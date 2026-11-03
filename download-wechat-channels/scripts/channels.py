#!/usr/bin/env python3
"""Resolve and download WeChat Channels videos from public share links.

This module intentionally uses only the Python standard library so the same
skill can run on macOS and Windows without installing project dependencies.
"""

from __future__ import annotations

import argparse
import contextlib
import dataclasses
import datetime as dt
import errno
import ipaddress
import json
import math
import os
import platform
import re
import shutil
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Iterable, Mapping


VERSION = "2.1.1"
LEGACY_PUBLIC_RESOLVER = "https://sph.litao.workers.dev/api/fetch_video_profile"
OFFICIAL_FEED_URL = "https://channels.weixin.qq.com/finder-preview/api/feed/get_feed_info"
OFFICIAL_PROFILE_MARKER = "download-wechat-channels-v2"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)
SHARE_URL_RE = re.compile(
    r"https?://(?:www\.)?weixin\.qq\.com/sph/[A-Za-z0-9_-]+(?:\?[^\s<>\"']*)?",
    re.IGNORECASE,
)
TRAILING_PUNCTUATION = "，。！？；：、）》】」』,.;:!?)]}"
DEFAULT_MEDIA_HOST_SUFFIXES = (
    "qq.com",
    "weixin.qq.com",
    "gtimg.com",
    "qpic.cn",
)
WINDOWS_RESERVED_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


class ChannelsError(Exception):
    def __init__(
        self,
        code: str,
        message: str,
        *,
        hint: str = "",
        details: Mapping[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint
        self.details = dict(details or {})

    def as_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "ok": False,
            "code": self.code,
            "message": self.message,
        }
        if self.hint:
            result["hint"] = self.hint
        if self.details:
            result["details"] = self.details
        return result


@dataclasses.dataclass(frozen=True)
class VideoCandidate:
    kind: str
    url: str
    width: int = 0
    height: int = 0
    size: int = 0


@dataclasses.dataclass(frozen=True)
class ResolvedVideo:
    share_url: str
    title: str
    author: str
    created_at: str
    cover_url: str
    candidates: tuple[VideoCandidate, ...]
    raw_response: Mapping[str, Any]


def _is_windows() -> bool:
    return platform.system().lower() == "windows"


def app_data_dir() -> Path:
    override = os.environ.get("DOWNLOAD_WECHAT_CHANNELS_HOME")
    if override:
        return Path(override).expanduser().resolve()
    if _is_windows():
        root = os.environ.get("APPDATA")
        if root:
            return Path(root) / "download-wechat-channels"
    if platform.system() == "Darwin":
        return Path.home() / "Library" / "Application Support" / "download-wechat-channels"
    root = os.environ.get("XDG_CONFIG_HOME")
    if root:
        return Path(root).expanduser() / "download-wechat-channels"
    return Path.home() / ".config" / "download-wechat-channels"


def default_output_dir() -> Path:
    override = os.environ.get("DOWNLOAD_WECHAT_CHANNELS_OUTPUT")
    if override:
        return Path(override).expanduser().resolve()
    downloads = _windows_downloads_dir() if _is_windows() else Path.home() / "Downloads"
    return downloads / "视频号"


def _windows_downloads_dir() -> Path:
    """Read the Windows Downloads known folder, with a portable fallback."""
    try:
        import ctypes
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", wintypes.DWORD),
                ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        folder_id = GUID(
            0x374DE290,
            0x123F,
            0x4565,
            (ctypes.c_ubyte * 8)(0x91, 0x64, 0x39, 0xC4, 0x92, 0x5E, 0x46, 0x7B),
        )
        path_pointer = ctypes.c_void_p()
        shell32 = ctypes.windll.shell32
        shell32.SHGetKnownFolderPath.argtypes = [
            ctypes.POINTER(GUID),
            wintypes.DWORD,
            wintypes.HANDLE,
            ctypes.POINTER(ctypes.c_void_p),
        ]
        shell32.SHGetKnownFolderPath.restype = ctypes.HRESULT
        ole32 = ctypes.windll.ole32
        ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
        ole32.CoTaskMemFree.restype = None
        result = shell32.SHGetKnownFolderPath(
            ctypes.byref(folder_id), 0, None, ctypes.byref(path_pointer)
        )
        if result == 0 and path_pointer.value:
            try:
                return Path(ctypes.wstring_at(path_pointer.value))
            finally:
                ole32.CoTaskMemFree(path_pointer)
    except (AttributeError, OSError, TypeError, ValueError):
        pass
    return Path.home() / "Downloads"


def config_path() -> Path:
    return app_data_dir() / "config.json"


def load_config() -> dict[str, Any]:
    path = config_path()
    if not path.exists():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ChannelsError(
            "CONFIG_INVALID",
            f"配置文件无法读取：{path}",
            hint="修复 JSON，或运行 config reset 后重新配置。",
            details={"error": str(exc)},
        ) from exc
    if not isinstance(value, dict):
        raise ChannelsError("CONFIG_INVALID", "配置文件根节点必须是 JSON 对象。")
    return value


def save_config(value: Mapping[str, Any]) -> Path:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix="config-", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(dict(value), handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(temp_name, path)
    except Exception:
        with contextlib.suppress(OSError):
            os.unlink(temp_name)
        raise
    return path


def resolver_url(override: str | None = None) -> str | None:
    url = (
        override
        or os.environ.get("DOWNLOAD_WECHAT_CHANNELS_RESOLVER")
        or load_config().get("resolver")
    )
    return validate_resolver_url(str(url)) if url else None


def validate_resolver_url(value: str) -> str:
    parsed = urllib.parse.urlsplit(value.strip())
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ChannelsError("INVALID_RESOLVER", "解析器必须是有效的 HTTP(S) URL。")
    if parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise ChannelsError("INVALID_RESOLVER", "解析器 URL 不能包含凭据、查询参数或片段。")
    try:
        parsed.port
    except ValueError as exc:
        raise ChannelsError("INVALID_RESOLVER", "解析器 URL 端口格式不正确。") from exc
    local_hosts = {"localhost", "127.0.0.1", "::1"}
    if parsed.scheme != "https" and parsed.hostname not in local_hosts:
        raise ChannelsError(
            "INSECURE_RESOLVER",
            "远程解析器必须使用 HTTPS。",
            hint="只有本机开发解析器允许使用 HTTP。",
        )
    return urllib.parse.urlunsplit(parsed)


def _url_origin(value: str) -> tuple[str, str, int | None]:
    parsed = urllib.parse.urlsplit(value)
    port = parsed.port
    if port is None:
        port = 443 if parsed.scheme.lower() == "https" else 80
    return parsed.scheme.lower(), (parsed.hostname or "").lower(), port


def resolver_token_for(resolver: str) -> str:
    token = os.environ.get("DOWNLOAD_WECHAT_CHANNELS_RESOLVER_TOKEN", "")
    if token and _url_origin(resolver) == _url_origin(LEGACY_PUBLIC_RESOLVER):
        raise ChannelsError(
            "RESOLVER_TOKEN_SCOPE_ERROR",
            "拒绝向已知公共解析器发送自定义 Bearer Token。",
            hint="仅在配置了你信任的自有解析器后设置 RESOLVER_TOKEN。",
        )
    return token


def extract_share_url(text: str) -> str:
    match = SHARE_URL_RE.search(text or "")
    if not match:
        raise ChannelsError(
            "INVALID_SHARE_LINK",
            "没有找到有效的视频号分享链接。",
            hint="请在微信视频号中选择“分享 → 复制链接”后重新提供。",
        )
    candidate = match.group(0).rstrip(TRAILING_PUNCTUATION)
    parsed = urllib.parse.urlsplit(candidate)
    host = (parsed.hostname or "").lower()
    if host not in {"weixin.qq.com", "www.weixin.qq.com"}:
        raise ChannelsError("INVALID_SHARE_LINK", "分享链接域名不受支持。")
    try:
        port = parsed.port
    except ValueError as exc:
        raise ChannelsError("INVALID_SHARE_LINK", "分享链接端口格式不正确。") from exc
    if port not in {None, 80, 443}:
        raise ChannelsError("INVALID_SHARE_LINK", "分享链接端口不受支持。")
    path_parts = [part for part in parsed.path.split("/") if part]
    if len(path_parts) != 2 or path_parts[0].lower() != "sph":
        raise ChannelsError("INVALID_SHARE_LINK", "分享链接路径不符合视频号格式。")
    return urllib.parse.urlunsplit(("https", "weixin.qq.com", parsed.path, parsed.query, ""))


def _read_input(positional: str | None, input_file: str | None) -> str:
    if positional:
        return positional
    if input_file:
        try:
            return Path(input_file).read_text(encoding="utf-8")
        except OSError as exc:
            raise ChannelsError("INPUT_READ_ERROR", f"无法读取输入文件：{input_file}") from exc
    if not sys.stdin.isatty():
        return sys.stdin.read()
    raise ChannelsError(
        "MISSING_INPUT",
        "请提供视频号分享文字或链接。",
        hint="可将链接作为参数传入，也可通过标准输入或 --input-file 提供。",
    )


def _json_request(
    url: str,
    payload: Mapping[str, Any],
    timeout: float,
    resolver_token: str = "",
) -> Mapping[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json; charset=utf-8",
        "User-Agent": USER_AGENT,
    }
    if resolver_token:
        headers["Authorization"] = f"Bearer {resolver_token}"
    request = urllib.request.Request(
        url,
        data=body,
        method="POST",
        headers=headers,
    )
    try:
        opener = urllib.request.build_opener(SafeResolverRedirectHandler(url))
        with opener.open(request, timeout=timeout) as response:
            raw = response.read(4 * 1024 * 1024)
    except urllib.error.HTTPError as exc:
        raw = exc.read(16 * 1024)
        message = ""
        with contextlib.suppress(Exception):
            parsed_error = json.loads(raw.decode("utf-8", errors="replace"))
            message = str(parsed_error.get("error") or parsed_error.get("message") or "")
        if exc.code in {404, 410, 422}:
            code = "VIDEO_UNAVAILABLE"
            hint = "链接可能已失效、内容已删除或不允许公开预览。"
        else:
            code = "RESOLVER_HTTP_ERROR"
            hint = "稍后重试，或配置自有解析器。"
        raise ChannelsError(
            code,
            f"解析服务返回 HTTP {exc.code}" + (f"：{message}" if message else "。"),
            hint=hint,
            details={"status": exc.code},
        ) from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise ChannelsError(
            "RESOLVER_UNAVAILABLE",
            "无法连接视频号解析服务。",
            hint="检查网络后重试，或配置其他解析器。",
            details={"error": str(exc)},
        ) from exc
    try:
        value = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ChannelsError(
            "RESOLVER_INVALID_RESPONSE",
            "解析服务返回了无效的 JSON。",
            hint="解析器接口可能已经变更。",
        ) from exc
    if not isinstance(value, dict):
        raise ChannelsError("RESOLVER_INVALID_RESPONSE", "解析服务返回格式不正确。")
    if value.get("error"):
        raise ChannelsError(
            "VIDEO_UNAVAILABLE",
            f"解析失败：{value['error']}",
            hint="确认链接仍能在微信中打开，或稍后重试。",
        )
    return value


class SafeResolverRedirectHandler(urllib.request.HTTPRedirectHandler):
    """Permit redirects only inside the configured resolver origin."""

    def __init__(self, resolver: str) -> None:
        super().__init__()
        self._origin = _url_origin(resolver)

    def redirect_request(
        self,
        req: urllib.request.Request,
        fp: Any,
        code: int,
        msg: str,
        headers: Any,
        newurl: str,
    ) -> urllib.request.Request | None:
        parsed = urllib.parse.urlsplit(newurl)
        try:
            unsafe = (
                parsed.scheme not in {"http", "https"}
                or not parsed.hostname
                or parsed.username
                or parsed.password
                or parsed.fragment
                or _url_origin(newurl) != self._origin
            )
        except ValueError:
            unsafe = True
        if unsafe:
            raise ChannelsError(
                "UNSAFE_RESOLVER_REDIRECT",
                "解析器返回了不安全的跨域或降级重定向。",
                hint="请修复解析器，使重定向保持在同一来源。",
            )
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _to_int(value: Any) -> int:
    try:
        return int(value or 0)
    except (TypeError, ValueError):
        return 0


def _candidate(kind: str, value: Any) -> VideoCandidate | None:
    if isinstance(value, str):
        url = value
        info: Mapping[str, Any] = {}
    elif isinstance(value, Mapping):
        url = str(value.get("videoUrl") or value.get("url") or "")
        info = value
    else:
        return None
    if not url:
        return None
    validate_media_url(url)
    return VideoCandidate(
        kind=kind,
        url=url,
        width=_to_int(info.get("width") or info.get("videoWidth")),
        height=_to_int(info.get("height") or info.get("videoHeight")),
        size=_to_int(info.get("fileSize") or info.get("size") or info.get("videoSize")),
    )


def _created_at(value: Any) -> str:
    timestamp = _to_int(value)
    if timestamp <= 0:
        return ""
    with contextlib.suppress(ValueError, OSError, OverflowError):
        return dt.datetime.fromtimestamp(timestamp).astimezone().isoformat(timespec="seconds")
    return ""


def parse_resolver_response(share_url: str, value: Mapping[str, Any]) -> ResolvedVideo:
    data = value.get("data")
    if not isinstance(data, Mapping):
        raise ChannelsError(
            "VIDEO_UNAVAILABLE",
            "解析结果中没有视频数据。",
            hint="链接可能过期、内容已删除或解析器接口已经变化。",
        )
    feed = data.get("feedInfo")
    if not isinstance(feed, Mapping):
        raise ChannelsError("VIDEO_UNAVAILABLE", "解析结果中没有 feedInfo。")
    author_info = data.get("authorInfo")
    author = str(author_info.get("nickname") or "") if isinstance(author_info, Mapping) else ""
    candidates = tuple(
        candidate
        for candidate in (
            _candidate("h264", feed.get("h264VideoInfo")),
            _candidate("h265", feed.get("h265VideoInfo")),
            _candidate("default", feed.get("videoUrl")),
        )
        if candidate is not None
    )
    if not candidates:
        raise ChannelsError(
            "VIDEO_UNAVAILABLE",
            "解析成功，但没有找到可下载的视频地址。",
            hint="当前内容可能不是普通视频，或不允许公开预览。",
        )
    title = str(feed.get("description") or "").strip()
    return ResolvedVideo(
        share_url=share_url,
        title=title,
        author=author,
        created_at=_created_at(feed.get("createtime") or feed.get("createTime")),
        cover_url=str(feed.get("coverUrl") or ""),
        candidates=candidates,
        raw_response=value,
    )


def resolve_video(
    text: str,
    resolver: str,
    timeout: float,
    resolver_token: str = "",
) -> ResolvedVideo:
    share_url = extract_share_url(text)
    value = _json_request(resolver, {"url": share_url}, timeout, resolver_token)
    return parse_resolver_response(share_url, value)


def _read_official_profile(path_value: str, expected_share_url: str) -> Mapping[str, Any]:
    supplied_path = Path(path_value).expanduser()
    if supplied_path.is_symlink():
        raise ChannelsError("PROFILE_INVALID", "官方解析结果不能是符号链接。")
    path = supplied_path.resolve()
    try:
        if path.stat().st_size > 1024 * 1024:
            raise ChannelsError("PROFILE_TOO_LARGE", "官方解析结果文件超过 1 MB。")
        value = json.loads(path.read_text(encoding="utf-8"))
    except ChannelsError:
        raise
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ChannelsError(
            "PROFILE_INVALID",
            f"无法读取官方解析结果：{path}",
            hint="重新执行浏览器解析流程，不要手工编辑临时文件。",
            details={"error": str(exc)},
        ) from exc
    if not isinstance(value, Mapping):
        raise ChannelsError("PROFILE_INVALID", "官方解析结果必须是 JSON 对象。")
    if value.get("created_by") != OFFICIAL_PROFILE_MARKER:
        raise ChannelsError("PROFILE_INVALID", "官方解析结果缺少可信的本地流程标记。")
    if value.get("share_url") != expected_share_url:
        raise ChannelsError(
            "PROFILE_LINK_MISMATCH",
            "官方解析结果与当前分享链接不匹配。",
            hint="重新为当前链接执行 prepare 和浏览器解析。",
        )
    response = value.get("yuanbao_response")
    if not isinstance(response, Mapping):
        raise ChannelsError("PROFILE_INVALID", "官方解析结果中没有 yuanbao_response。")
    return response


def _playable_parameters(value: Mapping[str, Any]) -> tuple[str, str]:
    data = value.get("data")
    if not isinstance(data, Mapping):
        raise ChannelsError(
            "YUANBAO_PROFILE_INVALID",
            "腾讯元宝没有返回视频号解析数据。",
            hint="确认已登录腾讯元宝后重新解析。",
        )
    playable_url = str(data.get("playable_url") or "")
    parsed = urllib.parse.urlsplit(playable_url)
    if (
        parsed.scheme != "https"
        or (parsed.hostname or "").lower() != "channels.weixin.qq.com"
        or not parsed.path.startswith("/finder-preview/pages/feed")
        or parsed.username
        or parsed.password
    ):
        raise ChannelsError("YUANBAO_PROFILE_INVALID", "腾讯元宝返回的播放页地址不受信任。")
    query = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
    token = (query.get("token") or [""])[0]
    export_id = (query.get("eid") or [""])[0]
    if (
        not token
        or not export_id
        or len(token) > 8192
        or len(export_id) > 8192
        or any(ord(char) < 32 for char in token + export_id)
    ):
        raise ChannelsError(
            "YUANBAO_PROFILE_INVALID",
            "腾讯元宝解析结果缺少 token 或 eid。",
            hint="解析接口可能已经变化，请更新 Skill。",
        )
    return token, export_id


def _official_feed_request(
    token: str,
    export_id: str,
    timeout: float,
) -> Mapping[str, Any]:
    rid = f"{int(dt.datetime.now().timestamp()):x}-{os.urandom(4).hex()}"
    page_url = "https://channels.weixin.qq.com/finder-preview/pages/feed"
    request_url = OFFICIAL_FEED_URL + "?" + urllib.parse.urlencode(
        {"_rid": rid, "_pageUrl": page_url}
    )
    referer = page_url + "?" + urllib.parse.urlencode(
        {
            "entry_card_type": "48",
            "comment_scene": "39",
            "appid": "0",
            "token": token,
            "entry_scene": "0",
            "eid": export_id,
        }
    )
    body = json.dumps(
        {"baseReq": {"generalToken": token}, "exportId": export_id},
        ensure_ascii=False,
    ).encode("utf-8")
    request = urllib.request.Request(
        request_url,
        data=body,
        method="POST",
        headers={
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/json; charset=utf-8",
            "Origin": "https://channels.weixin.qq.com",
            "Referer": referer,
            "User-Agent": USER_AGENT,
        },
    )
    try:
        opener = urllib.request.build_opener(SafeResolverRedirectHandler(OFFICIAL_FEED_URL))
        with opener.open(request, timeout=timeout) as response:
            raw = response.read(4 * 1024 * 1024)
    except urllib.error.HTTPError as exc:
        raise ChannelsError(
            "OFFICIAL_FEED_HTTP_ERROR",
            f"腾讯视频号官方接口返回 HTTP {exc.code}。",
            hint="登录令牌可能已过期，请重新执行浏览器解析。",
            details={"status": exc.code},
        ) from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise ChannelsError(
            "OFFICIAL_FEED_UNAVAILABLE",
            "无法连接腾讯视频号官方接口。",
            hint="检查网络后重新执行浏览器解析。",
            details={"error": str(exc)},
        ) from exc
    try:
        value = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ChannelsError("OFFICIAL_FEED_INVALID", "腾讯视频号官方接口返回了无效 JSON。") from exc
    if not isinstance(value, Mapping):
        raise ChannelsError("OFFICIAL_FEED_INVALID", "腾讯视频号官方接口返回格式不正确。")
    if _to_int(value.get("errCode")) != 0:
        raise ChannelsError(
            "OFFICIAL_FEED_ERROR",
            f"腾讯视频号官方接口解析失败：{value.get('errMsg') or value.get('errCode')}",
            hint="重新登录腾讯元宝并再次解析该链接。",
        )
    return value


def resolve_video_from_official_profile(
    text: str,
    profile_file: str,
    timeout: float,
) -> ResolvedVideo:
    share_url = extract_share_url(text)
    profile = _read_official_profile(profile_file, share_url)
    token, export_id = _playable_parameters(profile)
    value = _official_feed_request(token, export_id, timeout)
    return parse_resolver_response(share_url, value)


def _consume_official_profile_file(path_value: str | None) -> None:
    if not path_value:
        return
    supplied_path = Path(path_value).expanduser()
    if supplied_path.is_symlink():
        return
    path = supplied_path.resolve()
    temp_root = Path(tempfile.gettempdir()).resolve()
    if (
        path.parent != temp_root
        or not path.name.startswith("download-wechat-channels-")
        or path.suffix != ".json"
    ):
        return
    try:
        if path.stat().st_size > 1024 * 1024:
            return
        value = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(value, Mapping) or value.get("created_by") != OFFICIAL_PROFILE_MARKER:
            return
        path.unlink()
    except (OSError, UnicodeDecodeError, json.JSONDecodeError):
        return


def validate_media_url(value: str) -> str:
    parsed = urllib.parse.urlsplit(value)
    if parsed.scheme != "https" or not parsed.hostname:
        raise ChannelsError("INVALID_MEDIA_URL", "媒体地址必须是有效的 HTTPS URL。")
    if parsed.username or parsed.password:
        raise ChannelsError("INVALID_MEDIA_URL", "媒体地址不能包含 URL 凭据。")
    host = parsed.hostname.lower()
    if host in {"localhost", "localhost.localdomain"}:
        raise ChannelsError("INVALID_MEDIA_URL", "解析器返回了本机媒体地址。")
    with contextlib.suppress(ValueError):
        address = ipaddress.ip_address(host)
        if not address.is_global:
            raise ChannelsError("INVALID_MEDIA_URL", "解析器返回了非公开网络地址。")
    configured = os.environ.get("DOWNLOAD_WECHAT_CHANNELS_MEDIA_HOSTS", "")
    configured_suffixes = tuple(
        item.strip().lower().lstrip(".")
        for item in configured.split(",")
        if item.strip()
    )
    suffixes = DEFAULT_MEDIA_HOST_SUFFIXES + tuple(
        suffix for suffix in configured_suffixes if suffix not in DEFAULT_MEDIA_HOST_SUFFIXES
    )
    if not any(host == suffix or host.endswith(f".{suffix}") for suffix in suffixes):
        raise ChannelsError(
            "UNTRUSTED_MEDIA_HOST",
            f"媒体域名不在允许列表中：{host}",
            hint="确认解析器可信；如腾讯更换 CDN，可通过环境变量扩展允许域名。",
        )
    return urllib.parse.urlunsplit(parsed)


def original_video_url(url: str) -> str | None:
    parsed = urllib.parse.urlsplit(url)
    query = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
    file_key = (query.get("encfilekey") or [""])[0]
    token = (query.get("token") or [""])[0]
    if not file_key or not token:
        return None
    raw_query = urllib.parse.urlencode({"encfilekey": file_key, "token": token})
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, raw_query, ""))


def choose_candidate(video: ResolvedVideo, quality: str) -> VideoCandidate:
    by_kind: dict[str, VideoCandidate] = {}
    for candidate in video.candidates:
        by_kind.setdefault(candidate.kind, candidate)
    compatible = by_kind.get("h264") or by_kind.get("default")
    if quality == "auto":
        sources = [candidate for candidate in (compatible, by_kind.get("h265")) if candidate]
        if not sources:
            raise ChannelsError("VIDEO_UNAVAILABLE", "没有可下载的视频地址。")
        for source in sources:
            original_url = original_video_url(source.url)
            if original_url is not None:
                return dataclasses.replace(source, kind="original", url=original_url)
        return compatible or by_kind["h265"]
    if quality == "compatible":
        if compatible is None:
            raise ChannelsError(
                "COMPATIBLE_UNAVAILABLE",
                "该视频没有 H.264 兼容地址。",
                hint="可尝试 --quality h265。",
            )
        return compatible
    if quality == "h265":
        if "h265" not in by_kind:
            raise ChannelsError(
                "H265_UNAVAILABLE",
                "该视频没有 H.265 地址。",
                hint="请改用 --quality original 或 --quality compatible。",
            )
        return by_kind["h265"]
    if quality == "original":
        sources = [candidate for candidate in (compatible, by_kind.get("h265")) if candidate]
        if not sources:
            raise ChannelsError("VIDEO_UNAVAILABLE", "没有可下载的视频地址。")
        for source in sources:
            original_url = original_video_url(source.url)
            if original_url is not None:
                return dataclasses.replace(source, kind="original", url=original_url)
        raise ChannelsError(
            "ORIGINAL_UNAVAILABLE",
            "解析结果无法确认原始视频地址。",
            hint="请改用 --quality compatible 或 --quality h265。",
        )
    raise ChannelsError("INVALID_QUALITY", f"不支持的画质：{quality}")


def sanitize_filename(value: str, fallback: str = "视频号视频", max_length: int = 160) -> str:
    name = re.sub(r"[\x00-\x1f\x7f<>:\"/\\|?*]+", " ", value or "")
    name = re.sub(r"\s+", " ", name).strip(" .")
    if not name:
        name = fallback
    if name.split(".", 1)[0].upper() in WINDOWS_RESERVED_NAMES:
        name = f"_{name}"
    name = name[:max_length].rstrip(" .")
    while name and (
        len(name.encode("utf-8")) > 180
        or len(name.encode("utf-16-le")) // 2 > 180
    ):
        name = name[:-1].rstrip(" .")
    return name or fallback


def suggested_filename(video: ResolvedVideo) -> str:
    base = video.title or ""
    if not base and video.author:
        base = f"{video.author}的视频"
    if not base and video.created_at:
        base = f"视频号视频_{video.created_at[:10].replace('-', '')}"
    return sanitize_filename(base) + ".mp4"


def unique_destination(directory: Path, filename: str) -> Path:
    candidate = directory / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    for index in range(2, 10_000):
        numbered = directory / f"{stem} ({index}){suffix}"
        if not numbered.exists():
            return numbered
    raise ChannelsError("TOO_MANY_DUPLICATES", "同名文件过多，无法生成安全文件名。")


def _looks_like_error_page(prefix: bytes) -> bool:
    sample = prefix.lstrip(b"\xef\xbb\xbf\x00\t\r\n ").lower()
    return sample.startswith((b"<", b"{", b"["))


def _looks_like_mp4(prefix: bytes) -> bool:
    return len(prefix) >= 12 and prefix[4:8] == b"ftyp"


class SafeMediaRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(
        self,
        req: urllib.request.Request,
        fp: Any,
        code: int,
        msg: str,
        headers: Any,
        newurl: str,
    ) -> urllib.request.Request | None:
        validate_media_url(newurl)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _open_media(request: urllib.request.Request, timeout: float) -> Any:
    opener = urllib.request.build_opener(SafeMediaRedirectHandler())
    response = opener.open(request, timeout=timeout)
    validate_media_url(response.geturl())
    return response


def _commit_without_overwrite(temp_path: Path, directory: Path, filename: str) -> Path:
    unsupported_link_errors = {errno.EPERM, errno.EACCES, errno.EINVAL, errno.EXDEV}
    if hasattr(errno, "ENOSYS"):
        unsupported_link_errors.add(errno.ENOSYS)
    for error_name in ("ENOTSUP", "EOPNOTSUPP"):
        error_number = getattr(errno, error_name, None)
        if error_number is not None:
            unsupported_link_errors.add(error_number)
    for _ in range(10_000):
        destination = unique_destination(directory, filename)
        try:
            os.link(temp_path, destination)
            temp_path.unlink()
            return destination
        except FileExistsError:
            continue
        except OSError as exc:
            if exc.errno not in unsupported_link_errors:
                raise
            created = False
            try:
                with destination.open("xb") as target, temp_path.open("rb") as source:
                    created = True
                    shutil.copyfileobj(source, target, length=1024 * 1024)
                temp_path.unlink()
                return destination
            except FileExistsError:
                continue
            except Exception:
                if created:
                    with contextlib.suppress(OSError):
                        destination.unlink()
                raise
    raise ChannelsError("TOO_MANY_DUPLICATES", "同名文件过多，无法安全保存。")


def download_media(
    url: str,
    output_dir: Path,
    filename: str,
    *,
    timeout: float,
    max_bytes: int,
) -> tuple[Path, int, str]:
    validate_media_url(url)
    safe_filename = sanitize_filename(Path(filename).stem) + ".mp4"
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ChannelsError(
            "OUTPUT_DIRECTORY_ERROR",
            f"无法创建输出目录：{output_dir}",
            hint="使用 --output-dir 指定可写目录。",
            details={"error": str(exc)},
        ) from exc
    if not output_dir.is_dir():
        raise ChannelsError("OUTPUT_NOT_DIRECTORY", f"输出路径不是目录：{output_dir}")
    request = urllib.request.Request(
        url,
        method="GET",
        headers={
            "Accept": "video/*,application/octet-stream;q=0.9,*/*;q=0.1",
            "Referer": "https://channels.weixin.qq.com/",
            "User-Agent": USER_AGENT,
        },
    )
    temp_path: Path | None = None
    try:
        try:
            response = _open_media(request, timeout)
        except urllib.error.HTTPError as exc:
            raise ChannelsError(
                "DOWNLOAD_HTTP_ERROR",
                f"下载地址返回 HTTP {exc.code}。",
                hint="媒体地址可能已经过期，请重新解析链接后再试。",
                details={"status": exc.code},
            ) from exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            raise ChannelsError(
                "DOWNLOAD_UNAVAILABLE",
                "无法连接视频下载地址。",
                hint="检查网络后重新解析链接并重试。",
                details={"error": str(exc)},
            ) from exc
        with response:
            length = _to_int(response.headers.get("Content-Length"))
            if length and length > max_bytes:
                raise ChannelsError(
                    "FILE_TOO_LARGE",
                    f"视频大小超过限制：{length} 字节。",
                    hint="确认链接无误后增大 --max-size-gb。",
                )
            content_type = response.headers.get_content_type()
            allowed_content_types = {
                "application/octet-stream",
                "application/mp4",
                "binary/octet-stream",
            }
            if not (content_type.startswith("video/") or content_type in allowed_content_types):
                raise ChannelsError(
                    "INVALID_CONTENT_TYPE",
                    f"下载地址返回了非视频类型：{content_type}",
                    hint="媒体地址可能已经过期，请重新解析。",
                )
            fd, temp_name = tempfile.mkstemp(
                prefix=".channels-",
                suffix=".part",
                dir=output_dir,
            )
            temp_path = Path(temp_name)
            total = 0
            prefix = b""
            with os.fdopen(fd, "wb") as handle:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    if len(prefix) < 64:
                        prefix += chunk[: 64 - len(prefix)]
                    total += len(chunk)
                    if total > max_bytes:
                        raise ChannelsError(
                            "FILE_TOO_LARGE",
                            "实际下载大小超过限制，已停止。",
                            hint="确认链接无误后增大 --max-size-gb。",
                        )
                    handle.write(chunk)
            if length and total != length:
                raise ChannelsError(
                    "INCOMPLETE_DOWNLOAD",
                    "视频下载不完整，未保存残缺文件。",
                    hint="媒体地址可能已过期或网络中断，请重新解析后重试。",
                    details={"expected_bytes": length, "actual_bytes": total},
                )
            if total == 0:
                raise ChannelsError("EMPTY_DOWNLOAD", "下载结果为空文件。")
            if _looks_like_error_page(prefix):
                raise ChannelsError(
                    "INVALID_DOWNLOAD",
                    "下载结果是错误页面，不是视频。",
                    hint="媒体地址可能已经过期，请重新解析。",
                )
            if not _looks_like_mp4(prefix):
                raise ChannelsError(
                    "INVALID_VIDEO_FILE",
                    "下载结果不是有效的 MP4 文件。",
                    hint="媒体地址可能已经过期，或解析器返回了不支持的格式。",
                )
            destination = _commit_without_overwrite(temp_path, output_dir, safe_filename)
            temp_path = None
            return destination.resolve(), total, content_type
    except ChannelsError:
        raise
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise ChannelsError(
            "DOWNLOAD_IO_ERROR",
            "下载或保存视频时发生 I/O 错误。",
            hint="检查网络、磁盘空间和输出目录权限后重试。",
            details={"error": str(exc)},
        ) from exc
    finally:
        if temp_path is not None:
            with contextlib.suppress(OSError):
                temp_path.unlink()


def resolved_public_dict(video: ResolvedVideo, selected: VideoCandidate | None = None) -> dict[str, Any]:
    kinds = {candidate.kind for candidate in video.candidates}
    available: list[str] = []
    if "h264" in kinds or "default" in kinds:
        available.append("compatible")
    if "h265" in kinds:
        available.append("h265")
    if any(original_video_url(candidate.url) for candidate in video.candidates):
        available.append("original")
    result: dict[str, Any] = {
        "ok": True,
        "share_url": video.share_url,
        "title": video.title,
        "author": video.author,
        "created_at": video.created_at,
        "available_qualities": available,
    }
    if selected:
        result["selected_quality"] = selected.kind
        result["width"] = selected.width
        result["height"] = selected.height
        result["declared_size"] = selected.size
    return result


def _human_size(value: int) -> str:
    size = float(value)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024 or unit == "TB":
            return f"{size:.2f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024
    return f"{value} B"


def print_result(value: Mapping[str, Any], as_json: bool) -> None:
    if as_json:
        print(json.dumps(dict(value), ensure_ascii=False, indent=2))
        return
    if not value.get("ok"):
        print(f"错误 [{value.get('code', 'UNKNOWN')}]：{value.get('message', '')}", file=sys.stderr)
        if value.get("hint"):
            print(f"建议：{value['hint']}", file=sys.stderr)
        return
    if value.get("path"):
        print("下载完成")
        print(f"标题：{value.get('title') or '未提供'}")
        print(f"作者：{value.get('author') or '未提供'}")
        print(f"画质：{value.get('selected_quality')}")
        print(f"大小：{value.get('size_human')}")
        print(f"位置：{value.get('path')}")
        return
    if value.get("tests"):
        print(value.get("message") or "内置测试通过。")
        return
    if value.get("profile_file"):
        print(value.get("message") or "已准备官方解析临时文件。")
        print(f"链接：{value.get('share_url')}")
        print(f"临时文件：{value.get('profile_file')}")
        return
    if "platform" in value and "output_dir" in value:
        print(value.get("message") or "本地环境检查完成。")
        print(f"系统：{value.get('platform')} {value.get('architecture')}")
        print(f"Python：{value.get('python')}")
        print(f"模式：{value.get('mode')}")
        print(f"输出目录：{value.get('output_dir')}")
        return
    if value.get("config_path") and "mode" in value:
        print(f"模式：{value['mode']}")
        if value.get("resolver"):
            print(f"解析器：{value['resolver']}")
        print(f"配置：{value['config_path']}")
        return
    if value.get("config_path") and value.get("reset"):
        print(f"配置已恢复默认值：{value['config_path']}")
        return
    if value.get("message") and "title" not in value:
        print(value["message"])
        return
    print("解析成功")
    print(f"标题：{value.get('title') or '未提供'}")
    print(f"作者：{value.get('author') or '未提供'}")
    print(f"可用格式：{', '.join(value.get('available_qualities', []))}")


def command_doctor(args: argparse.Namespace) -> dict[str, Any]:
    system = platform.system()
    supported = system in {"Darwin", "Windows"}
    output = Path(args.output_dir).expanduser().resolve() if args.output_dir else default_output_dir()
    output_is_file = output.exists() and not output.is_dir()
    parent = output if output.exists() and output.is_dir() else output.parent
    writable_ancestor = parent
    while not writable_ancestor.exists() and writable_ancestor != writable_ancestor.parent:
        writable_ancestor = writable_ancestor.parent
    python_supported = sys.version_info >= (3, 9)
    parent_exists = parent.exists()
    ancestor_exists = writable_ancestor.exists()
    parent_writable = False
    write_error = ""
    if ancestor_exists and writable_ancestor.is_dir() and not output_is_file:
        probe_path = ""
        try:
            fd, probe_path = tempfile.mkstemp(prefix=".channels-doctor-", dir=writable_ancestor)
            os.close(fd)
            parent_writable = True
        except OSError as exc:
            write_error = str(exc)
        finally:
            if probe_path:
                with contextlib.suppress(OSError):
                    os.unlink(probe_path)
    ok = supported and python_supported and ancestor_exists and parent_writable and not output_is_file
    if not supported:
        code = "UNSUPPORTED_PLATFORM"
        message = "当前仅正式支持 macOS 和 Windows。"
    elif not python_supported:
        code = "PYTHON_TOO_OLD"
        message = "需要 Python 3.9 或更高版本。"
    elif output_is_file:
        code = "OUTPUT_NOT_DIRECTORY"
        message = "输出路径是普通文件，不是目录。"
    elif not ancestor_exists or not parent_writable:
        code = "OUTPUT_NOT_WRITABLE"
        message = "输出目录最近的现有父目录不可写。"
    else:
        code = "OK"
        message = "本地环境检查通过；官方浏览器登录状态未检查。"
    resolver = resolver_url(args.resolver)
    return {
        "ok": ok,
        "code": code,
        "platform": system,
        "architecture": platform.machine(),
        "python": platform.python_version(),
        "python_supported": python_supported,
        "mode": "custom-resolver" if resolver else "official-browser",
        "resolver": resolver,
        "resolver_network_checked": False,
        "output_dir": str(output),
        "output_parent_exists": parent_exists,
        "output_writable_ancestor": str(writable_ancestor),
        "output_parent_writable": parent_writable,
        "output_write_error": write_error,
        "config_path": str(config_path()),
        "message": message,
    }


def command_prepare(args: argparse.Namespace) -> dict[str, Any]:
    text = _read_input(args.text, args.input_file)
    share_url = extract_share_url(text)
    fd, temp_name = tempfile.mkstemp(
        prefix="download-wechat-channels-",
        suffix=".json",
    )
    value = {
        "created_by": OFFICIAL_PROFILE_MARKER,
        "share_url": share_url,
        "yuanbao_response": None,
    }
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
    except Exception:
        with contextlib.suppress(OSError):
            os.unlink(temp_name)
        raise
    return {
        "ok": True,
        "mode": "official-browser",
        "share_url": share_url,
        "profile_file": str(Path(temp_name).resolve()),
        "browser_url": "https://yuanbao.tencent.com/",
        "message": "已创建安全的官方解析临时文件。",
    }


def _resolve_from_args(args: argparse.Namespace, text: str) -> tuple[ResolvedVideo, str]:
    resolver = resolver_url(args.resolver)
    official_profile = getattr(args, "official_profile", None)
    if official_profile and resolver:
        raise ChannelsError(
            "CONFLICTING_RESOLUTION_MODE",
            "--official-profile 与自定义解析器不能同时使用。",
        )
    if official_profile:
        return (
            resolve_video_from_official_profile(text, official_profile, args.timeout),
            "official-browser",
        )
    if resolver:
        if _url_origin(resolver) == _url_origin(LEGACY_PUBLIC_RESOLVER):
            raise ChannelsError(
                "DEPRECATED_PUBLIC_RESOLVER",
                "已停用失效的默认公共解析器。",
                hint="运行 config reset，然后按 Skill 的腾讯元宝浏览器流程解析。",
            )
        return (
            resolve_video(text, resolver, args.timeout, resolver_token_for(resolver)),
            "custom-resolver",
        )
    raise ChannelsError(
        "OFFICIAL_PROFILE_REQUIRED",
        "下载前需要腾讯元宝官方登录态生成解析结果。",
        hint="先运行 prepare，再由 Skill 在腾讯元宝浏览器页面完成官方解析。",
    )


def _resolve_and_select(
    args: argparse.Namespace,
    text: str,
    *,
    consume_on_success: bool = True,
) -> tuple[ResolvedVideo, VideoCandidate, str]:
    profile_file = getattr(args, "official_profile", None)
    try:
        video, mode = _resolve_from_args(args, text)
        selected = choose_candidate(video, args.quality)
    except ChannelsError as exc:
        if exc.code not in {
            "ORIGINAL_UNAVAILABLE",
            "H265_UNAVAILABLE",
            "COMPATIBLE_UNAVAILABLE",
        }:
            _consume_official_profile_file(profile_file)
        raise
    if consume_on_success:
        _consume_official_profile_file(profile_file)
    return video, selected, mode


def command_resolve(args: argparse.Namespace) -> dict[str, Any]:
    text = _read_input(args.text, args.input_file)
    video, selected, mode = _resolve_and_select(args, text)
    result = resolved_public_dict(video, selected)
    result["mode"] = mode
    if resolver := resolver_url(args.resolver):
        result["resolver"] = resolver
    return result


def _fallback_candidate(
    video: ResolvedVideo, selected: VideoCandidate
) -> VideoCandidate | None:
    """Compatible candidate to retry with when an "original" URL is rejected."""
    if selected.kind != "original":
        return None
    by_kind = {candidate.kind: candidate for candidate in video.candidates}
    return by_kind.get("h264") or by_kind.get("default")


def command_download(args: argparse.Namespace) -> dict[str, Any]:
    text = _read_input(args.text, args.input_file)
    video, selected, mode = _resolve_and_select(args, text, consume_on_success=False)
    output = Path(args.output_dir).expanduser().resolve() if args.output_dir else default_output_dir()
    filename = args.filename or suggested_filename(video)
    try:
        path, size, content_type = download_media(
            selected.url,
            output,
            filename,
            timeout=args.timeout,
            max_bytes=int(args.max_size_gb * 1024**3),
        )
    except ChannelsError as exc:
        fallback = _fallback_candidate(video, selected)
        if exc.code != "DOWNLOAD_HTTP_ERROR" or fallback is None:
            raise
        # finder.video.qq.com 现在要求完整签名参数；裁剪后的“原画”地址会 400，
        # 这里退回到带完整参数的兼容地址，而不是让整个下载失败。
        path, size, content_type = download_media(
            fallback.url,
            output,
            filename,
            timeout=args.timeout,
            max_bytes=int(args.max_size_gb * 1024**3),
        )
        selected = fallback
    _consume_official_profile_file(getattr(args, "official_profile", None))
    result = resolved_public_dict(video, selected)
    result.update(
        {
            "path": str(path),
            "size": size,
            "size_human": _human_size(size),
            "content_type": content_type,
            "mode": mode,
        }
    )
    if resolver := resolver_url(args.resolver):
        result["resolver"] = resolver
    return result


def command_config(args: argparse.Namespace) -> dict[str, Any]:
    if args.config_command == "show":
        resolver = resolver_url()
        return {
            "ok": True,
            "config_path": str(config_path()),
            "mode": "custom-resolver" if resolver else "official-browser",
            "resolver": resolver,
            "is_default": resolver is None,
        }
    if args.config_command == "set-resolver":
        checked = validate_resolver_url(args.url)
        if _url_origin(checked) == _url_origin(LEGACY_PUBLIC_RESOLVER):
            raise ChannelsError(
                "DEPRECATED_PUBLIC_RESOLVER",
                "不能再把已失效的公共服务设为解析器。",
                hint="部署你自己的解析器，或使用默认的官方浏览器模式。",
            )
        value = load_config()
        value["resolver"] = checked
        path = save_config(value)
        return {"ok": True, "config_path": str(path), "resolver": value["resolver"]}
    if args.config_command == "reset":
        path = config_path()
        if path.exists():
            path.unlink()
        return {"ok": True, "config_path": str(path), "reset": True}
    raise ChannelsError("INVALID_COMMAND", "未知配置命令。")


def command_self_test(_: argparse.Namespace) -> dict[str, Any]:
    assert extract_share_url("看看这个 https://weixin.qq.com/sph/AbC_123。") == (
        "https://weixin.qq.com/sph/AbC_123"
    )
    assert sanitize_filename('A<B>:C/ D?*"') == "A B C D"
    assert sanitize_filename("CON") == "_CON"
    assert sanitize_filename("CON.txt") == "_CON.txt"
    assert len(sanitize_filename("😀" * 160).encode("utf-8")) <= 180
    raw = original_video_url(
        "https://example.com/video.mp4?foo=1&encfilekey=abc%2B123&token=t%2F1"
    )
    assert urllib.parse.parse_qs(urllib.parse.urlsplit(raw).query) == {
        "encfilekey": ["abc+123"],
        "token": ["t/1"],
    }
    fixture = {
        "data": {
            "feedInfo": {
                "description": "示例视频",
                "createtime": 1_700_000_000,
                "coverUrl": "https://example.com/cover.jpg",
                "h264VideoInfo": {"videoUrl": "https://finder.video.qq.com/video.mp4"},
            },
            "authorInfo": {"nickname": "作者"},
        }
    }
    parsed = parse_resolver_response("https://weixin.qq.com/sph/Test123", fixture)
    assert parsed.title == "示例视频" and parsed.author == "作者"
    assert choose_candidate(parsed, "compatible").kind == "h264"
    assert choose_candidate(parsed, "auto").kind == "h264"
    mixed = dataclasses.replace(
        parsed,
        candidates=(
            VideoCandidate("h264", "https://finder.video.qq.com/compatible.mp4"),
            VideoCandidate(
                "h265",
                "https://finder.video.qq.com/h265.mp4?encfilekey=key&token=token",
            ),
        ),
    )
    assert choose_candidate(mixed, "original").url.endswith("encfilekey=key&token=token")
    assert choose_candidate(mixed, "auto").kind == "original"
    token, export_id = _playable_parameters(
        {
            "data": {
                "playable_url": (
                    "https://channels.weixin.qq.com/finder-preview/pages/feed"
                    "?token=test-token&eid=export%2Ftest"
                )
            }
        }
    )
    assert token == "test-token" and export_id == "export/test"
    assert timeout_seconds("0.5") == 0.5
    return {"ok": True, "tests": 14, "message": "内置测试通过。"}


def bounded_positive_float(value: str, *, maximum: float, label: str) -> float:
    try:
        number = float(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("必须是数字") from exc
    if not math.isfinite(number) or number <= 0:
        raise argparse.ArgumentTypeError("必须是有限的正数")
    if number > maximum:
        raise argparse.ArgumentTypeError(f"{label}不能超过 {maximum:g}")
    return number


def timeout_seconds(value: str) -> float:
    return bounded_positive_float(value, maximum=3600.0, label="超时秒数")


def maximum_size_gb(value: str) -> float:
    return bounded_positive_float(value, maximum=1024.0, label="最大下载大小（GB）")


def add_input_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("text", nargs="?", help="视频号分享文字或链接")
    parser.add_argument("--input-file", help="从 UTF-8 文本文件读取分享内容")
    parser.add_argument("--resolver", help="临时覆盖解析器 URL")
    parser.add_argument("--official-profile", help="读取浏览器生成的腾讯官方解析临时文件")
    parser.add_argument(
        "--quality",
        choices=("auto", "original", "compatible", "h265"),
        default="auto",
        help="画质：auto 自动优先原画并回退，默认 auto",
    )
    parser.add_argument(
        "--timeout", type=timeout_seconds, default=30.0, help="网络超时秒数，默认 30，最大 3600"
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="通过分享链接下载微信视频号视频")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    parser.add_argument("--version", action="version", version=VERSION)
    subparsers = parser.add_subparsers(dest="command", required=True)

    doctor = subparsers.add_parser("doctor", help="检查本地运行环境")
    doctor.add_argument("--resolver", help="临时覆盖解析器 URL")
    doctor.add_argument("--output-dir", help="检查指定输出目录")
    doctor.set_defaults(handler=command_doctor)

    prepare = subparsers.add_parser("prepare", help="校验链接并创建官方解析临时文件")
    prepare.add_argument("text", nargs="?", help="视频号分享文字或链接")
    prepare.add_argument("--input-file", help="从 UTF-8 文本文件读取分享内容")
    prepare.set_defaults(handler=command_prepare)

    resolve = subparsers.add_parser("resolve", help="解析链接但不下载")
    add_input_arguments(resolve)
    resolve.set_defaults(handler=command_resolve)

    download = subparsers.add_parser("download", help="解析并下载视频")
    add_input_arguments(download)
    download.add_argument("--output-dir", help="视频保存目录")
    download.add_argument("--filename", help="目标文件名")
    download.add_argument(
        "--max-size-gb",
        type=maximum_size_gb,
        default=8.0,
        help="最大下载大小，默认 8 GB，最大 1024 GB",
    )
    download.set_defaults(handler=command_download)

    config = subparsers.add_parser("config", help="管理解析器配置")
    config_subparsers = config.add_subparsers(dest="config_command", required=True)
    config_show = config_subparsers.add_parser("show", help="显示当前配置")
    config_show.set_defaults(handler=command_config)
    config_set = config_subparsers.add_parser("set-resolver", help="设置默认解析器")
    config_set.add_argument("url")
    config_set.set_defaults(handler=command_config)
    config_reset = config_subparsers.add_parser("reset", help="恢复默认配置")
    config_reset.set_defaults(handler=command_config)

    self_test = subparsers.add_parser("self-test", help="运行无网络内置测试")
    self_test.set_defaults(handler=command_self_test)
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    try:
        if sys.version_info < (3, 9):
            raise ChannelsError("PYTHON_TOO_OLD", "需要 Python 3.9 或更高版本。")
        result = args.handler(args)
        print_result(result, args.json)
        return 0 if result.get("ok", False) else 1
    except ChannelsError as exc:
        result = exc.as_dict()
        print_result(result, args.json)
        return 2
    except OSError as exc:
        result = ChannelsError(
            "LOCAL_IO_ERROR",
            "读取或写入本地文件时发生错误。",
            hint="检查路径、磁盘空间和文件权限后重试。",
            details={"error": str(exc)},
        ).as_dict()
        print_result(result, args.json)
        return 2
    except KeyboardInterrupt:
        result = ChannelsError("CANCELLED", "操作已取消。", hint="临时下载文件已清理。").as_dict()
        print_result(result, args.json)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
