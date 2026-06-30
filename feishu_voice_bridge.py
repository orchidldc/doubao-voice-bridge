#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Feishu Voice Bridge

在 Windows 上轮询一个指定的飞书新版云文档 docx，把 F8 后新增的纯文本
粘贴到用户启动时捕获的目标输入窗口。它不做语音识别，不接豆包/火山/OCR。
"""

from __future__ import annotations

import argparse
import base64
import ctypes
from ctypes import wintypes
import html
import hashlib
import io
import json
import logging
import os
import re
import shutil
import struct
import subprocess
import sys
import threading
import time
import uuid
import zlib
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional
from urllib.parse import urlparse

try:
    import requests
except Exception:  # pragma: no cover - dependency check happens at runtime.
    requests = None  # type: ignore[assignment]

try:
    import pyperclip
except Exception:  # pragma: no cover
    pyperclip = None  # type: ignore[assignment]

try:
    from PIL import Image
except Exception:  # pragma: no cover
    Image = None  # type: ignore[assignment]

try:
    import keyboard
except Exception:  # pragma: no cover
    keyboard = None  # type: ignore[assignment]

try:
    import win32gui
except Exception:  # pragma: no cover
    win32gui = None  # type: ignore[assignment]

try:
    import win32api
    import win32con
except Exception:  # pragma: no cover
    win32api = None  # type: ignore[assignment]
    win32con = None  # type: ignore[assignment]

try:
    import win32process
except Exception:  # pragma: no cover
    win32process = None  # type: ignore[assignment]

try:
    import win32clipboard
except Exception:  # pragma: no cover
    win32clipboard = None  # type: ignore[assignment]

try:
    import winreg
except Exception:  # pragma: no cover
    winreg = None  # type: ignore[assignment]


FEISHU_HOST = "https://open.feishu.cn"
DEFAULT_TEST_TEXT = "这是飞书云文档语音桥接工具的固定粘贴测试。"
FEISHU_DOCX_READ_SCOPE = "docx:document:readonly"
FEISHU_DOCX_WRITE_SCOPE = "docx:document"
DELETE_CHILDREN_BATCH_SIZE = 200
TAIL_ANCHOR_MIN_CHARS = 24
TAIL_ANCHOR_MAX_CHARS = 160
TAIL_REWRITE_TOLERANCE_CHARS = 12
TAIL_REWRITE_MIN_PREFIX_RATIO = 0.72
TAIL_CONTEXT_MIN_CHARS = 16
TAIL_CONTEXT_MAX_SUFFIX_CHARS = 96
TAIL_CONTEXT_MAX_DRIFT_CHARS = 360
TAIL_CONTEXT_STEP_CHARS = 8
DEFAULT_CONFIG: dict[str, Any] = {
    "auth_mode": "lark_cli",
    "lark_cli_path": "lark-cli",
    "lark_cli_profile": "",
    "doc_url": "https://xxx.feishu.cn/docx/xxxxxxxx",
    "poll_interval_seconds": 1.2,
    "stable_delay_ms": 1000,
    "min_chars_to_paste": 1,
    "trim_trailing_whitespace_for_delta": True,
    "append_tolerance_chars": 2,
    "leading_newline_policy": "smart",
    "leading_newline_idle_threshold_seconds": 2.0,
    "monitor_after_marker": "",
    "insert_mode": "clipboard",
    "restore_clipboard": True,
    "enable_image_bridge": False,
    "image": {
        "enabled": False,
        "detect_route": "cli",
        "insert_mode": "clipboard_bitmap",
        "allow_file_drop_fallback": False,
        "stable_delay_ms": 1500,
        "download_retry_count": 3,
        "download_retry_interval_ms": 800,
        "temp_dir": "./data/images",
        "keep_downloaded_images_hours": 24,
        "max_image_size_mb": 30,
        "restore_clipboard": True,
        "restore_clipboard_delay_ms": 1500,
    },
    "target_window_mode": "locked",
    "require_same_foreground_window": True,
    "require_same_focused_control": False,
    "allow_refocus_target_window": False,
    "log_level": "INFO",
    "auto_start_watch": True,
    "auto_begin_on_watch": True,
    "hotkeys": {
        "start": "F8",
        "stop": "F9",
        "reset_baseline": "F10",
        "quit": "F12",
    },
}


class ConfigError(RuntimeError):
    pass


class FeishuApiError(RuntimeError):
    def __init__(self, action: str, code: Any, msg: Any):
        super().__init__(f"{action}失败：code={code} msg={msg}")
        self.action = action
        self.code = code
        self.msg = msg


class MarkerMissingError(RuntimeError):
    pass


def set_dpi_aware() -> None:
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass


def setup_logging(level: str) -> None:
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%H:%M:%S",
        force=True,
    )


def load_json(path: Path, required: bool = True) -> dict[str, Any]:
    if not path.exists():
        if required:
            raise FileNotFoundError(f"配置文件不存在：{path}")
        return {}
    with path.open("r", encoding="utf-8-sig") as f:
        data = json.load(f)
    if not isinstance(data, dict):
        raise ConfigError("配置文件根节点必须是 JSON object。")
    return data


def get_auth_mode(config: dict[str, Any]) -> str:
    raw = str(config.get("auth_mode", "lark_cli") or "lark_cli").strip().lower()
    if raw in {"lark_cli", "cli", "lark-cli", "feishu_cli", "feishu-cli"}:
        return "lark_cli"
    if raw in {"app", "tenant", "app_secret"}:
        return "app"
    raise ConfigError("auth_mode 只能是 lark_cli 或 app。")


def resolve_config_path(raw_path: str) -> Path:
    path = Path(raw_path)
    if path.is_absolute() or path.exists():
        return path
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        packaged_path = exe_dir / raw_path
        return packaged_path
    return path


def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def ensure_config_for_double_click(config_path: Path) -> bool:
    if config_path.exists():
        return False

    config_path.parent.mkdir(parents=True, exist_ok=True)
    example_candidates = [
        config_path.parent / "config.example.json",
        app_dir() / "config.example.json",
    ]
    for example in example_candidates:
        if example.exists():
            shutil.copyfile(example, config_path)
            return True

    with config_path.open("w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        f.write("\n")
    return True


def config_not_ready_reasons(config: dict[str, Any]) -> list[str]:
    reasons: list[str] = []
    auth_mode = get_auth_mode(config)

    source = str(config.get("document_id") or config.get("doc_url") or "").strip()
    if not source:
        reasons.append("config.json 里缺少 doc_url 或 document_id。")
    elif "xxx.feishu.cn" in source or "xxxxxxxx" in source:
        reasons.append("config.json 里的 doc_url 还是示例值，请替换成真实飞书 docx 链接。")
    else:
        try:
            extract_docx_id(source)
        except Exception as exc:
            reasons.append(str(exc))

    if auth_mode == "app":
        if not get_env_value("FEISHU_APP_ID"):
            reasons.append("auth_mode=app 时缺少环境变量 FEISHU_APP_ID。")
        if not get_env_value("FEISHU_APP_SECRET"):
            reasons.append("auth_mode=app 时缺少环境变量 FEISHU_APP_SECRET。")
    else:
        reasons.extend(check_lark_cli_ready(config))
    return reasons


def print_double_click_help(config_path: Path, created_config: bool, reasons: list[str]) -> None:
    print()
    print("飞书云文档语音桥接工具")
    print("=" * 30)
    print(f"配置文件：{config_path}")
    if created_config:
        print("已自动创建 config.json。请先打开它，把 doc_url 改成真实飞书 docx 链接。")
    if reasons:
        print()
        print("当前还不能启动监听，需要先处理：")
        for index, reason in enumerate(reasons, 1):
            print(f"{index}. {reason}")
        print()
        print("最小配置步骤：")
        print("1. 编辑 config.json，填写 https://xxx.feishu.cn/docx/xxxxxxxx 这种新版文档链接。")
        print("2. 确保本机 lark-cli 已登录你的飞书账号。")
        print("3. 重新打开本 exe。配置完整后，双击会直接进入监听模式。")
        print()
        print("飞书 CLI 登录/授权示例：")
        print("lark-cli auth status --json")
        print('lark-cli auth login --scope "docx:document:readonly offline_access"')
        print('清空正文按钮还需要写权限，可运行：lark-cli auth login --scope "docx:document offline_access"')
        print()
        print("如果你仍想用自建应用密钥模式，把 config.json 里的 auth_mode 改成 app，")
        print("再设置 FEISHU_APP_ID 和 FEISHU_APP_SECRET。")
        print()
        print("处理完成后请重新双击 exe。")
    else:
        print("配置看起来完整，将进入监听模式。")
        print("把光标放到目标输入框，按 F8 开始；F9 暂停；F10 重设 baseline；F12 退出。")


def pause_before_exit() -> None:
    try:
        input("\n按 Enter 退出...")
    except Exception:
        pass


def sha12(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:12]


def normalize_text(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "\n")


def mask_secret(value: str) -> str:
    if not value:
        return "<empty>"
    if len(value) <= 8:
        return "***"
    return f"{value[:4]}...{value[-4:]}"


def get_env_value(name: str) -> str:
    value = os.getenv(name, "").strip()
    if value or winreg is None:
        return value
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as key:  # type: ignore[union-attr]
            registry_value, _value_type = winreg.QueryValueEx(key, name)  # type: ignore[union-attr]
        return str(registry_value).strip()
    except Exception:
        return ""


def lark_cli_base_args(config: dict[str, Any]) -> list[str]:
    cli_path = str(config.get("lark_cli_path", "lark-cli") or "lark-cli").strip()
    profile = str(config.get("lark_cli_profile", "") or "").strip()
    resolved = shutil.which(cli_path)
    args = [resolved or cli_path]
    if profile:
        args.extend(["--profile", profile])
    return args


def lark_cli_exists(config: dict[str, Any]) -> bool:
    cli_path = str(config.get("lark_cli_path", "lark-cli") or "lark-cli").strip()
    path_obj = Path(cli_path)
    if path_obj.is_absolute() or "\\" in cli_path or "/" in cli_path:
        return path_obj.exists()
    return shutil.which(cli_path) is not None


def run_lark_cli_json(config: dict[str, Any], args: list[str], timeout: int = 30) -> dict[str, Any]:
    command = lark_cli_base_args(config) + args
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
        )
    except FileNotFoundError as exc:
        raise RuntimeError("找不到 lark-cli。请先安装飞书 CLI，或在 config.json 设置 lark_cli_path。") from exc
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError("lark-cli 调用超时。") from exc

    stdout = (completed.stdout or "").strip()
    stderr = (completed.stderr or "").strip()
    if completed.returncode != 0:
        detail = stderr or stdout or f"exit={completed.returncode}"
        raise RuntimeError(f"lark-cli 调用失败：{detail}")
    if not stdout:
        raise RuntimeError("lark-cli 没有返回 JSON 输出。")

    try:
        parsed = json.loads(stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"lark-cli 返回的不是 JSON：{stdout[:300]}") from exc
    if not isinstance(parsed, dict):
        raise RuntimeError("lark-cli JSON 输出不是 object。")
    return parsed


def lark_cli_missing_scopes(config: dict[str, Any], scope: str, timeout: int = 15) -> list[str]:
    command = lark_cli_base_args(config) + ["auth", "check", "--scope", scope, "--json"]
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            check=False,
        )
    except Exception:
        return [scope]

    stdout = (completed.stdout or "").strip()
    try:
        parsed = json.loads(stdout) if stdout else {}
    except json.JSONDecodeError:
        return [scope] if completed.returncode != 0 else []

    missing = parsed.get("missing")
    if isinstance(missing, list):
        return [str(item) for item in missing]
    if parsed.get("ok") is True or parsed.get("granted") is True:
        return []
    return [scope] if completed.returncode != 0 else []


def check_lark_cli_ready(config: dict[str, Any]) -> list[str]:
    reasons: list[str] = []
    if not lark_cli_exists(config):
        reasons.append("找不到 lark-cli。请先安装飞书 CLI，或在 config.json 设置 lark_cli_path。")
        return reasons

    try:
        status = run_lark_cli_json(config, ["auth", "status", "--json"], timeout=15)
    except Exception as exc:
        reasons.append(f"无法读取 lark-cli 登录状态：{exc}")
        return reasons

    user = (status.get("identities") or {}).get("user") if isinstance(status.get("identities"), dict) else None
    user_status = user.get("status") if isinstance(user, dict) else None
    if not isinstance(user, dict) or not user.get("available") or user_status not in {"ready", "needs_refresh"}:
        reasons.append(f'lark-cli 用户身份未登录。请运行：lark-cli auth login --scope "{FEISHU_DOCX_READ_SCOPE} offline_access"')
        return reasons

    try:
        check = run_lark_cli_json(config, ["auth", "check", "--scope", FEISHU_DOCX_READ_SCOPE, "--json"], timeout=15)
    except Exception as exc:
        reasons.append(f"无法检查 lark-cli 文档读取权限：{exc}")
        return reasons

    missing = check.get("missing")
    if missing:
        reasons.append(f'lark-cli 缺少 {FEISHU_DOCX_READ_SCOPE} 权限。请运行：lark-cli auth login --scope "{FEISHU_DOCX_READ_SCOPE} offline_access"')
    return reasons


def extract_docx_id(doc_url_or_id: str) -> str:
    raw = (doc_url_or_id or "").strip()
    if not raw:
        raise ConfigError("缺少 doc_url 或 document_id。")

    if re.fullmatch(r"[A-Za-z0-9_-]{6,}", raw):
        return raw

    parsed = urlparse(raw)
    path = parsed.path or ""
    if "/wiki/" in path:
        raise ConfigError("暂不支持 wiki 链接。请先填写新版 docx 链接，形如 https://xxx.feishu.cn/docx/xxxxxxxx。")
    if parsed.scheme and "feishu.cn" not in parsed.netloc and "larksuite.com" not in parsed.netloc:
        raise ConfigError("doc_url 必须是飞书/larksuite 文档链接，或直接填写 document_id。")

    match = re.search(r"/docx/([A-Za-z0-9_-]+)", path)
    if match:
        return match.group(1)

    if parsed.scheme:
        raise ConfigError("无法解析 document_id：当前只支持新版飞书 docx 链接，不支持旧版 doc/wiki/sheets 链接。")
    raise ConfigError("无法解析 document_id。请填写 docx 链接或直接填写 document_id。")


def extract_scope_text(text: str, marker: str) -> str:
    normalized = normalize_text(text)
    if not marker:
        return normalized
    idx = normalized.rfind(marker)
    if idx < 0:
        raise MarkerMissingError(f"未找到 monitor_after_marker：{marker!r}。不会监控整篇文档。")
    return normalized[idx + len(marker):]


def common_prefix_len(a: str, b: str) -> int:
    limit = min(len(a), len(b))
    i = 0
    while i < limit and a[i] == b[i]:
        i += 1
    return i


def require_dependency(value: Any, package_name: str) -> None:
    if value is None:
        raise RuntimeError(f"缺少依赖 {package_name}。请先运行：pip install -r requirements.txt")


@dataclass(frozen=True)
class FeishuConfig:
    app_id: str
    app_secret: str
    document_id: str


@dataclass(frozen=True)
class TailAnchorDelta:
    delta: str
    anchor_len: int
    offset_drift: int


@dataclass(frozen=True)
class NearTailRewriteDelta:
    delta: str
    rewritten_tail_len: int


def find_tail_anchor_append_delta(previous: str, current: str) -> Optional[TailAnchorDelta]:
    """Recover append-only deltas when Feishu rewrites earlier paragraph text."""
    if len(current) <= len(previous):
        return None

    base = previous.rstrip()
    if len(base) < TAIL_ANCHOR_MIN_CHARS:
        return None

    candidate_lengths = {
        min(TAIL_ANCHOR_MAX_CHARS, len(base)),
        120,
        80,
        48,
        32,
        TAIL_ANCHOR_MIN_CHARS,
    }
    plausible_shift = max(240, abs(len(current) - len(previous)) + 80)
    for anchor_len in sorted((n for n in candidate_lengths if TAIL_ANCHOR_MIN_CHARS <= n <= len(base)), reverse=True):
        anchor = base[-anchor_len:]
        if len(anchor.strip()) < TAIL_ANCHOR_MIN_CHARS:
            continue

        current_anchor_start = current.rfind(anchor)
        if current_anchor_start < 0:
            continue

        delta_start = current_anchor_start + anchor_len
        delta = current[delta_start:]
        if not delta:
            continue

        expected_anchor_start = len(base) - anchor_len
        drift = current_anchor_start - expected_anchor_start
        if abs(drift) > plausible_shift:
            continue

        return TailAnchorDelta(delta=delta, anchor_len=anchor_len, offset_drift=drift)
    return None


def find_near_tail_rewrite_append_delta(previous: str, current: str, lcp: int) -> Optional[NearTailRewriteDelta]:
    """Treat a tiny rewrite at the old tail as append-only input.

    Older mobile clients can rewrite the final character or separator while appending
    text. In that case a strict prefix check fails although the user only appended.
    """
    if len(current) <= len(previous) or not previous:
        return None

    rewritten_tail_len = len(previous) - lcp
    if rewritten_tail_len <= 0 or rewritten_tail_len > TAIL_REWRITE_TOLERANCE_CHARS:
        return None

    min_prefix_len = min(6, max(1, int(len(previous) * TAIL_REWRITE_MIN_PREFIX_RATIO)))
    if lcp < min_prefix_len:
        return None

    delta = current[len(previous):]
    if not delta or not delta.strip():
        return None
    return NearTailRewriteDelta(delta=delta, rewritten_tail_len=rewritten_tail_len)


def find_tail_context_append_delta(previous: str, current: str) -> Optional[TailAnchorDelta]:
    """Recover appends when the old document tail was lightly rewritten.

    Instead of anchoring only on the exact old suffix, scan the last small window of
    the previous text for a stable context anchor and trim the old suffix after it.
    """
    if len(current) <= len(previous):
        return None

    base = previous.rstrip()
    if len(base) < TAIL_CONTEXT_MIN_CHARS:
        return None

    candidate_lengths = {
        min(TAIL_ANCHOR_MAX_CHARS, len(base)),
        120,
        80,
        48,
        32,
        24,
        TAIL_CONTEXT_MIN_CHARS,
    }
    max_drift = min(TAIL_CONTEXT_MAX_DRIFT_CHARS, max(120, len(base) // 3))

    for anchor_len in sorted((n for n in candidate_lengths if TAIL_CONTEXT_MIN_CHARS <= n <= len(base)), reverse=True):
        latest_start = len(base) - anchor_len
        earliest_start = max(0, latest_start - TAIL_CONTEXT_MAX_SUFFIX_CHARS)
        starts = list(range(latest_start, earliest_start - 1, -TAIL_CONTEXT_STEP_CHARS))
        if starts[-1] != earliest_start:
            starts.append(earliest_start)

        for anchor_start in starts:
            anchor = base[anchor_start:anchor_start + anchor_len]
            if len(anchor.strip()) < TAIL_CONTEXT_MIN_CHARS:
                continue

            previous_suffix = base[anchor_start + anchor_len:]
            low = max(0, anchor_start - max_drift)
            high = min(len(current), anchor_start + max_drift + anchor_len)
            current_anchor_start = current.rfind(anchor, low, high)
            if current_anchor_start < 0:
                continue

            after_anchor = current[current_anchor_start + anchor_len:]
            if previous_suffix:
                if after_anchor.startswith(previous_suffix):
                    delta = after_anchor[len(previous_suffix):]
                else:
                    if len(after_anchor) < len(previous_suffix):
                        continue
                    suffix_lcp = common_prefix_len(previous_suffix, after_anchor)
                    rewritten_suffix_len = len(previous_suffix) - suffix_lcp
                    suffix_mostly_matches = suffix_lcp >= int(len(previous_suffix) * TAIL_REWRITE_MIN_PREFIX_RATIO)
                    if rewritten_suffix_len > TAIL_REWRITE_TOLERANCE_CHARS and not suffix_mostly_matches:
                        continue
                    delta = after_anchor[len(previous_suffix):]
            else:
                delta = after_anchor

            if not delta or not delta.strip():
                continue

            drift = current_anchor_start - anchor_start
            if abs(drift) > max_drift:
                continue

            return TailAnchorDelta(delta=delta, anchor_len=anchor_len, offset_drift=drift)
    return None


@dataclass(frozen=True)
class TargetWindow:
    hwnd: Optional[int]
    focus_hwnd: Optional[int]
    title: str
    process_name: str

    @property
    def summary(self) -> str:
        return f"hwnd={self.hwnd} focus={self.focus_hwnd} process={self.process_name!r} title={self.title!r}"


@dataclass(frozen=True)
class BridgeSettings:
    poll_interval_seconds: float = 1.2
    stable_delay_ms: int = 1000
    min_chars_to_paste: int = 1
    trim_trailing_whitespace_for_delta: bool = True
    append_tolerance_chars: int = 2
    leading_newline_policy: str = "smart"
    leading_newline_idle_threshold_seconds: float = 2.0
    monitor_after_marker: str = ""
    insert_mode: str = "clipboard"
    restore_clipboard: bool = True
    target_window_mode: str = "locked"
    require_same_foreground_window: bool = True
    require_same_focused_control: bool = False
    allow_refocus_target_window: bool = False
    log_level: str = "INFO"

    @classmethod
    def from_config(cls, config: dict[str, Any]) -> "BridgeSettings":
        settings = cls(
            poll_interval_seconds=float(config.get("poll_interval_seconds", 1.2)),
            stable_delay_ms=int(config.get("stable_delay_ms", 1000)),
            min_chars_to_paste=int(config.get("min_chars_to_paste", 1)),
            trim_trailing_whitespace_for_delta=bool(config.get("trim_trailing_whitespace_for_delta", True)),
            append_tolerance_chars=int(config.get("append_tolerance_chars", 2)),
            leading_newline_policy=str(config.get("leading_newline_policy", "smart") or "smart").lower(),
            leading_newline_idle_threshold_seconds=float(config.get("leading_newline_idle_threshold_seconds", 2.0)),
            monitor_after_marker=str(config.get("monitor_after_marker", "") or ""),
            insert_mode=str(config.get("insert_mode", "clipboard") or "clipboard").lower(),
            restore_clipboard=bool(config.get("restore_clipboard", True)),
            target_window_mode=str(config.get("target_window_mode", "locked") or "locked").lower(),
            require_same_foreground_window=bool(config.get("require_same_foreground_window", True)),
            require_same_focused_control=bool(config.get("require_same_focused_control", False)),
            allow_refocus_target_window=bool(config.get("allow_refocus_target_window", False)),
            log_level=str(config.get("log_level", "INFO")),
        )
        if settings.poll_interval_seconds <= 0:
            raise ConfigError("poll_interval_seconds 必须大于 0。")
        if settings.poll_interval_seconds < 0.8:
            logging.warning("poll_interval_seconds=%s 偏低，可能触发飞书接口限频。", settings.poll_interval_seconds)
        if settings.stable_delay_ms < 0:
            raise ConfigError("stable_delay_ms 不能为负数。")
        if settings.min_chars_to_paste < 1:
            raise ConfigError("min_chars_to_paste 必须至少为 1。")
        if settings.append_tolerance_chars < 0:
            raise ConfigError("append_tolerance_chars 不能为负数。")
        if settings.leading_newline_policy not in {"smart", "keep", "strip"}:
            raise ConfigError("leading_newline_policy 只能是 smart、keep 或 strip。")
        if settings.leading_newline_idle_threshold_seconds < 0:
            raise ConfigError("leading_newline_idle_threshold_seconds 不能为负数。")
        if settings.insert_mode not in {"clipboard", "sendinput"}:
            raise ConfigError("insert_mode 只能是 clipboard 或 sendinput。")
        if settings.target_window_mode not in {"locked", "process", "any"}:
            raise ConfigError("target_window_mode 只能是 locked、process 或 any。")
        return settings


def config_bool(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}
    return bool(value)


@dataclass(frozen=True)
class ImageSettings:
    enabled: bool = False
    detect_route: str = "cli"
    insert_mode: str = "clipboard_bitmap"
    allow_file_drop_fallback: bool = False
    stable_delay_ms: int = 1500
    download_retry_count: int = 3
    download_retry_interval_ms: int = 800
    temp_dir: str = "./data/images"
    keep_downloaded_images_hours: int = 24
    max_image_size_mb: int = 30
    restore_clipboard: bool = True
    restore_clipboard_delay_ms: int = 1500

    @classmethod
    def from_config(cls, config: dict[str, Any]) -> "ImageSettings":
        raw_image = config.get("image")
        image = raw_image if isinstance(raw_image, dict) else {}
        settings = cls(
            enabled=config_bool(config.get("enable_image_bridge"), False) or config_bool(image.get("enabled"), False),
            detect_route=str(image.get("detect_route", "cli") or "cli").lower(),
            insert_mode=str(image.get("insert_mode", "clipboard_bitmap") or "clipboard_bitmap").lower(),
            allow_file_drop_fallback=config_bool(image.get("allow_file_drop_fallback"), False),
            stable_delay_ms=int(image.get("stable_delay_ms", 1500)),
            download_retry_count=int(image.get("download_retry_count", 3)),
            download_retry_interval_ms=int(image.get("download_retry_interval_ms", 800)),
            temp_dir=str(image.get("temp_dir", "./data/images") or "./data/images"),
            keep_downloaded_images_hours=int(image.get("keep_downloaded_images_hours", 24)),
            max_image_size_mb=int(image.get("max_image_size_mb", 30)),
            restore_clipboard=config_bool(image.get("restore_clipboard"), True),
            restore_clipboard_delay_ms=int(image.get("restore_clipboard_delay_ms", 1500)),
        )
        if settings.detect_route != "cli":
            raise ConfigError("image.detect_route currently only supports cli.")
        if settings.insert_mode not in {"auto", "clipboard_bitmap", "clipboard_file", "save_only"}:
            raise ConfigError("image.insert_mode must be auto, clipboard_bitmap, clipboard_file, or save_only.")
        if settings.stable_delay_ms < 0:
            raise ConfigError("image.stable_delay_ms cannot be negative.")
        if settings.download_retry_count < 1:
            raise ConfigError("image.download_retry_count must be at least 1.")
        if settings.download_retry_interval_ms < 0:
            raise ConfigError("image.download_retry_interval_ms cannot be negative.")
        if settings.keep_downloaded_images_hours < 0:
            raise ConfigError("image.keep_downloaded_images_hours cannot be negative.")
        if settings.max_image_size_mb < 1:
            raise ConfigError("image.max_image_size_mb must be at least 1.")
        return settings


@dataclass(frozen=True)
class ImageEvent:
    event_id: str
    token: str
    occurrence_index: int
    width: Optional[int]
    height: Optional[int]
    tag_position: int
    stable_id: str
    name: str = ""
    mime: str = ""

    @property
    def token_masked(self) -> str:
        return mask_secret(self.token)


@dataclass
class PendingImage:
    event: ImageEvent
    detected_at: float
    attempts: int = 0


@dataclass(frozen=True)
class DownloadedImage:
    path: Path
    size_bytes: int
    sha256_8: str
    kind: str
    width: Optional[int]
    height: Optional[int]


def parse_int_attr(value: str) -> Optional[int]:
    try:
        return int(float(value))
    except Exception:
        return None


def parse_image_attrs(raw_attrs: str) -> dict[str, str]:
    attrs: dict[str, str] = {}
    pattern = re.compile(r"([A-Za-z0-9_:-]+)\s*=\s*(?:\"([^\"]*)\"|'([^']*)')")
    for match in pattern.finditer(raw_attrs):
        value = match.group(2) if match.group(2) is not None else match.group(3)
        attrs[match.group(1)] = value or ""
    return attrs


def pick_attr(attrs: dict[str, str], names: tuple[str, ...]) -> str:
    lowered = {key.lower(): value for key, value in attrs.items()}
    for name in names:
        value = lowered.get(name.lower())
        if value:
            return value
    return ""


def parse_image_events_from_content(content: str) -> list[ImageEvent]:
    events: list[ImageEvent] = []
    pattern = re.compile(r"<(?:img|image)\b(?P<attrs>[^>]*)/?>", re.IGNORECASE)
    for occurrence_index, match in enumerate(pattern.finditer(content or "")):
        attrs = parse_image_attrs(match.group("attrs"))
        token = pick_attr(attrs, ("token", "src", "file_token", "file-token", "image_key", "imagekey", "media_token", "mediatoken"))
        if not token:
            continue
        width = parse_int_attr(pick_attr(attrs, ("width",)))
        height = parse_int_attr(pick_attr(attrs, ("height",)))
        stable_id = pick_attr(attrs, ("block_id", "blockid", "id"))
        name = pick_attr(attrs, ("name", "filename", "file_name"))
        mime = pick_attr(attrs, ("mime", "mime_type", "type"))
        if stable_id:
            basis = f"id:{stable_id}"
        else:
            basis = f"{token}|{occurrence_index}|{width}|{height}|{match.start()}"
        events.append(
            ImageEvent(
                event_id=sha12(basis),
                token=token,
                occurrence_index=occurrence_index,
                width=width,
                height=height,
                tag_position=match.start(),
                stable_id=stable_id,
                name=name,
                mime=mime,
            )
        )
    return events


def strip_media_spans(content: str) -> str:
    without_paired = re.sub(
        r"<(?:img|image)\b[^>]*>.*?</(?:img|image)>",
        "",
        content or "",
        flags=re.IGNORECASE | re.DOTALL,
    )
    return re.sub(r"<(?:img|image)\b[^>]*/?>", "", without_paired, flags=re.IGNORECASE | re.DOTALL)


def document_content_to_plain_text_without_media(content: str) -> str:
    text = strip_media_spans(content)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</(?:p|h[1-6]|title|li|blockquote)>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    lines = [line.rstrip() for line in normalize_text(text).split("\n")]
    return "\n".join(lines).strip()


def extract_lark_fetch_content(data: dict[str, Any]) -> str:
    document = data.get("data", {}).get("document") if isinstance(data.get("data"), dict) else None
    if isinstance(document, dict) and isinstance(document.get("content"), str):
        return document["content"]
    nested_data = data.get("data")
    if isinstance(nested_data, dict) and isinstance(nested_data.get("content"), str):
        return nested_data["content"]
    content = find_first_content(data)
    if content is not None:
        return content
    raise RuntimeError("lark-cli docs +fetch response does not contain document content.")


def resolve_image_temp_dir(raw_dir: str) -> Path:
    root = Path(raw_dir)
    if not root.is_absolute():
        root = app_dir() / root
    root = root.resolve()
    root.mkdir(parents=True, exist_ok=True)
    return root


def safe_image_output_stem(event_id: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9_.-]+", "_", event_id)
    return f"doubao_image_{clean[:32]}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def image_metadata(path: Path) -> tuple[str, Optional[int], Optional[int]]:
    data = path.read_bytes()
    if data.startswith(b"\x89PNG\r\n\x1a\n") and len(data) >= 24:
        width, height = struct.unpack(">II", data[16:24])
        return "png", width, height
    if data.startswith(b"\xff\xd8"):
        offset = 2
        while offset + 9 < len(data):
            if data[offset] != 0xFF:
                offset += 1
                continue
            marker = data[offset + 1]
            offset += 2
            if marker in {0xD8, 0xD9}:
                continue
            if offset + 2 > len(data):
                break
            segment_length = int.from_bytes(data[offset:offset + 2], "big")
            if segment_length < 2 or offset + segment_length > len(data):
                break
            if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}:
                height = int.from_bytes(data[offset + 3:offset + 5], "big")
                width = int.from_bytes(data[offset + 5:offset + 7], "big")
                return "jpg", width, height
            offset += segment_length
        return "jpg", None, None
    if data.startswith((b"GIF87a", b"GIF89a")) and len(data) >= 10:
        width, height = struct.unpack("<HH", data[6:10])
        return "gif", width, height
    if len(data) >= 12 and data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp", None, None
    return "unknown", None, None


def newest_matching_file(directory: Path, stem: str) -> Path:
    candidates = [path for path in directory.glob(f"{stem}*") if path.is_file()]
    if not candidates:
        raise RuntimeError("media-download did not create an image file.")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def sanitize_cli_output(text: str, *secrets: str) -> str:
    sanitized = text or ""
    for secret in secrets:
        if secret:
            sanitized = sanitized.replace(secret, mask_secret(secret))
    return sanitized.strip()[:500]


def cleanup_old_downloaded_images(settings: ImageSettings) -> None:
    if settings.keep_downloaded_images_hours <= 0:
        return
    root = resolve_image_temp_dir(settings.temp_dir)
    cutoff = time.time() - settings.keep_downloaded_images_hours * 3600
    for path in root.rglob("doubao_image_*"):
        try:
            if path.is_file() and path.stat().st_mtime < cutoff:
                path.unlink()
        except Exception as exc:
            logging.debug("Failed to clean old image temp file %s: %s", path, exc)


def write_test_png(path: Path, width: int = 320, height: int = 180) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            row.extend(((x * 255) // max(1, width - 1), (y * 255) // max(1, height - 1), 180))
        rows.append(bytes(row))

    def chunk(kind: bytes, payload: bytes) -> bytes:
        body = kind + payload
        return struct.pack(">I", len(payload)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    payload = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    png = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", payload) + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b"")
    path.write_bytes(png)


class FeishuClient:
    def __init__(self, cfg: FeishuConfig):
        require_dependency(requests, "requests")
        self.cfg = cfg
        self._tenant_token: Optional[str] = None
        self._refresh_at = 0.0
        self.session = requests.Session()  # type: ignore[union-attr]

    def _api_json(
        self,
        method: str,
        path: str,
        *,
        params: Optional[dict[str, Any]] = None,
        body: Optional[dict[str, Any]] = None,
        timeout: int = 20,
        action: str = "调用飞书接口",
    ) -> dict[str, Any]:
        token = self._ensure_token()
        url = f"{FEISHU_HOST}{path}"
        try:
            resp = self.session.request(
                method,
                url,
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                params=params,
                json=body,
                timeout=timeout,
            )
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:
            raise RuntimeError(f"{action}网络请求失败：{exc}") from exc

        if data.get("code") != 0:
            raise FeishuApiError(action, data.get("code"), data.get("msg"))
        return data

    def _ensure_token(self) -> str:
        now = time.time()
        if self._tenant_token and now < self._refresh_at:
            return self._tenant_token

        url = f"{FEISHU_HOST}/open-apis/auth/v3/tenant_access_token/internal"
        try:
            resp = self.session.post(
                url,
                json={"app_id": self.cfg.app_id, "app_secret": self.cfg.app_secret},
                timeout=15,
            )
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:
            raise RuntimeError(f"获取 tenant_access_token 网络请求失败：{exc}") from exc

        if data.get("code") != 0:
            raise FeishuApiError("获取 tenant_access_token", data.get("code"), data.get("msg"))

        token = data.get("tenant_access_token")
        expire = data.get("expire")
        if not isinstance(token, str) or not token:
            raise RuntimeError("飞书返回中没有 tenant_access_token。")
        if not isinstance(expire, int) or expire <= 0:
            raise RuntimeError("飞书返回中没有有效 expire，无法安全刷新 tenant_access_token。")

        safety_margin = min(120, max(10, int(expire * 0.1)))
        self._tenant_token = token
        self._refresh_at = now + max(1, expire - safety_margin)
        logging.debug("tenant_access_token 已刷新，token=%s expire=%ss", mask_secret(token), expire)
        return token

    def get_raw_content(self) -> str:
        token = self._ensure_token()
        url = f"{FEISHU_HOST}/open-apis/docx/v1/documents/{self.cfg.document_id}/raw_content"
        try:
            resp = self.session.get(
                url,
                headers={"Authorization": f"Bearer {token}"},
                timeout=15,
            )
            resp.raise_for_status()
            data = resp.json()
        except Exception as exc:
            raise RuntimeError(f"获取文档纯文本网络请求失败：{exc}") from exc

        if data.get("code") != 0:
            raise FeishuApiError("获取文档纯文本", data.get("code"), data.get("msg"))

        content = find_first_content(data.get("data", data))
        if content is None:
            raise RuntimeError("飞书 raw_content 响应中没有可识别的纯文本字段。")
        return normalize_text(content)

    def count_document_body_blocks(self) -> int:
        total = 0
        page_token = ""
        path = f"/open-apis/docx/v1/documents/{self.cfg.document_id}/blocks/{self.cfg.document_id}/children"
        while True:
            params: dict[str, Any] = {"page_size": 500, "document_revision_id": -1}
            if page_token:
                params["page_token"] = page_token
            data = self._api_json("GET", path, params=params, action="获取文档正文块")
            items, has_more, page_token = extract_page_items(data)
            total += len(items)
            if not has_more or not page_token:
                return total

    def delete_document_body_range(self, start_index: int, end_index: int) -> None:
        path = f"/open-apis/docx/v1/documents/{self.cfg.document_id}/blocks/{self.cfg.document_id}/children/batch_delete"
        self._api_json(
            "DELETE",
            path,
            params={"document_revision_id": -1, "client_token": str(uuid.uuid4())},
            body={"start_index": start_index, "end_index": end_index},
            action="删除文档正文块",
        )

    def clear_document_body(self) -> int:
        remaining = self.count_document_body_blocks()
        deleted = 0
        while remaining > 0:
            batch = min(remaining, DELETE_CHILDREN_BATCH_SIZE)
            self.delete_document_body_range(0, batch)
            deleted += batch
            remaining -= batch
            if remaining:
                time.sleep(0.4)
        return deleted


class LarkCliClient:
    def __init__(self, config: dict[str, Any], document_id: str):
        self.config = config
        self.document_id = document_id

    def _api_json(
        self,
        method: str,
        path: str,
        *,
        params: Optional[dict[str, Any]] = None,
        body: Optional[dict[str, Any]] = None,
        timeout: int = 30,
        action: str = "通过 lark-cli 调用飞书接口",
    ) -> dict[str, Any]:
        args = ["api", method, path, "--as", "user", "--json"]
        if params:
            args.extend(["--params", json.dumps(params, ensure_ascii=False)])
        if body is not None:
            args.extend(["--data", json.dumps(body, ensure_ascii=False)])
        data = run_lark_cli_json(self.config, args, timeout=timeout)
        if data.get("code") not in (None, 0):
            raise FeishuApiError(action, data.get("code"), data.get("msg"))
        return data

    def get_raw_content(self) -> str:
        path = f"/open-apis/docx/v1/documents/{self.document_id}/raw_content"
        data = self._api_json("GET", path, action="通过 lark-cli 获取文档纯文本")
        if data.get("code") not in (None, 0):
            raise FeishuApiError("通过 lark-cli 获取文档纯文本", data.get("code"), data.get("msg"))
        content = find_first_content(data.get("data", data))
        if content is None:
            raise RuntimeError("lark-cli raw_content 响应中没有可识别的纯文本字段。")
        return normalize_text(content)

    def get_document_content(self) -> str:
        source = str(self.config.get("doc_url") or self.config.get("document_id") or self.document_id)
        data = run_lark_cli_json(
            self.config,
            ["docs", "+fetch", "--api-version", "v2", "--doc", source, "--detail", "with-ids", "--format", "json"],
            timeout=45,
        )
        return extract_lark_fetch_content(data)

    def _ensure_write_scope(self) -> None:
        missing = lark_cli_missing_scopes(self.config, FEISHU_DOCX_WRITE_SCOPE)
        if missing:
            raise ConfigError(
                f'lark-cli 缺少 {FEISHU_DOCX_WRITE_SCOPE} 写权限。请运行：'
                f'lark-cli auth login --scope "{FEISHU_DOCX_WRITE_SCOPE} offline_access"'
            )

    def count_document_body_blocks(self) -> int:
        total = 0
        page_token = ""
        path = f"/open-apis/docx/v1/documents/{self.document_id}/blocks/{self.document_id}/children"
        while True:
            params: dict[str, Any] = {"page_size": 500, "document_revision_id": -1}
            if page_token:
                params["page_token"] = page_token
            data = self._api_json("GET", path, params=params, action="通过 lark-cli 获取文档正文块")
            items, has_more, page_token = extract_page_items(data)
            total += len(items)
            if not has_more or not page_token:
                return total

    def delete_document_body_range(self, start_index: int, end_index: int) -> None:
        path = f"/open-apis/docx/v1/documents/{self.document_id}/blocks/{self.document_id}/children/batch_delete"
        self._api_json(
            "DELETE",
            path,
            params={"document_revision_id": -1, "client_token": str(uuid.uuid4())},
            body={"start_index": start_index, "end_index": end_index},
            timeout=30,
            action="通过 lark-cli 删除文档正文块",
        )

    def clear_document_body(self) -> int:
        self._ensure_write_scope()
        remaining = self.count_document_body_blocks()
        deleted = 0
        while remaining > 0:
            batch = min(remaining, DELETE_CHILDREN_BATCH_SIZE)
            self.delete_document_body_range(0, batch)
            deleted += batch
            remaining -= batch
            if remaining:
                time.sleep(0.4)
        return deleted


def find_first_content(obj: Any) -> Optional[str]:
    if isinstance(obj, dict):
        for key in ("content", "raw_content", "text"):
            val = obj.get(key)
            if isinstance(val, str):
                return val
        for val in obj.values():
            got = find_first_content(val)
            if got is not None:
                return got
    if isinstance(obj, list):
        for item in obj:
            got = find_first_content(item)
            if got is not None:
                return got
    return None


def extract_page_items(data: dict[str, Any]) -> tuple[list[Any], bool, str]:
    payload: Any = data.get("data", data)
    if not isinstance(payload, dict):
        return [], False, ""

    items: Any = None
    for key in ("items", "children", "blocks"):
        value = payload.get(key)
        if isinstance(value, list):
            items = value
            break
    if items is None:
        items = []

    has_more = bool(payload.get("has_more"))
    next_page_token = payload.get("page_token") or payload.get("next_page_token") or payload.get("next_page")
    if not isinstance(next_page_token, str):
        next_page_token = ""
    return items, has_more, next_page_token


def _query_process_name(pid: int) -> str:
    if not sys.platform.startswith("win") or pid <= 0:
        return ""
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return f"pid:{pid}"
    try:
        size = wintypes.DWORD(32768)
        buf = ctypes.create_unicode_buffer(size.value)
        ok = kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size))
        if ok:
            return Path(buf.value).name
        return f"pid:{pid}"
    finally:
        kernel32.CloseHandle(handle)


def _get_focus_hwnd_for_window(hwnd: Optional[int]) -> Optional[int]:
    if not hwnd or not sys.platform.startswith("win") or win32process is None:
        return None

    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    class GUITHREADINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.c_ulong),
            ("flags", ctypes.c_ulong),
            ("hwndActive", ctypes.c_void_p),
            ("hwndFocus", ctypes.c_void_p),
            ("hwndCapture", ctypes.c_void_p),
            ("hwndMenuOwner", ctypes.c_void_p),
            ("hwndMoveSize", ctypes.c_void_p),
            ("hwndCaret", ctypes.c_void_p),
            ("rcCaret", RECT),
        ]

    try:
        thread_id, _pid = win32process.GetWindowThreadProcessId(hwnd)
        info = GUITHREADINFO()
        info.cbSize = ctypes.sizeof(GUITHREADINFO)
        if ctypes.windll.user32.GetGUIThreadInfo(thread_id, ctypes.byref(info)):
            return int(info.hwndFocus) if info.hwndFocus else None
    except Exception:
        return None
    return None


def get_foreground_window() -> TargetWindow:
    if win32gui is None:
        return TargetWindow(hwnd=None, focus_hwnd=None, title="", process_name="")
    try:
        hwnd = int(win32gui.GetForegroundWindow())
        title = win32gui.GetWindowText(hwnd) or ""
        pid = 0
        if win32process is not None:
            _thread_id, pid = win32process.GetWindowThreadProcessId(hwnd)
        return TargetWindow(
            hwnd=hwnd,
            focus_hwnd=_get_focus_hwnd_for_window(hwnd),
            title=title,
            process_name=_query_process_name(pid),
        )
    except Exception:
        return TargetWindow(hwnd=None, focus_hwnd=None, title="", process_name="")


def refocus_window(hwnd: Optional[int]) -> bool:
    if not hwnd or win32gui is None:
        return False
    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.15)
        return get_foreground_window().hwnd == hwnd
    except Exception:
        return False


class TextInserter:
    def __init__(self, mode: str = "clipboard", restore_clipboard: bool = True):
        self.mode = mode
        self.restore_clipboard = restore_clipboard

    def paste_text(self, text: str) -> None:
        if not text:
            return
        if self.mode == "clipboard":
            self._paste_by_clipboard(text)
            return
        self._paste_by_sendinput(text)

    def _paste_by_clipboard(self, text: str) -> None:
        require_dependency(pyperclip, "pyperclip")

        old_clip = None
        had_old_clip = False
        if self.restore_clipboard:
            try:
                old_clip = pyperclip.paste()  # type: ignore[union-attr]
                had_old_clip = True
            except Exception as exc:
                logging.warning("读取当前剪贴板失败，粘贴后无法恢复：%s", exc)

        pyperclip.copy(text)  # type: ignore[union-attr]
        time.sleep(0.08)
        send_ctrl_v()
        time.sleep(0.08)

        if self.restore_clipboard and had_old_clip:
            try:
                pyperclip.copy(old_clip)  # type: ignore[union-attr]
            except Exception as exc:
                logging.warning("恢复剪贴板失败：%s", exc)

    def _paste_by_sendinput(self, text: str) -> None:
        if not sys.platform.startswith("win"):
            raise RuntimeError("sendinput 模式当前只支持 Windows。")
        send_unicode_text(text)


STANDARD_CLIPBOARD_FORMATS = {
    1: "Text",
    2: "Bitmap",
    8: "DIB",
    13: "UnicodeText",
    15: "FileDrop",
    17: "DIBV5",
}


def require_clipboard_dependencies() -> None:
    require_dependency(win32clipboard, "pywin32 win32clipboard")
    require_dependency(win32con, "pywin32 win32con")


def open_clipboard_with_retry(retries: int = 12, delay: float = 0.12) -> None:
    require_clipboard_dependencies()
    last_exc: Optional[Exception] = None
    for _ in range(retries):
        try:
            win32clipboard.OpenClipboard()  # type: ignore[union-attr]
            return
        except Exception as exc:
            last_exc = exc
            time.sleep(delay)
    raise RuntimeError(f"Cannot open Windows clipboard: {last_exc}")


def list_clipboard_formats() -> list[dict[str, Any]]:
    require_clipboard_dependencies()
    formats: list[dict[str, Any]] = []
    open_clipboard_with_retry()
    try:
        fmt = 0
        while True:
            fmt = win32clipboard.EnumClipboardFormats(fmt)  # type: ignore[union-attr]
            if fmt == 0:
                break
            name = STANDARD_CLIPBOARD_FORMATS.get(fmt, "")
            if not name:
                try:
                    name = win32clipboard.GetClipboardFormatName(fmt)  # type: ignore[union-attr]
                except Exception:
                    name = f"Format{fmt}"
            formats.append({"id": fmt, "name": name})
    finally:
        win32clipboard.CloseClipboard()  # type: ignore[union-attr]
    return formats


def clipboard_has_image_format(formats: list[dict[str, Any]]) -> bool:
    ids = {int(item["id"]) for item in formats}
    names = {str(item["name"]).lower() for item in formats}
    return 8 in ids or 2 in ids or 17 in ids or "bitmap" in names or "dib" in names


def clipboard_has_text_format(formats: list[dict[str, Any]]) -> bool:
    ids = {int(item["id"]) for item in formats}
    names = {str(item["name"]).lower() for item in formats}
    return 1 in ids or 13 in ids or "text" in names or "unicodetext" in names


def clipboard_has_file_drop_format(formats: list[dict[str, Any]]) -> bool:
    ids = {int(item["id"]) for item in formats}
    names = {str(item["name"]).lower() for item in formats}
    return 15 in ids or "filedrop" in names or "filenamew" in names or "filename" in names


def image_file_to_dib_bytes(image_path: Path) -> bytes:
    require_dependency(Image, "Pillow")
    with Image.open(image_path) as img:  # type: ignore[union-attr]
        if img.mode in {"RGBA", "LA"} or ("transparency" in img.info):
            rgba = img.convert("RGBA")
            background = Image.new("RGB", rgba.size, (255, 255, 255))  # type: ignore[union-attr]
            background.paste(rgba, mask=rgba.getchannel("A"))
            img = background
        else:
            img = img.convert("RGB")
        output = io.BytesIO()
        img.save(output, "BMP")
    return output.getvalue()[14:]


def set_clipboard_dib(image_path: Path) -> list[dict[str, Any]]:
    dib_bytes = image_file_to_dib_bytes(image_path)
    open_clipboard_with_retry()
    try:
        win32clipboard.EmptyClipboard()  # type: ignore[union-attr]
        win32clipboard.SetClipboardData(win32con.CF_DIB, dib_bytes)  # type: ignore[union-attr]
    finally:
        win32clipboard.CloseClipboard()  # type: ignore[union-attr]
    formats = list_clipboard_formats()
    if not clipboard_has_image_format(formats):
        raise RuntimeError("CF_DIB write succeeded but clipboard has no image format.")
    if clipboard_has_file_drop_format(formats):
        raise RuntimeError("Bitmap mode unexpectedly set FileDrop format.")
    return formats


def set_clipboard_bitmap_with_dotnet(image_path: Path) -> list[dict[str, Any]]:
    resolved = image_path.resolve()
    powershell = shutil.which("powershell.exe") or shutil.which("powershell") or "powershell.exe"
    path_literal = str(resolved).replace("'", "''")
    script = f"""
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$imagePath = '{path_literal}'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
function Invoke-ClipboardRetry([scriptblock]$Block) {{
  for ($i = 0; $i -lt 10; $i++) {{
    try {{ return & $Block }} catch {{
      if ($i -eq 9) {{ throw }}
      Start-Sleep -Milliseconds 150
    }}
  }}
}}
$img = [System.Drawing.Image]::FromFile($imagePath)
try {{
  Invoke-ClipboardRetry {{ [System.Windows.Forms.Clipboard]::SetImage($img) }} | Out-Null
}} finally {{
  $img.Dispose()
}}
"""
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    completed = subprocess.run(
        [powershell, "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        timeout=20,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()[:500]
        raise RuntimeError(f".NET SetImage failed: {detail}")
    formats = list_clipboard_formats()
    if not clipboard_has_image_format(formats):
        raise RuntimeError(".NET SetImage completed but clipboard has no image format.")
    if clipboard_has_file_drop_format(formats):
        raise RuntimeError("Bitmap fallback unexpectedly set FileDrop format.")
    return formats


def set_clipboard_file_drop(image_path: Path) -> list[dict[str, Any]]:
    resolved = image_path.resolve()
    powershell = shutil.which("powershell.exe") or shutil.which("powershell") or "powershell.exe"
    path_literal = str(resolved).replace("'", "''")
    script = f"""
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$imagePath = '{path_literal}'
Add-Type -AssemblyName System.Windows.Forms
$files = New-Object System.Collections.Specialized.StringCollection
[void]$files.Add($imagePath)
$data = New-Object System.Windows.Forms.DataObject
$data.SetFileDropList($files)
[System.Windows.Forms.Clipboard]::SetDataObject($data, $true, 10, 200)
"""
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    completed = subprocess.run(
        [powershell, "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        timeout=20,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()[:500]
        raise RuntimeError(f"FileDrop clipboard write failed: {detail}")
    formats = list_clipboard_formats()
    if not clipboard_has_file_drop_format(formats):
        raise RuntimeError("FileDrop write completed but clipboard has no FileDrop format.")
    return formats


class ImageInserter:
    def paste_image(self, image: DownloadedImage, settings: ImageSettings) -> None:
        if settings.insert_mode == "save_only":
            logging.info(
                "Image saved only. file=%s size_bytes=%s sha256=%s",
                image.path.name,
                image.size_bytes,
                image.sha256_8,
            )
            return
        if not sys.platform.startswith("win"):
            raise RuntimeError("Image clipboard paste is only supported on Windows.")

        old_clip = None
        had_old_clip = False
        if settings.restore_clipboard and pyperclip is not None:
            try:
                old_clip = pyperclip.paste()  # type: ignore[union-attr]
                had_old_clip = True
            except Exception as exc:
                logging.debug("Could not read text clipboard before image paste: %s", exc)

        formats = self._set_windows_clipboard_image(image.path, settings.insert_mode, settings.allow_file_drop_fallback)
        file_fallback_allowed = (
            settings.insert_mode == "clipboard_file"
            or (
                settings.insert_mode == "auto"
                and settings.allow_file_drop_fallback
                and clipboard_has_file_drop_format(formats)
            )
        )
        if not clipboard_has_image_format(formats) and not file_fallback_allowed:
            raise RuntimeError("Clipboard has no bitmap image format; Ctrl+V was blocked.")
        if clipboard_has_text_format(formats) and not clipboard_has_image_format(formats):
            raise RuntimeError("Clipboard only has text formats; Ctrl+V was blocked to avoid pasting a file name.")

        time.sleep(0.08)
        send_ctrl_v()
        time.sleep(max(0, settings.restore_clipboard_delay_ms) / 1000.0)

        if settings.restore_clipboard and had_old_clip and pyperclip is not None:
            try:
                pyperclip.copy(old_clip)  # type: ignore[union-attr]
            except Exception as exc:
                logging.debug("Could not restore text clipboard after image paste: %s", exc)

    def set_clipboard_only(self, image_path: Path, mode: str = "clipboard_bitmap", allow_file_drop_fallback: bool = False) -> list[dict[str, Any]]:
        return self._set_windows_clipboard_image(image_path, mode, allow_file_drop_fallback)

    @staticmethod
    def _set_windows_clipboard_image(image_path: Path, mode: str, allow_file_drop_fallback: bool = False) -> list[dict[str, Any]]:
        if mode not in {"auto", "clipboard_bitmap", "clipboard_file"}:
            raise ConfigError("image insert mode cannot set clipboard: " + mode)
        resolved = image_path.resolve()
        if not resolved.exists():
            raise FileNotFoundError(str(resolved))
        if mode == "clipboard_file":
            return set_clipboard_file_drop(resolved)

        try:
            return set_clipboard_dib(resolved)
        except Exception as dib_exc:
            logging.warning("CF_DIB clipboard write failed, trying .NET SetImage. error=%s", dib_exc)
            try:
                return set_clipboard_bitmap_with_dotnet(resolved)
            except Exception as bitmap_exc:
                if mode == "auto" and allow_file_drop_fallback:
                    logging.warning("Bitmap clipboard failed; using explicit FileDrop fallback. error=%s", bitmap_exc)
                    return set_clipboard_file_drop(resolved)
                raise RuntimeError(
                    "目标应用可能不支持剪贴板图片粘贴，请尝试 clipboard_file 模式或手动上传。"
                ) from bitmap_exc


def send_ctrl_v() -> None:
    errors: list[str] = []
    try:
        send_ctrl_v_by_sendinput()
        return
    except Exception as exc:
        errors.append(f"SendInput: {exc}")
        logging.warning("SendInput 发送 Ctrl+V 失败，尝试 keybd_event。%s", exc)

    try:
        send_ctrl_v_by_keybd_event()
        return
    except Exception as exc:
        errors.append(f"keybd_event: {exc}")
        logging.warning("keybd_event 发送 Ctrl+V 失败，尝试 keyboard 库。%s", exc)

    try:
        send_ctrl_v_by_keyboard_library()
        return
    except Exception as exc:
        errors.append(f"keyboard: {exc}")
        raise RuntimeError("Ctrl+V 发送失败：" + " | ".join(errors)) from exc


def send_ctrl_v_by_sendinput() -> None:
    if not sys.platform.startswith("win"):
        raise RuntimeError("剪贴板粘贴热键当前只支持 Windows。")

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002
    VK_CONTROL = 0x11
    VK_V = 0x56
    ULONG_PTR = ctypes.c_size_t

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", ctypes.c_ushort),
            ("wScan", ctypes.c_ushort),
            ("dwFlags", ctypes.c_ulong),
            ("time", ctypes.c_ulong),
            ("dwExtraInfo", ULONG_PTR),
        ]

    class INPUT_UNION(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", ctypes.c_ulong), ("union", INPUT_UNION)]

    events = [
        KEYBDINPUT(VK_CONTROL, 0, 0, 0, 0),
        KEYBDINPUT(VK_V, 0, 0, 0, 0),
        KEYBDINPUT(VK_V, 0, KEYEVENTF_KEYUP, 0, 0),
        KEYBDINPUT(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0, 0),
    ]
    inputs = (INPUT * len(events))(*[INPUT(type=INPUT_KEYBOARD, union=INPUT_UNION(ki=event)) for event in events])
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.SendInput.argtypes = (wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
    user32.SendInput.restype = wintypes.UINT
    ctypes.set_last_error(0)
    sent = user32.SendInput(len(events), inputs, ctypes.sizeof(INPUT))
    if sent != len(events):
        error_code = ctypes.get_last_error()
        raise RuntimeError(f"仅发送 {sent}/{len(events)} 个输入事件，GetLastError={error_code}")


def send_ctrl_v_by_keybd_event() -> None:
    if win32api is None or win32con is None:
        raise RuntimeError("pywin32 win32api/win32con 不可用。")
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    time.sleep(0.02)
    try:
        win32api.keybd_event(ord("V"), 0, 0, 0)
        time.sleep(0.02)
        win32api.keybd_event(ord("V"), 0, win32con.KEYEVENTF_KEYUP, 0)
    finally:
        time.sleep(0.02)
        win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)


def send_ctrl_v_by_keyboard_library() -> None:
    if keyboard is None:
        raise RuntimeError("keyboard 模块不可用。")
    keyboard.press_and_release("ctrl+v")


def send_unicode_text(text: str) -> None:
    if not sys.platform.startswith("win"):
        raise RuntimeError("sendinput 模式当前只支持 Windows。")

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 0x0002
    KEYEVENTF_UNICODE = 0x0004
    ULONG_PTR = ctypes.c_size_t

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", ctypes.c_ushort),
            ("wScan", ctypes.c_ushort),
            ("dwFlags", ctypes.c_ulong),
            ("time", ctypes.c_ulong),
            ("dwExtraInfo", ULONG_PTR),
        ]

    class INPUT_UNION(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", ctypes.c_ulong), ("union", INPUT_UNION)]

    units = [
        int.from_bytes(text.encode("utf-16-le")[i:i + 2], "little")
        for i in range(0, len(text.encode("utf-16-le")), 2)
    ]
    user32 = ctypes.windll.user32
    for unit in units:
        for flags in (KEYEVENTF_UNICODE, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP):
            inp = INPUT(type=INPUT_KEYBOARD, union=INPUT_UNION(ki=KEYBDINPUT(0, unit, flags, 0, 0)))
            sent = user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
            if sent != 1:
                raise RuntimeError("SendInput 写入失败。")
        time.sleep(0.001)


def download_image_with_lark_cli(config: dict[str, Any], event: ImageEvent, settings: ImageSettings) -> DownloadedImage:
    root = resolve_image_temp_dir(settings.temp_dir)
    downloads_dir = root / "downloads"
    downloads_dir.mkdir(parents=True, exist_ok=True)
    stem = safe_image_output_stem(event.event_id)
    output_arg = str(Path("downloads") / stem)
    command = lark_cli_base_args(config) + [
        "docs",
        "+media-download",
        "--token",
        event.token,
        "--output",
        output_arg,
        "--overwrite",
    ]
    last_detail = ""
    for attempt in range(1, settings.download_retry_count + 1):
        completed = subprocess.run(
            command,
            cwd=str(root),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            timeout=60,
        )
        if completed.returncode == 0:
            path = newest_matching_file(downloads_dir, stem)
            size_bytes = path.stat().st_size
            max_bytes = settings.max_image_size_mb * 1024 * 1024
            if size_bytes > max_bytes:
                raise RuntimeError(f"Downloaded image is too large: {size_bytes} bytes.")
            digest = sha256_file(path)
            kind, width, height = image_metadata(path)
            return DownloadedImage(
                path=path,
                size_bytes=size_bytes,
                sha256_8=digest[:8],
                kind=kind,
                width=width,
                height=height,
            )
        last_detail = sanitize_cli_output((completed.stderr or "") + "\n" + (completed.stdout or ""), event.token)
        if attempt < settings.download_retry_count:
            time.sleep(settings.download_retry_interval_ms / 1000.0)
    raise RuntimeError(f"lark-cli media-download failed: {last_detail}")


class Bridge:
    def __init__(self, config: dict[str, Any]):
        auth_mode = get_auth_mode(config)
        document_id = extract_docx_id(str(config.get("document_id") or config.get("doc_url") or ""))
        self.settings = BridgeSettings.from_config(config)
        self.image_settings = ImageSettings.from_config(config)
        self.config = config
        if auth_mode == "app":
            app_id = get_env_value("FEISHU_APP_ID")
            app_secret = get_env_value("FEISHU_APP_SECRET")
            if not app_id or not app_secret:
                raise ConfigError("auth_mode=app 时缺少 FEISHU_APP_ID / FEISHU_APP_SECRET。请用环境变量设置，不要写入配置文件。")
            self.client = FeishuClient(FeishuConfig(app_id=app_id, app_secret=app_secret, document_id=document_id))
            logging.debug("已加载飞书 app 鉴权配置：app_id=%s app_secret=%s document_id=%s", app_id, mask_secret(app_secret), document_id)
        else:
            ready_reasons = check_lark_cli_ready(config)
            if ready_reasons:
                raise ConfigError("；".join(ready_reasons))
            self.client = LarkCliClient(config, document_id)
            logging.debug("已加载 lark-cli 用户鉴权配置：document_id=%s", document_id)
        if self.image_settings.enabled and not isinstance(self.client, LarkCliClient):
            raise ConfigError("Image bridge currently requires auth_mode=lark_cli.")
        self.inserter = TextInserter(self.settings.insert_mode, self.settings.restore_clipboard)
        self.image_inserter = ImageInserter()

        self.document_id = document_id
        self.active = False
        self.quitting = False
        self.last_seen = ""
        self.pending_text = ""
        self.pending_since = 0.0
        self.last_delta_at = 0.0
        self.image_seen_ids: set[str] = set()
        self.image_order: list[str] = []
        self.pending_images: list[PendingImage] = []
        self.image_baseline_ready = not self.image_settings.enabled
        self.target_window: Optional[TargetWindow] = None
        self.lock = threading.RLock()

    def fetch_scope_text(self) -> str:
        if isinstance(self.client, LarkCliClient):
            raw = document_content_to_plain_text_without_media(self.client.get_document_content())
        else:
            raw = document_content_to_plain_text_without_media(self.client.get_raw_content())
        text = extract_scope_text(raw, self.settings.monitor_after_marker)
        if self.settings.trim_trailing_whitespace_for_delta:
            return text.rstrip()
        return text

    def fetch_scope_text_and_image_events(self) -> tuple[str, Optional[list[ImageEvent]]]:
        if not isinstance(self.client, LarkCliClient):
            return self.fetch_scope_text(), None
        content = self.client.get_document_content()
        raw_text = document_content_to_plain_text_without_media(content)
        text = extract_scope_text(raw_text, self.settings.monitor_after_marker)
        if self.settings.trim_trailing_whitespace_for_delta:
            text = text.rstrip()
        image_events = parse_image_events_from_content(content) if self.image_settings.enabled else None
        return text, image_events

    def count_document_body_blocks(self) -> int:
        return self.client.count_document_body_blocks()

    def clear_document_body(self) -> int:
        return self.client.clear_document_body()

    def fetch_image_events(self) -> list[ImageEvent]:
        if not isinstance(self.client, LarkCliClient):
            raise RuntimeError("Image bridge requires lark-cli document fetch.")
        content = self.client.get_document_content()
        return parse_image_events_from_content(content)

    def _set_image_baseline(self, events: list[ImageEvent]) -> None:
        self.image_seen_ids = {event.event_id for event in events}
        self.image_order = [event.event_id for event in events]
        self.pending_images = []
        self.image_baseline_ready = True
        logging.info("Image baseline set. image_count=%s", len(events))

    def _append_image_events_or_reset(self, events: list[ImageEvent], now: float) -> None:
        if not self.image_settings.enabled:
            return
        current_order = [event.event_id for event in events]
        if not self.image_baseline_ready:
            self._set_image_baseline(events)
            return
        if len(current_order) < len(self.image_order) or current_order[:len(self.image_order)] != self.image_order:
            self._set_image_baseline(events)
            logging.warning("Image document order changed; image baseline reset and old pending images cleared.")
            return
        new_events = [event for event in events if event.event_id not in self.image_seen_ids]
        for event in new_events:
            self.image_seen_ids.add(event.event_id)
            self.image_order.append(event.event_id)
            self.pending_images.append(PendingImage(event=event, detected_at=now))
            logging.info(
                "Detected new image. event=%s occurrence=%s token=%s width=%s height=%s pending_images=%s",
                event.event_id,
                event.occurrence_index,
                event.token_masked,
                event.width,
                event.height,
                len(self.pending_images),
            )

    def _pop_ready_image(self, now: float) -> Optional[PendingImage]:
        if not self.image_settings.enabled or not self.pending_images:
            return None
        first = self.pending_images[0]
        if now - first.detected_at < self.image_settings.stable_delay_ms / 1000.0:
            return None
        if self.image_settings.insert_mode != "save_only" and not self._target_is_active():
            return None
        return self.pending_images.pop(0)

    def _handle_image_paste_failure(self, pending: PendingImage) -> None:
        pending.attempts += 1
        if pending.attempts < self.image_settings.download_retry_count:
            pending.detected_at = time.time()
            self.pending_images.insert(0, pending)
        else:
            logging.error("Dropping image after repeated failures. event=%s attempts=%s", pending.event.event_id, pending.attempts)

    def start(self) -> None:
        try:
            text, image_events = self.fetch_scope_text_and_image_events()
        except MarkerMissingError as exc:
            logging.warning("%s", exc)
            return
        except Exception as exc:
            logging.warning("读取飞书文档失败：%s", exc)
            return
        if self.image_settings.enabled:
            try:
                cleanup_old_downloaded_images(self.image_settings)
            except Exception as exc:
                logging.debug("Image temp cleanup failed: %s", exc)
        with self.lock:
            self.last_seen = text
            self.pending_text = ""
            self.pending_since = 0.0
            self.last_delta_at = 0.0
            self.target_window = get_foreground_window()
            self.active = True
            if self.image_settings.enabled:
                if image_events is not None:
                    self._set_image_baseline(image_events)
                else:
                    self.image_baseline_ready = False
                    self.pending_images = []
            logging.info(
                "已开始监听。baseline_len=%s hash=%s target=%s",
                len(text),
                sha12(text),
                self.target_window.summary if self.target_window else "<none>",
            )

    def stop(self, clear_pending: bool = False) -> None:
        with self.lock:
            self.active = False
            if clear_pending:
                self.pending_text = ""
            logging.info("已暂停监听。pending_chars=%s", len(self.pending_text))

    def reset_baseline(self) -> None:
        try:
            text, image_events = self.fetch_scope_text_and_image_events()
        except MarkerMissingError as exc:
            logging.warning("%s", exc)
            return
        except Exception as exc:
            logging.warning("读取飞书文档失败：%s", exc)
            return
        with self.lock:
            self.last_seen = text
            self.pending_text = ""
            self.pending_since = 0.0
            self.last_delta_at = 0.0
            if self.image_settings.enabled:
                if image_events is not None:
                    self._set_image_baseline(image_events)
                else:
                    self.image_baseline_ready = False
                    self.pending_images = []
            logging.info("已重设 baseline。baseline_len=%s hash=%s", len(text), sha12(text))

    def quit(self) -> None:
        with self.lock:
            self.quitting = True
            self.active = False
            logging.info("准备退出。")

    def _strip_delta_leading_newline(self, delta: str, now: float) -> tuple[str, bool, float]:
        if not delta.startswith("\n"):
            return delta, False, 0.0

        policy = self.settings.leading_newline_policy
        idle_seconds = float("inf") if self.last_delta_at <= 0 else max(0.0, now - self.last_delta_at)
        should_strip = policy == "strip" or (
            policy == "smart"
            and not self.pending_text
            and idle_seconds >= self.settings.leading_newline_idle_threshold_seconds
        )
        if not should_strip:
            return delta, False, idle_seconds
        return delta[1:], True, idle_seconds

    def _append_delta_or_reset(self, current: str, now: float) -> None:
        previous = self.last_seen
        if current == previous:
            return

        if current.startswith(previous):
            delta = current[len(previous):]
            delta, stripped, idle_seconds = self._strip_delta_leading_newline(delta, now)
            self.last_seen = current
            if len(delta) >= self.settings.min_chars_to_paste:
                self.pending_text += delta
                self.pending_since = now
                self.last_delta_at = now
                logging.info(
                    "检测到新增文本。chars=%s pending_chars=%s stripped_leading_newline=%s idle=%.2fs hash=%s",
                    len(delta),
                    len(self.pending_text),
                    stripped,
                    idle_seconds,
                    sha12(current),
                )
            elif stripped:
                logging.info("已去掉飞书段落开头换行；本次没有剩余待粘贴文本。idle=%.2fs hash=%s", idle_seconds, sha12(current))
            return

        lcp = common_prefix_len(previous, current)
        changed_tail = previous[lcp:]
        if (
            len(current) > len(previous)
            and len(changed_tail) <= self.settings.append_tolerance_chars
            and changed_tail.strip() == ""
        ):
            delta = current[lcp:].lstrip()
            delta, stripped, idle_seconds = self._strip_delta_leading_newline(delta, now)
            self.last_seen = current
            if len(delta) >= self.settings.min_chars_to_paste:
                self.pending_text += delta
                self.pending_since = now
                self.last_delta_at = now
                logging.info(
                    "检测到尾部空白调整，按追加处理。chars=%s pending_chars=%s stripped_leading_newline=%s idle=%.2fs old_len=%s new_len=%s lcp=%s hash=%s",
                    len(delta),
                    len(self.pending_text),
                    stripped,
                    idle_seconds,
                    len(previous),
                    len(current),
                    lcp,
                    sha12(current),
                )
            elif stripped:
                logging.info("尾部空白调整后已去掉飞书段落开头换行；本次没有剩余待粘贴文本。idle=%.2fs hash=%s", idle_seconds, sha12(current))
            return

        near_tail_delta = find_near_tail_rewrite_append_delta(previous, current, lcp)
        if near_tail_delta is not None:
            delta = near_tail_delta.delta
            delta, stripped, idle_seconds = self._strip_delta_leading_newline(delta, now)
            self.last_seen = current
            if len(delta) >= self.settings.min_chars_to_paste:
                self.pending_text += delta
                self.pending_since = now
                self.last_delta_at = now
                logging.info(
                    "检测到旧尾部短改写，按追加处理。chars=%s pending_chars=%s stripped_leading_newline=%s idle=%.2fs old_len=%s new_len=%s lcp=%s rewritten_tail=%s hash=%s",
                    len(delta),
                    len(self.pending_text),
                    stripped,
                    idle_seconds,
                    len(previous),
                    len(current),
                    lcp,
                    near_tail_delta.rewritten_tail_len,
                    sha12(current),
                )
            elif stripped:
                logging.info(
                    "旧尾部短改写后已去掉飞书段落开头换行；本次没有剩余待粘贴文本。idle=%.2fs old_len=%s new_len=%s lcp=%s rewritten_tail=%s hash=%s",
                    idle_seconds,
                    len(previous),
                    len(current),
                    lcp,
                    near_tail_delta.rewritten_tail_len,
                    sha12(current),
                )
            return

        anchored_delta = find_tail_anchor_append_delta(previous, current)
        if anchored_delta is not None:
            delta = anchored_delta.delta
            delta, stripped, idle_seconds = self._strip_delta_leading_newline(delta, now)
            self.last_seen = current
            if len(delta) >= self.settings.min_chars_to_paste:
                self.pending_text += delta
                self.pending_since = now
                self.last_delta_at = now
                logging.info(
                    "检测到尾部锚点追加，按新增处理。chars=%s pending_chars=%s stripped_leading_newline=%s idle=%.2fs old_len=%s new_len=%s lcp=%s anchor_len=%s drift=%s hash=%s",
                    len(delta),
                    len(self.pending_text),
                    stripped,
                    idle_seconds,
                    len(previous),
                    len(current),
                    lcp,
                    anchored_delta.anchor_len,
                    anchored_delta.offset_drift,
                    sha12(current),
                )
            elif stripped:
                logging.info(
                    "尾部锚点追加后已去掉飞书段落开头换行；本次没有剩余待粘贴文本。idle=%.2fs old_len=%s new_len=%s lcp=%s anchor_len=%s drift=%s hash=%s",
                    idle_seconds,
                    len(previous),
                    len(current),
                    lcp,
                    anchored_delta.anchor_len,
                    anchored_delta.offset_drift,
                    sha12(current),
                )
            return

        context_delta = find_tail_context_append_delta(previous, current)
        if context_delta is not None:
            delta = context_delta.delta
            delta, stripped, idle_seconds = self._strip_delta_leading_newline(delta, now)
            self.last_seen = current
            if len(delta) >= self.settings.min_chars_to_paste:
                self.pending_text += delta
                self.pending_since = now
                self.last_delta_at = now
                logging.info(
                    "检测到近尾部上下文追加，按新增处理。chars=%s pending_chars=%s stripped_leading_newline=%s idle=%.2fs old_len=%s new_len=%s lcp=%s anchor_len=%s drift=%s hash=%s",
                    len(delta),
                    len(self.pending_text),
                    stripped,
                    idle_seconds,
                    len(previous),
                    len(current),
                    lcp,
                    context_delta.anchor_len,
                    context_delta.offset_drift,
                    sha12(current),
                )
            elif stripped:
                logging.info(
                    "近尾部上下文追加后已去掉飞书段落开头换行；本次没有剩余待粘贴文本。idle=%.2fs old_len=%s new_len=%s lcp=%s anchor_len=%s drift=%s hash=%s",
                    idle_seconds,
                    len(previous),
                    len(current),
                    lcp,
                    context_delta.anchor_len,
                    context_delta.offset_drift,
                    sha12(current),
                )
            return

        if len(current) < len(previous):
            logging.info("文档内容变短，已重设 baseline，不粘贴。old_len=%s new_len=%s hash=%s", len(previous), len(current), sha12(current))
        else:
            logging.warning("检测到非追加式修改，已重设 baseline，不粘贴。old_len=%s new_len=%s lcp=%s hash=%s", len(previous), len(current), lcp, sha12(current))
        self.last_seen = current
        self.pending_text = ""
        self.pending_since = 0.0
        self.last_delta_at = 0.0

    def _target_is_active(self) -> bool:
        if self.settings.target_window_mode == "any" or not self.settings.require_same_foreground_window:
            return True
        if self.target_window is None:
            logging.warning("未捕获目标窗口，暂停粘贴。")
            return False

        current = get_foreground_window()
        if self.settings.target_window_mode == "process" and current.process_name == self.target_window.process_name:
            return True

        if current.hwnd == self.target_window.hwnd:
            if (
                self.settings.require_same_focused_control
                and self.target_window.focus_hwnd
                and current.focus_hwnd
                and current.focus_hwnd != self.target_window.focus_hwnd
            ):
                logging.warning(
                    "目标窗口内焦点控件已变化，暂停粘贴。target_focus=%s current_focus=%s target_title=%r",
                    self.target_window.focus_hwnd,
                    current.focus_hwnd,
                    self.target_window.title,
                )
                return False
            return True

        if self.settings.allow_refocus_target_window and refocus_window(self.target_window.hwnd):
            logging.info("已尝试切回目标窗口。target=%s", self.target_window.summary)
            return True

        logging.warning(
            "当前前台窗口已变化，暂停粘贴并保留 pending。target=%s current=%s",
            self.target_window.summary,
            current.summary,
        )
        return False

    def tick(self) -> None:
        with self.lock:
            if not self.active:
                return

        try:
            current, image_events = self.fetch_scope_text_and_image_events()
        except MarkerMissingError as exc:
            logging.warning("%s", exc)
            return
        except Exception as exc:
            logging.warning("读取飞书文档失败：%s", exc)
            return

        now = time.time()
        text_to_paste = ""
        image_to_paste: Optional[PendingImage] = None
        with self.lock:
            self._append_delta_or_reset(current, now)
            if image_events is not None:
                self._append_image_events_or_reset(image_events, now)

            if (
                self.pending_text
                and len(self.pending_text) >= self.settings.min_chars_to_paste
                and now - self.pending_since >= self.settings.stable_delay_ms / 1000.0
                and self._target_is_active()
            ):
                text_to_paste = self.pending_text
                self.pending_text = ""

            image_to_paste = self._pop_ready_image(now)

        if text_to_paste:
            try:
                logging.info("准备粘贴新增文本。chars=%s", len(text_to_paste))
                self.inserter.paste_text(text_to_paste)
                logging.info("粘贴完成。chars=%s", len(text_to_paste))
            except Exception as exc:
                logging.error("粘贴失败：%s", exc)
                with self.lock:
                    self.pending_text = text_to_paste + self.pending_text
                    self.pending_since = time.time()

        if image_to_paste is not None:
            try:
                downloaded = download_image_with_lark_cli(self.config, image_to_paste.event, self.image_settings)
                logging.info(
                    "Prepared image paste. event=%s occurrence=%s size_bytes=%s kind=%s width=%s height=%s sha256=%s mode=%s",
                    image_to_paste.event.event_id,
                    image_to_paste.event.occurrence_index,
                    downloaded.size_bytes,
                    downloaded.kind,
                    downloaded.width,
                    downloaded.height,
                    downloaded.sha256_8,
                    self.image_settings.insert_mode,
                )
                self.image_inserter.paste_image(downloaded, self.image_settings)
                logging.info("Image paste completed. event=%s", image_to_paste.event.event_id)
            except Exception as exc:
                logging.error("Image paste failed. event=%s error=%s", image_to_paste.event.event_id, exc)
                with self.lock:
                    self._handle_image_paste_failure(image_to_paste)

    def run_loop(self) -> None:
        logging.info(
            "轮询已启动。热键：F8 开始，F9 暂停，F10 重设 baseline，F12 退出。"
        )
        while True:
            with self.lock:
                if self.quitting:
                    break
            self.tick()
            time.sleep(self.settings.poll_interval_seconds)


def register_hotkeys(bridge: Bridge, hotkeys: dict[str, Any]) -> bool:
    if keyboard is None:
        logging.warning("keyboard 模块不可用，改用控制台命令模式。")
        return False
    try:
        keyboard.add_hotkey(str(hotkeys.get("start", "F8")), bridge.start)
        keyboard.add_hotkey(str(hotkeys.get("stop", "F9")), bridge.stop)
        keyboard.add_hotkey(str(hotkeys.get("reset_baseline", "F10")), bridge.reset_baseline)
        keyboard.add_hotkey(str(hotkeys.get("quit", "F12")), bridge.quit)
        return True
    except Exception as exc:
        logging.warning("注册全局热键失败，改用控制台命令模式：%s", exc)
        return False


def start_console_command_thread(bridge: Bridge) -> None:
    def worker() -> None:
        logging.info("控制台命令：输入 start/stop/reset/quit 后回车。")
        while True:
            try:
                command = input("> ").strip().lower()
            except EOFError:
                return
            if command in {"start", "s"}:
                bridge.start()
            elif command in {"stop", "pause", "p"}:
                bridge.stop()
            elif command in {"reset", "r"}:
                bridge.reset_baseline()
            elif command in {"quit", "exit", "q"}:
                bridge.quit()
                return
            elif command:
                logging.info("未知命令：%s。可用命令：start/stop/reset/quit", command)

    threading.Thread(target=worker, daemon=True).start()


def build_bridge_config(config: dict[str, Any]) -> dict[str, Any]:
    if "feishu_app_id" in config or "feishu_app_secret" in config:
        logging.warning("配置文件中的 feishu_app_id/feishu_app_secret 会被忽略；请使用环境变量。")
    return config


def print_doc_id(config: dict[str, Any], override_doc_url: str = "") -> None:
    source = override_doc_url or str(config.get("document_id") or config.get("doc_url") or "")
    print(extract_docx_id(source))


def check_document(bridge: Bridge) -> None:
    text = bridge.fetch_scope_text()
    logging.info(
        "读取成功。document_id=%s len=%s hash=%s。不会打印文档全文。",
        bridge.document_id,
        len(text),
        sha12(text),
    )


def dry_run_clear_body(config: dict[str, Any]) -> int:
    bridge = Bridge(build_bridge_config(config))
    count = bridge.count_document_body_blocks()
    logging.info("dry-run: 将删除 %s 个正文块；页面标题会保留。", count)
    logging.info("本工具不冻结飞书页面标题；如果有人有文档编辑权限，仍可手动改标题。")
    return 0


def clear_body(config: dict[str, Any]) -> int:
    bridge = Bridge(build_bridge_config(config))
    before = bridge.count_document_body_blocks()
    if before <= 0:
        logging.info("文档正文已为空；页面标题保持不变。")
        return 0
    deleted = bridge.clear_document_body()
    after = bridge.count_document_body_blocks()
    logging.info("已清空文档正文：删除 %s 个正文块，剩余 %s 个；页面标题保持不变。", deleted, after)
    logging.info("本工具不冻结飞书页面标题；如果有人有文档编辑权限，仍可手动改标题。")
    return 0


def run_watch(config: dict[str, Any]) -> int:
    bridge = Bridge(build_bridge_config(config))
    hotkeys_ok = register_hotkeys(bridge, config.get("hotkeys", {}) if isinstance(config.get("hotkeys"), dict) else {})
    if not hotkeys_ok:
        start_console_command_thread(bridge)
    if bool(config.get("auto_begin_on_watch", True)):
        try:
            bridge.start()
        except Exception as exc:
            logging.warning("auto_begin_on_watch failed: %s", exc)
    try:
        bridge.run_loop()
    except KeyboardInterrupt:
        logging.info("收到 Ctrl+C，退出。")
    return 0


def run_test_output(config: dict[str, Any]) -> int:
    settings = BridgeSettings.from_config(config)
    target = get_foreground_window()
    logging.info("将向当前前台窗口发送固定测试文本。target=%s chars=%s", target.summary, len(DEFAULT_TEST_TEXT))
    TextInserter(settings.insert_mode, settings.restore_clipboard).paste_text(DEFAULT_TEST_TEXT)
    logging.info("测试文本已发送。")
    return 0


def image_event_summary(event: ImageEvent) -> dict[str, Any]:
    return {
        "occurrence_index": event.occurrence_index,
        "event_id": event.event_id,
        "token_masked": event.token_masked,
        "width": event.width,
        "height": event.height,
        "tag_position": event.tag_position,
        "has_stable_id": bool(event.stable_id),
    }


def downloaded_image_summary(image: DownloadedImage) -> dict[str, Any]:
    return {
        "path": str(image.path),
        "size_bytes": image.size_bytes,
        "kind": image.kind,
        "width": image.width,
        "height": image.height,
        "sha256_8": image.sha256_8,
    }


def run_dump_image_tokens(config: dict[str, Any]) -> int:
    bridge = Bridge(build_bridge_config(config))
    events = bridge.fetch_image_events()
    payload = {"image_count": len(events), "images": [image_event_summary(event) for event in events]}
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def run_download_latest_image(config: dict[str, Any]) -> int:
    bridge = Bridge(build_bridge_config(config))
    events = bridge.fetch_image_events()
    if not events:
        raise RuntimeError("No image tags found in the configured document.")
    event = events[-1]
    image = download_image_with_lark_cli(config, event, bridge.image_settings)
    payload = {"image": image_event_summary(event), "download": downloaded_image_summary(image)}
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def run_test_image_output(config: dict[str, Any], file_path: str = "", mode: str = "") -> int:
    settings = ImageSettings.from_config(config)
    if mode:
        image_config = dict(config.get("image", {}) if isinstance(config.get("image"), dict) else {})
        image_config["insert_mode"] = mode
        settings = ImageSettings.from_config({**config, "image": image_config})
    temp_dir = resolve_image_temp_dir(settings.temp_dir)
    if file_path:
        image_path = Path(file_path).expanduser().resolve()
        if not image_path.exists():
            raise FileNotFoundError(str(image_path))
    else:
        image_path = temp_dir / "doubao_image_local_test.png"
        write_test_png(image_path)
    digest = sha256_file(image_path)
    kind, width, height = image_metadata(image_path)
    image = DownloadedImage(
        path=image_path,
        size_bytes=image_path.stat().st_size,
        sha256_8=digest[:8],
        kind=kind,
        width=width,
        height=height,
    )
    logging.info("Setting local test image clipboard. size_bytes=%s sha256=%s mode=%s", image.size_bytes, image.sha256_8, settings.insert_mode)
    ImageInserter().paste_image(image, settings)
    logging.info("Local test image output completed.")
    return 0


def run_debug_image_pipeline(config: dict[str, Any]) -> int:
    bridge = Bridge(build_bridge_config(config))
    if not isinstance(bridge.client, LarkCliClient):
        raise RuntimeError("--debug-image-pipeline requires auth_mode=lark_cli.")
    content = bridge.client.get_document_content()
    raw = bridge.client.get_raw_content()
    text_without_media = document_content_to_plain_text_without_media(content)
    events = parse_image_events_from_content(content)
    payload = {
        "image_count": len(events),
        "images": [image_event_summary(event) for event in events],
        "fetch_content_len": len(content),
        "raw_content_len": len(raw),
        "text_without_media_len": len(text_without_media),
        "text_without_media_hash": sha12(text_without_media),
        "jpg_like_in_fetch_content": bool(re.search(r"\.(?:jpg|jpeg|png)\b", content, re.IGNORECASE)),
        "jpg_like_in_raw_content": bool(re.search(r"\.(?:jpg|jpeg|png)\b", raw, re.IGNORECASE)),
        "jpg_like_in_text_without_media": bool(re.search(r"\.(?:jpg|jpeg|png)\b", text_without_media, re.IGNORECASE)),
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def run_inspect_clipboard() -> int:
    formats = list_clipboard_formats()
    payload = {
        "formats": formats,
        "has_image": clipboard_has_image_format(formats),
        "has_text": clipboard_has_text_format(formats),
        "has_file_drop": clipboard_has_file_drop_format(formats),
    }
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


def run_test_image_bridge_once(config: dict[str, Any]) -> int:
    bridge = Bridge(build_bridge_config(config))
    events = bridge.fetch_image_events()
    if not events:
        raise RuntimeError("No image tags found in the configured document.")
    event = events[-1]
    image = download_image_with_lark_cli(config, event, bridge.image_settings)
    logging.info(
        "Testing latest image bridge once. event=%s occurrence=%s size_bytes=%s kind=%s width=%s height=%s sha256=%s mode=%s",
        event.event_id,
        event.occurrence_index,
        image.size_bytes,
        image.kind,
        image.width,
        image.height,
        image.sha256_8,
        bridge.image_settings.insert_mode,
    )
    bridge.image_inserter.paste_image(image, bridge.image_settings)
    return 0


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="飞书云文档语音桥接工具")
    parser.add_argument("--config", default="config.json", help="配置文件路径，默认 config.json")
    parser.add_argument("--doc-url", default="", help="仅用于 --print-doc-id 时覆盖配置中的 doc_url/document_id")
    parser.add_argument("--check", action="store_true", help="检查配置、鉴权和文档读取，只输出长度和 hash")
    parser.add_argument("--print-doc-id", action="store_true", help="从 doc_url 解析 document_id 并打印")
    parser.add_argument("--test-output", action="store_true", help="不访问飞书，把固定测试文本粘贴到当前光标处")
    parser.add_argument("--dump-image-tokens", action="store_true", help="读取文档并只输出图片摘要，不输出全文")
    parser.add_argument("--download-latest-image", action="store_true", help="下载当前文档中最新图片到 image.temp_dir，不粘贴")
    parser.add_argument("--test-image-output", action="store_true", help="不访问飞书，生成或读取本地测试图片并粘贴到当前光标处")
    parser.add_argument("--file", default="", help="仅用于 --test-image-output，指定要测试粘贴的图片文件")
    parser.add_argument("--mode", default="", choices=("", "auto", "clipboard_bitmap", "clipboard_file", "save_only"), help="仅用于 --test-image-output，覆盖 image.insert_mode")
    parser.add_argument("--test-image-bridge-once", action="store_true", help="读取最新图片、下载并尝试粘贴一次")
    parser.add_argument("--debug-image-pipeline", action="store_true", help="读取文档并输出图片事件与去媒体文字摘要")
    parser.add_argument("--inspect-clipboard", action="store_true", help="只枚举当前 Windows 剪贴板格式，不输出内容")
    parser.add_argument("--once", action="store_true", help="读取一次文档，只显示长度和 hash，不粘贴")
    parser.add_argument("--watch", action="store_true", help="启动常驻轮询和热键")
    parser.add_argument("--clear-body", action="store_true", help="删除文档根正文块，保留飞书页面标题")
    parser.add_argument("--clear-body-dry-run", action="store_true", help="只统计将删除多少个正文块，不真正删除")
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    configure_stdio()
    set_dpi_aware()
    setup_logging("INFO")
    arg_list = sys.argv[1:] if argv is None else argv
    double_click_mode = len(arg_list) == 0
    args = parse_args(arg_list)

    config_required = not (args.test_output or args.test_image_output or args.inspect_clipboard or (args.print_doc_id and args.doc_url))
    config_path = resolve_config_path(args.config)
    created_config = False

    try:
        if double_click_mode:
            created_config = ensure_config_for_double_click(config_path)

        config = load_json(config_path, required=config_required)
        setup_logging(str(config.get("log_level", "INFO")))

        if double_click_mode:
            reasons = config_not_ready_reasons(config)
            print_double_click_help(config_path, created_config, reasons)
            if reasons:
                pause_before_exit()
                return 1

        if args.print_doc_id:
            print_doc_id(config, args.doc_url)
            return 0

        if args.test_output:
            return run_test_output(config)

        if args.test_image_output:
            return run_test_image_output(config, args.file, args.mode)

        if args.inspect_clipboard:
            return run_inspect_clipboard()

        if args.debug_image_pipeline:
            return run_debug_image_pipeline(config)

        if args.dump_image_tokens:
            return run_dump_image_tokens(config)

        if args.download_latest_image:
            return run_download_latest_image(config)

        if args.test_image_bridge_once:
            return run_test_image_bridge_once(config)

        if args.clear_body_dry_run:
            return dry_run_clear_body(config)

        if args.clear_body:
            return clear_body(config)

        if args.check or args.once:
            bridge = Bridge(build_bridge_config(config))
            check_document(bridge)
            return 0

        return run_watch(config)
    except Exception as exc:
        logging.error("%s", exc)
        if double_click_mode:
            pause_before_exit()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
