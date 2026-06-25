#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Feishu Voice Bridge

在 Windows 上轮询一个指定的飞书新版云文档 docx，把 F8 后新增的纯文本
粘贴到用户启动时捕获的目标输入窗口。它不做语音识别，不接豆包/火山/OCR。
"""

from __future__ import annotations

import argparse
import ctypes
from ctypes import wintypes
import hashlib
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid
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
    "target_window_mode": "locked",
    "require_same_foreground_window": True,
    "require_same_focused_control": False,
    "allow_refocus_target_window": False,
    "log_level": "INFO",
    "auto_start_watch": True,
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
    with path.open("r", encoding="utf-8") as f:
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


class Bridge:
    def __init__(self, config: dict[str, Any]):
        auth_mode = get_auth_mode(config)
        document_id = extract_docx_id(str(config.get("document_id") or config.get("doc_url") or ""))
        self.settings = BridgeSettings.from_config(config)
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
        self.inserter = TextInserter(self.settings.insert_mode, self.settings.restore_clipboard)

        self.document_id = document_id
        self.active = False
        self.quitting = False
        self.last_seen = ""
        self.pending_text = ""
        self.pending_since = 0.0
        self.last_delta_at = 0.0
        self.target_window: Optional[TargetWindow] = None
        self.lock = threading.RLock()

    def fetch_scope_text(self) -> str:
        raw = self.client.get_raw_content()
        text = extract_scope_text(raw, self.settings.monitor_after_marker)
        if self.settings.trim_trailing_whitespace_for_delta:
            return text.rstrip()
        return text

    def count_document_body_blocks(self) -> int:
        return self.client.count_document_body_blocks()

    def clear_document_body(self) -> int:
        return self.client.clear_document_body()

    def start(self) -> None:
        try:
            text = self.fetch_scope_text()
        except MarkerMissingError as exc:
            logging.warning("%s", exc)
            return
        with self.lock:
            self.last_seen = text
            self.pending_text = ""
            self.pending_since = 0.0
            self.last_delta_at = 0.0
            self.target_window = get_foreground_window()
            self.active = True
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
            text = self.fetch_scope_text()
        except MarkerMissingError as exc:
            logging.warning("%s", exc)
            return
        with self.lock:
            self.last_seen = text
            self.pending_text = ""
            self.pending_since = 0.0
            self.last_delta_at = 0.0
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
            current = self.fetch_scope_text()
        except MarkerMissingError as exc:
            logging.warning("%s", exc)
            return
        except Exception as exc:
            logging.warning("读取飞书文档失败：%s", exc)
            return

        now = time.time()
        text_to_paste = ""
        with self.lock:
            self._append_delta_or_reset(current, now)
            if not self.pending_text:
                return
            if len(self.pending_text) < self.settings.min_chars_to_paste:
                return
            if now - self.pending_since < self.settings.stable_delay_ms / 1000.0:
                return
            if not self._target_is_active():
                return
            text_to_paste = self.pending_text
            self.pending_text = ""

        try:
            logging.info("准备粘贴新增文本。chars=%s", len(text_to_paste))
            self.inserter.paste_text(text_to_paste)
            logging.info("粘贴完成。chars=%s", len(text_to_paste))
        except Exception as exc:
            logging.error("粘贴失败：%s", exc)
            with self.lock:
                self.pending_text = text_to_paste + self.pending_text
                self.pending_since = time.time()

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


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="飞书云文档语音桥接工具")
    parser.add_argument("--config", default="config.json", help="配置文件路径，默认 config.json")
    parser.add_argument("--doc-url", default="", help="仅用于 --print-doc-id 时覆盖配置中的 doc_url/document_id")
    parser.add_argument("--check", action="store_true", help="检查配置、鉴权和文档读取，只输出长度和 hash")
    parser.add_argument("--print-doc-id", action="store_true", help="从 doc_url 解析 document_id 并打印")
    parser.add_argument("--test-output", action="store_true", help="不访问飞书，把固定测试文本粘贴到当前光标处")
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

    config_required = not (args.test_output or (args.print_doc_id and args.doc_url))
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
