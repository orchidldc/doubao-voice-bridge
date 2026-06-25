#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
DouBao Voice Bridge GUI.

A focused Windows GUI wrapper around the local Feishu/Lark document bridge.
The GUI stores config, starts/stops the bridge subprocess, and streams logs.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import BOTH, DISABLED, END, LEFT, NORMAL, RIGHT, VERTICAL, W, X, Y, BooleanVar, StringVar, Tk, messagebox
from tkinter import ttk
import tkinter as tk


APP_NAME = "DouBao Voice Bridge"
APP_VERSION = "0.2.1"
CREATE_NO_WINDOW = 0x08000000
APP_ICON_RELATIVE_PATH = "assets/doubao_d.ico"
BRIDGE_PROCESS_IMAGES = ("doubao_voice_bridge_cli.exe", "feishu_voice_bridge.exe")


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_resource_path(relative_path: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent / relative_path


def get_icon_path() -> Path:
    candidates = [
        APP_DIR / APP_ICON_RELATIVE_PATH,
        get_resource_path(APP_ICON_RELATIVE_PATH),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


APP_DIR = get_app_dir()
LOG_DIR = APP_DIR / "logs"
APP_LOG_PATH = LOG_DIR / "doubao_voice_bridge_gui.log"
BRIDGE_LOG_PATH = LOG_DIR / "doubao_voice_bridge.log"
CONFIG_PATH = APP_DIR / "doubao_voice_bridge_config.json"
LEGACY_CONFIG_PATHS = [
    APP_DIR / "config.json",
    Path.home()
    / "Documents"
    / "Codex"
    / "2026-06-23"
    / "files-mentioned-by-the-user-feishu"
    / "outputs"
    / "config.json",
]


class GuiLogger:
    def __init__(self) -> None:
        LOG_DIR.mkdir(parents=True, exist_ok=True)

    def emit(self, message: str, level: str = "INFO") -> None:
        line = f"{datetime.now():%Y-%m-%d %H:%M:%S} [{level}] {message}\n"
        APP_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with APP_LOG_PATH.open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(line)

    def emit_bridge(self, message: str, level: str = "INFO") -> None:
        line = f"{datetime.now():%Y-%m-%d %H:%M:%S} [{level}] {message}\n"
        BRIDGE_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with BRIDGE_LOG_PATH.open("a", encoding="utf-8", newline="\n") as handle:
            handle.write(line)


class DouBaoVoiceBridgeApp(ttk.Frame):
    def __init__(self, master: Tk, logger: GuiLogger) -> None:
        super().__init__(master, padding=18)
        self.master = master
        self.logger = logger
        self.process: subprocess.Popen | None = None
        self.reader: threading.Thread | None = None
        self._config_cache: dict[str, object] = {}

        self.status_var = StringVar(value="未运行")
        self.last_action_var = StringVar(value="尚未启动")
        self.config_path_var = StringVar(value=str(CONFIG_PATH))
        self.doc_url_var = StringVar()
        self.auth_mode_var = StringVar(value="lark_cli")
        self.lark_cli_path_var = StringVar(value="lark-cli")
        self.lark_cli_profile_var = StringVar(value="")
        self.target_window_mode_var = StringVar(value="any")
        self.poll_interval_var = StringVar(value="1.2")
        self.stable_delay_var = StringVar(value="1000")
        self.leading_newline_threshold_var = StringVar(value="2.0")
        self.marker_var = StringVar(value="")
        self.auto_start_var = BooleanVar(value=True)

        self.output_text: tk.Text | None = None
        self.start_button: ttk.Button | None = None
        self.stop_button: ttk.Button | None = None

        self._build_ui()
        self.load_config_into_ui()

    def _build_ui(self) -> None:
        self.pack(fill=BOTH, expand=True)

        header = ttk.Frame(self)
        header.pack(fill=X)
        ttk.Label(header, text=APP_NAME, font=("Microsoft YaHei UI", 22, "bold")).pack(anchor=W)
        ttk.Label(
            header,
            text="豆包跨屏输入桥接工具：手机端用豆包/输入法把语音转成文字写入飞书文档，电脑端把新增文字粘贴到当前输入框。",
            foreground="#555555",
            font=("Microsoft YaHei UI", 11),
            wraplength=840,
        ).pack(anchor=W, pady=(6, 16))

        status = ttk.LabelFrame(self, text="状态", padding=16)
        status.pack(fill=X, pady=(0, 14))
        ttk.Label(status, text="运行状态:").grid(row=0, column=0, sticky=W, padx=(0, 12), pady=6)
        ttk.Label(status, textvariable=self.status_var, font=("Microsoft YaHei UI", 12, "bold")).grid(row=0, column=1, sticky=W, pady=6)
        ttk.Label(status, text="最近动作:").grid(row=1, column=0, sticky=W, padx=(0, 12), pady=6)
        ttk.Label(status, textvariable=self.last_action_var, wraplength=760).grid(row=1, column=1, sticky=W, pady=6)
        ttk.Label(status, text="配置文件:").grid(row=2, column=0, sticky=W, padx=(0, 12), pady=6)
        ttk.Label(status, textvariable=self.config_path_var, wraplength=760, foreground="#444444").grid(row=2, column=1, sticky=W, pady=6)

        config = ttk.LabelFrame(self, text="桥接配置", padding=16)
        config.pack(fill=X, pady=(0, 14))
        ttk.Label(config, text="飞书 docx 链接:").grid(row=0, column=0, sticky=W, padx=(0, 10), pady=7)
        ttk.Entry(config, textvariable=self.doc_url_var, width=66, font=("Microsoft YaHei UI", 10)).grid(
            row=0, column=1, columnspan=5, sticky=W, pady=7
        )

        ttk.Label(config, text="鉴权方式:").grid(row=1, column=0, sticky=W, padx=(0, 10), pady=7)
        ttk.Combobox(config, textvariable=self.auth_mode_var, values=("lark_cli", "app"), width=12, state="readonly").grid(
            row=1, column=1, sticky=W, pady=7
        )
        ttk.Label(config, text="lark-cli:").grid(row=1, column=2, sticky=W, padx=(20, 10), pady=7)
        ttk.Entry(config, textvariable=self.lark_cli_path_var, width=18).grid(row=1, column=3, sticky=W, pady=7)
        ttk.Label(config, text="profile:").grid(row=1, column=4, sticky=W, padx=(20, 10), pady=7)
        ttk.Entry(config, textvariable=self.lark_cli_profile_var, width=18).grid(row=1, column=5, sticky=W, pady=7)

        ttk.Label(config, text="目标窗口模式:").grid(row=2, column=0, sticky=W, padx=(0, 10), pady=7)
        ttk.Combobox(config, textvariable=self.target_window_mode_var, values=("any", "locked", "process"), width=12, state="readonly").grid(
            row=2, column=1, sticky=W, pady=7
        )
        ttk.Label(config, text="轮询秒:").grid(row=2, column=2, sticky=W, padx=(20, 10), pady=7)
        ttk.Entry(config, textvariable=self.poll_interval_var, width=10).grid(row=2, column=3, sticky=W, pady=7)
        ttk.Label(config, text="稳定毫秒:").grid(row=2, column=4, sticky=W, padx=(20, 10), pady=7)
        ttk.Entry(config, textvariable=self.stable_delay_var, width=10).grid(row=2, column=5, sticky=W, pady=7)

        ttk.Label(config, text="只监听 marker 后内容:").grid(row=3, column=0, sticky=W, padx=(0, 10), pady=7)
        ttk.Entry(config, textvariable=self.marker_var, width=48).grid(row=3, column=1, columnspan=4, sticky=W, pady=7)
        ttk.Label(config, text="开头换行阈值秒:").grid(row=4, column=0, sticky=W, padx=(0, 10), pady=7)
        ttk.Entry(config, textvariable=self.leading_newline_threshold_var, width=10).grid(row=4, column=1, sticky=W, pady=7)
        ttk.Label(
            config,
            text="超过该间隔的新语音片段会去掉飞书自动段落开头换行；短时间连续输入仍保留换行。",
            foreground="#555555",
            wraplength=500,
        ).grid(row=4, column=2, columnspan=4, sticky=W, padx=(20, 0), pady=7)
        ttk.Checkbutton(config, text="打开软件后自动启动桥接监听", variable=self.auto_start_var).grid(
            row=5, column=0, columnspan=3, sticky=W, pady=(8, 0)
        )
        ttk.Label(
            config,
            text="热键：F8 开始并建立 baseline，F9 暂停，F10 重设 baseline，F12 退出桥接进程。",
            foreground="#444444",
            wraplength=820,
        ).grid(row=6, column=0, columnspan=6, sticky=W, pady=(10, 0))

        controls = ttk.Frame(self)
        controls.pack(fill=X, pady=(0, 14))
        ttk.Button(controls, text="保存配置", command=self.save_config).pack(side=LEFT, padx=(0, 8))
        ttk.Button(controls, text="检查连接", command=self.run_check).pack(side=LEFT, padx=(0, 8))
        ttk.Button(controls, text="单次读取", command=self.run_once).pack(side=LEFT, padx=(0, 8))
        ttk.Button(controls, text="测试粘贴", command=self.run_test_output).pack(side=LEFT, padx=(0, 8))
        ttk.Button(controls, text="一键清理文档", command=self.run_clear_body).pack(side=LEFT, padx=(0, 8))
        self.start_button = ttk.Button(controls, text="启动监听", command=self.start_watch)
        self.start_button.pack(side=LEFT, padx=(0, 8))
        self.stop_button = ttk.Button(controls, text="停止监听", command=self.stop_watch, state=DISABLED)
        self.stop_button.pack(side=LEFT, padx=(0, 8))
        ttk.Button(controls, text="打开配置", command=self.open_config).pack(side=RIGHT, padx=(8, 0))
        ttk.Button(controls, text="打开日志目录", command=self.open_logs_dir).pack(side=RIGHT)

        output_frame = ttk.LabelFrame(self, text="运行输出", padding=8)
        output_frame.pack(fill=BOTH, expand=True)
        self.output_text = tk.Text(output_frame, height=18, wrap="word", state=DISABLED, font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(output_frame, orient=VERTICAL, command=self.output_text.yview)
        self.output_text.configure(yscrollcommand=scrollbar.set)
        self.output_text.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

    def start_after_ui_ready(self) -> None:
        self._cleanup_orphan_bridge_processes()
        if self.auto_start_var.get():
            self.start_watch()

    def _bridge_exe_path(self) -> Path:
        candidates = [
            APP_DIR / "tools" / "doubao_voice_bridge_cli.exe",
            get_resource_path("tools/doubao_voice_bridge_cli.exe"),
            APP_DIR / "doubao_voice_bridge_cli.exe",
            APP_DIR / "feishu_voice_bridge.exe",
            get_resource_path("tools/feishu_voice_bridge.exe"),
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return candidates[0]

    def _bridge_script_path(self) -> Path:
        candidates = [
            APP_DIR / "tools" / "feishu_voice_bridge.py",
            get_resource_path("tools/feishu_voice_bridge.py"),
            APP_DIR / "feishu_voice_bridge.py",
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return candidates[0]

    def _example_config_path(self) -> Path:
        candidates = [
            APP_DIR / "tools" / "config.example.json",
            get_resource_path("tools/config.example.json"),
            APP_DIR / "config.example.json",
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return candidates[0]

    @staticmethod
    def _default_config() -> dict[str, object]:
        return {
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
            "target_window_mode": "any",
            "require_same_foreground_window": False,
            "require_same_focused_control": False,
            "allow_refocus_target_window": False,
            "log_level": "INFO",
            "auto_start_watch": True,
            "hotkeys": {"start": "F8", "stop": "F9", "reset_baseline": "F10", "quit": "F12"},
        }

    def _load_config(self) -> dict[str, object]:
        source: Path | None = None
        if CONFIG_PATH.exists():
            source = CONFIG_PATH
        else:
            for candidate in LEGACY_CONFIG_PATHS:
                if candidate.exists():
                    source = candidate
                    break
        if source is None and self._example_config_path().exists():
            source = self._example_config_path()

        config = self._default_config()
        if source is not None:
            try:
                data = json.loads(source.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    config.update(data)
            except Exception as exc:
                self._log(f"读取配置失败，使用默认配置：{exc}", "WARN")
        self._config_cache = dict(config)
        return config

    def load_config_into_ui(self) -> None:
        config = self._load_config()
        self.doc_url_var.set(str(config.get("doc_url", "")))
        self.auth_mode_var.set(str(config.get("auth_mode", "lark_cli")))
        self.lark_cli_path_var.set(str(config.get("lark_cli_path", "lark-cli")))
        self.lark_cli_profile_var.set(str(config.get("lark_cli_profile", "")))
        self.target_window_mode_var.set(str(config.get("target_window_mode", "any")))
        self.poll_interval_var.set(str(config.get("poll_interval_seconds", 1.2)))
        self.stable_delay_var.set(str(config.get("stable_delay_ms", 1000)))
        self.leading_newline_threshold_var.set(str(config.get("leading_newline_idle_threshold_seconds", 2.0)))
        self.marker_var.set(str(config.get("monitor_after_marker", "")))
        self.auto_start_var.set(bool(config.get("auto_start_watch", True)))

    def _config_from_ui(self) -> dict[str, object]:
        config = self._default_config()
        config.update(self._config_cache)
        config["doc_url"] = self.doc_url_var.get().strip()
        config["auth_mode"] = self.auth_mode_var.get().strip() or "lark_cli"
        config["lark_cli_path"] = self.lark_cli_path_var.get().strip() or "lark-cli"
        config["lark_cli_profile"] = self.lark_cli_profile_var.get().strip()
        config["target_window_mode"] = self.target_window_mode_var.get().strip() or "any"
        config["monitor_after_marker"] = self.marker_var.get()
        config["leading_newline_policy"] = "smart"
        config["auto_start_watch"] = bool(self.auto_start_var.get())

        try:
            config["poll_interval_seconds"] = max(0.2, float(self.poll_interval_var.get()))
        except ValueError:
            config["poll_interval_seconds"] = 1.2
            self.poll_interval_var.set("1.2")

        try:
            config["stable_delay_ms"] = max(0, int(float(self.stable_delay_var.get())))
        except ValueError:
            config["stable_delay_ms"] = 1000
            self.stable_delay_var.set("1000")

        try:
            config["leading_newline_idle_threshold_seconds"] = max(0.0, float(self.leading_newline_threshold_var.get()))
        except ValueError:
            config["leading_newline_idle_threshold_seconds"] = 2.0
            self.leading_newline_threshold_var.set("2.0")

        config["hotkeys"] = {"start": "F8", "stop": "F9", "reset_baseline": "F10", "quit": "F12"}
        return config

    def save_config(self) -> None:
        config = self._config_from_ui()
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")
        self._config_cache = dict(config)
        self.config_path_var.set(str(CONFIG_PATH))
        self._set_last_action("配置已保存")
        self._log(f"配置已保存：{CONFIG_PATH}")

    def _base_command(self) -> list[str]:
        exe = self._bridge_exe_path()
        if exe.exists():
            return [str(exe), "--config", str(CONFIG_PATH)]

        script = self._bridge_script_path()
        if script.exists():
            python_runner = sys.executable
            if getattr(sys, "frozen", False):
                python_runner = shutil.which("python") or shutil.which("py") or ""
            if python_runner:
                return [python_runner, str(script), "--config", str(CONFIG_PATH)]

        raise FileNotFoundError(f"桥接工具不存在：{exe}；源代码脚本也不可用：{script}")

    def run_check(self) -> None:
        self.save_config()
        self._run_one_shot("检查连接", ["--check"])

    def run_once(self) -> None:
        self.save_config()
        self._run_one_shot("单次读取", ["--once"])

    def run_test_output(self) -> None:
        self.save_config()
        self._run_one_shot("测试粘贴", ["--test-output"])

    def run_clear_body(self) -> None:
        if self.is_running():
            messagebox.showwarning(APP_NAME, "请先停止监听，再清理文档正文。")
            return
        if not messagebox.askyesno(APP_NAME, "确定清理当前配置文档的正文吗？\n\n会删除正文块，只保留飞书页面标题。"):
            return
        self.save_config()
        self._run_one_shot("清理文档正文", ["--clear-body"])

    def _run_one_shot(self, label: str, extra_args: list[str]) -> None:
        def worker() -> None:
            try:
                command = self._base_command() + extra_args
                self.after(0, lambda: self._set_last_action(f"{label}运行中"))
                self._log(f"{label}开始：{' '.join(command)}")
                process = self._spawn(command)
                self._stream_process_output(process, managed=False)
                code = process.wait()
                level = "INFO" if code == 0 else "WARN"
                self._log(f"{label}结束，退出码 {code}", level)
                self.after(0, lambda: self._set_last_action(f"{label}结束，退出码 {code}"))
            except Exception as exc:
                self._log(f"{label}失败：{exc}", "ERROR")
                self.after(0, lambda: self._set_last_action(f"{label}失败"))

        threading.Thread(target=worker, daemon=True).start()

    def start_watch(self) -> None:
        if self.is_running():
            return
        self.save_config()
        try:
            command = self._base_command() + ["--watch"]
            self.process = self._spawn(command)
        except Exception as exc:
            self._log(f"启动监听失败：{exc}", "ERROR")
            self._set_last_action("启动监听失败")
            return
        self.status_var.set("运行中")
        self._set_last_action("桥接监听已启动，请把光标放到目标输入框后按 F8")
        self._set_buttons_running(True)
        self._log(f"监听已启动：{' '.join(command)}")
        self.reader = threading.Thread(target=self._stream_process_output, args=(self.process, True), daemon=True)
        self.reader.start()

    def stop_watch(self) -> None:
        process = self.process
        if process is None or process.poll() is not None:
            self.process = None
            self.status_var.set("未运行")
            self._set_buttons_running(False)
            self._cleanup_orphan_bridge_processes()
            return
        self._log("正在停止桥接监听")
        self._terminate_process_tree(process)
        self.process = None
        self._cleanup_orphan_bridge_processes()
        self.status_var.set("已停止")
        self._set_last_action("桥接监听已停止")
        self._set_buttons_running(False)

    def shutdown(self) -> None:
        if self.is_running():
            self.stop_watch()
        else:
            self._cleanup_orphan_bridge_processes()

    def is_running(self) -> bool:
        return self.process is not None and self.process.poll() is None

    def _run_hidden_command(self, command: list[str], timeout: int = 8) -> subprocess.CompletedProcess:
        creationflags = CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
        return subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True,
            encoding="mbcs" if sys.platform.startswith("win") else "utf-8",
            errors="replace",
            creationflags=creationflags,
            timeout=timeout,
        )

    def _terminate_process_tree(self, process: subprocess.Popen) -> None:
        if process.poll() is not None:
            return

        if sys.platform.startswith("win"):
            try:
                completed = self._run_hidden_command(["taskkill", "/PID", str(process.pid), "/T", "/F"])
                if completed.returncode != 0 and completed.stdout:
                    self._log(f"结束桥接进程树返回 {completed.returncode}：{completed.stdout.strip()}", "WARN")
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    pass
            except Exception as exc:
                self._log(f"taskkill 结束桥接进程树失败：{exc}", "WARN")

        if process.poll() is not None:
            return

        try:
            process.terminate()
            try:
                process.wait(timeout=4)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=4)
        except Exception as exc:
            self._log(f"停止监听失败：{exc}", "WARN")

    def _cleanup_orphan_bridge_processes(self) -> None:
        if not sys.platform.startswith("win"):
            return
        for image in BRIDGE_PROCESS_IMAGES:
            try:
                completed = self._run_hidden_command(["taskkill", "/IM", image, "/T", "/F"], timeout=5)
                if completed.returncode == 0:
                    self._log(f"已清理残留桥接后台进程：{image}")
            except Exception as exc:
                self._log(f"清理残留桥接进程失败：{image}：{exc}", "WARN")

    def _spawn(self, command: list[str]) -> subprocess.Popen:
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        creationflags = CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
        return subprocess.Popen(
            command,
            cwd=str(APP_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            creationflags=creationflags,
        )

    def _stream_process_output(self, process: subprocess.Popen, managed: bool = False) -> None:
        try:
            if process.stdout is not None:
                for line in process.stdout:
                    clean = line.rstrip()
                    if clean:
                        self._log(clean, self._level_from_line(clean))
        finally:
            code = process.poll()
            if code is None:
                code = process.wait()
            if managed and process is self.process:
                self.process = None
                self.after(0, lambda: self.status_var.set(f"已退出 {code}"))
                self.after(0, lambda: self._set_buttons_running(False))
                self.after(0, lambda: self._set_last_action(f"桥接进程已退出，退出码 {code}"))
                self._log(f"桥接进程已退出，退出码 {code}", "INFO" if code == 0 else "WARN")

    @staticmethod
    def _level_from_line(line: str) -> str:
        upper = line.upper()
        if "ERROR" in upper or "失败" in line:
            return "ERROR"
        if "WARNING" in upper or "WARN" in upper or "警告" in line:
            return "WARN"
        return "INFO"

    def _set_buttons_running(self, running: bool) -> None:
        if self.start_button:
            self.start_button.configure(state=DISABLED if running else NORMAL)
        if self.stop_button:
            self.stop_button.configure(state=NORMAL if running else DISABLED)

    def _set_last_action(self, message: str) -> None:
        self.last_action_var.set(message)

    def _log(self, message: str, level: str = "INFO") -> None:
        self.logger.emit_bridge(message, level)
        if self.output_text:
            self.after(0, lambda: self._append_output(message, level))

    def _append_output(self, message: str, level: str) -> None:
        if not self.output_text:
            return
        self.output_text.configure(state=NORMAL)
        self.output_text.insert(END, f"{datetime.now():%H:%M:%S} [{level}] {message}\n")
        self.output_text.see(END)
        self.output_text.configure(state=DISABLED)

    def open_config(self) -> None:
        if not CONFIG_PATH.exists():
            self.save_config()
        subprocess.Popen(["notepad", str(CONFIG_PATH)])

    def open_logs_dir(self) -> None:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        subprocess.Popen(["explorer", str(LOG_DIR)])


class MainWindow(Tk):
    def __init__(self) -> None:
        super().__init__()
        self.logger = GuiLogger()
        self.title(f"{APP_NAME} {APP_VERSION}")
        icon_path = get_icon_path()
        if icon_path.exists():
            try:
                self.iconbitmap(default=str(icon_path))
            except Exception as exc:
                self.logger.emit(f"设置窗口图标失败：{exc}", "WARN")
        self.geometry("960x820")
        self.minsize(900, 720)
        self.app = DouBaoVoiceBridgeApp(self, self.logger)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(250, self.app.start_after_ui_ready)
        self.logger.emit(f"{APP_NAME} 已启动", "INFO")

    def on_close(self) -> None:
        if self.app.is_running():
            if not messagebox.askyesno(APP_NAME, "豆包跨屏输入桥接正在运行。关闭软件会停止桥接，确定关闭吗？"):
                return
            self.app.stop_watch()
        self.app.shutdown()
        self.destroy()


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--self-test", action="store_true", help="启动并立即关闭 GUI，用于构建后自检")
    parser.add_argument("--version", action="store_true", help="打印版本")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    if args.version:
        print(f"{APP_NAME} {APP_VERSION}")
        return 0

    window = MainWindow()
    if args.self_test:
        window.update_idletasks()
        window.app._bridge_exe_path()
        window.destroy()
        print("self-test ok")
        return 0

    window.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
