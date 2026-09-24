#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""StoryBoardStudio 交互式更新器（macOS / Windows / Linux 通用）

由 update-sbs.command / update-sbs_win.bat 调用，也可以直接 `python sbs_update.py`。

流程：
  1. git fetch 远程，比较本地 HEAD 与 origin/<当前分支>
  2. 一样 → 直接提示「本地已是最新版本」，什么都不改就退出
  3. 不一样 → 列出 origin 上最近 20 个提交，输入序号 1-20 选择要部署到本地的
     版本（直接回车 = 1 = 最新版，q = 退出）
  4. 切过去之后，用项目 venv（如果有）再检查一遍依赖

安全设计（只动 git 指针，不删任何文件）：
  · 切换前若工作区有未提交改动，自动 git stash 留档，并打印恢复命令
  · 本地有「不在远程上的提交」时，会额外确认一次，免得把本地提交冲掉

为什么逻辑写在 Python 而不是批处理里：
  · cmd.exe 解析 UTF-8 批处理文件有已知的字节偏移 bug（中文会突然变成
    「'xxx' 不是内部或外部命令」），所以 .bat 必须保持纯 ASCII；
    交互式菜单放在 Python 里，mac / win 行为完全一致，中文输出也不会乱码。
"""
from __future__ import annotations

import os
import subprocess
import sys
from datetime import datetime

REPO = os.path.dirname(os.path.abspath(__file__))
DEFAULT_BRANCH = "main"
MAX_LIST = 20

# 运行器本体：更新时万一回退到「还没有这些文件的旧版本」，要把它们放回来，
# 否则用户下次就找不到启动/更新脚本了。
RUNNERS = ("sbs_update.py", "update-sbs.command", "update-sbs_win.bat")


def _fix_console_encoding() -> None:
    """输出被重定向到管道/文件时强制 UTF-8，免得中文标题把脚本搞崩。

    直接双击运行（stdout 是终端）时不动它 —— Windows 上 Python 走控制台 API
    输出，中文本来就正常，强行改反而可能出问题。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            enc = (getattr(stream, "encoding", "") or "").lower()
            if not stream.isatty() and enc not in ("utf-8", "utf8"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def git(*args: str):
    """跑 git，返回 (returncode, stdout, stderr)。commit message 一律按 UTF-8 解。"""
    p = subprocess.run(["git", *args], cwd=REPO, capture_output=True)
    return (p.returncode,
            p.stdout.decode("utf-8", "replace"),
            p.stderr.decode("utf-8", "replace"))


def say(msg: str = "") -> None:
    print(msg, flush=True)


def die(msg: str, code: int = 1):
    say(msg)
    sys.exit(code)


def get_branch() -> str:
    rc, out, _ = git("rev-parse", "--abbrev-ref", "HEAD")
    name = out.strip() if rc == 0 else ""
    return name if name and name != "HEAD" else DEFAULT_BRANCH     # detached 时兜底


def show_local_head() -> str:
    rc, out, _ = git("log", "-1", "--date=short", "--pretty=format:%h  %ad  %s")
    return out.strip() if rc == 0 and out.strip() else "(未知)"


def venv_python() -> str:
    """返回项目 venv 里的 python 路径；没建或不可用返回 ""。"""
    for rel in (("venv", "Scripts", "python.exe"), ("venv", "bin", "python")):
        p = os.path.join(REPO, *rel)
        if os.path.exists(p):
            return p
    return ""


def install_deps() -> None:
    if not os.path.exists(os.path.join(REPO, "requirements.txt")):
        return
    py = venv_python()
    if not py:
        say()
        say("ℹ️  未发现可用的 venv，跳过依赖更新（双击启动脚本会自动创建）。")
        return
    say()
    say("📦 正在检查并更新依赖包 ...")
    try:
        rc = subprocess.run([py, "-m", "pip", "install", "-q", "-r", "requirements.txt"],
                            cwd=REPO).returncode
    except Exception as e:                                   # noqa: BLE001
        say(f"⚠️  依赖安装没能启动：{e}")
        return
    if rc == 0:
        say("✅ 依赖已就绪。")
    else:
        say("⚠️  依赖安装失败；不影响代码，双击启动脚本时会再试一次。")


def snapshot_runners() -> dict:
    """把运行器脚本内容读进内存（纯文本、几 KB），切换版本后用于自愈。"""
    snap = {}
    for name in RUNNERS:
        p = os.path.join(REPO, name)
        try:
            if os.path.isfile(p):
                with open(p, "rb") as f:
                    snap[name] = f.read()
        except Exception:
            pass
    return snap


def restore_runners(snap: dict) -> None:
    """回退到旧版本后，若运行器脚本被一起回退没了，就地写回。"""
    fixed = []
    for name, data in snap.items():
        p = os.path.join(REPO, name)
        if not os.path.exists(p):
            try:
                with open(p, "wb") as f:
                    f.write(data)
                if name.endswith(".command"):
                    os.chmod(p, 0o755)
                fixed.append(name)
            except Exception:
                pass
    if fixed:
        say()
        say("🛠️  该版本里还没有这些脚本，已为你保留：" + "、".join(fixed))


def ask(prompt: str):
    """读一行输入；Ctrl+C / Ctrl+D 返回 None 表示取消。"""
    try:
        return input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        say()
        return None


def main() -> int:
    _fix_console_encoding()
    os.chdir(REPO)
    runners = snapshot_runners()

    say("=================================================")
    say("  🔄 StoryBoardStudio 更新器")
    say("=================================================")
    say()

    if not os.path.isdir(os.path.join(REPO, ".git")):
        die("❌ 当前目录不是 git 仓库，请先 git clone 本项目。")

    branch = get_branch()
    remote_ref = f"origin/{branch}"

    say("⏳ 正在读取远程仓库 ...")
    rc, _, err = git("fetch", "--prune", "origin")
    if rc != 0:
        die("❌ 连接远程仓库失败，请检查网络（或 GitHub SSH 密钥）后重试。\n" + err.strip())

    rc, local_sha, _ = git("rev-parse", "HEAD")
    rc2, remote_sha, err2 = git("rev-parse", remote_ref)
    if rc != 0 or rc2 != 0:
        die(f"❌ 找不到远程分支 {remote_ref}，请确认仓库状态。\n{err2.strip()}")
    local_sha, remote_sha = local_sha.strip(), remote_sha.strip()

    say(f"📌 本地当前版本：{show_local_head()}")

    if local_sha == remote_sha:
        say()
        say(f"✅ 本地已是最新版本（{remote_ref}），无需更新。")
        return 0

    # ---- 列出远程最近 20 个提交 ----
    rc, out, err = git("log", f"-n{MAX_LIST}", "--date=short",
                       "--pretty=format:%H%x1f%h%x1f%ad%x1f%s", remote_ref)
    if rc != 0 or not out.strip():
        die("❌ 读取远程提交记录失败。\n" + err.strip())

    commits = [ln.split("\x1f") for ln in out.splitlines() if len(ln.split("\x1f")) == 4]
    if not commits:
        die("❌ 远程没有任何提交记录。")

    say()
    say(f"============== {remote_ref} 最近 {len(commits)} 个提交 ==============")
    for i, (_full, short, date, subject) in enumerate(commits, 1):
        mark = "   ← 你现在的版本" if _full == local_sha else ""
        say(f"  {i:>2}) {short}  {date}  {subject}{mark}")
    say("=" * 56)
    say()
    say(f"请输入要部署的版本序号（1-{len(commits)}，直接回车 = 1 最新版，q = 退出）：")

    # ---- 选择 ----
    while True:
        raw = ask("> ")
        if raw is None:
            say("已取消，未做任何改动。")
            return 1
        if raw.lower() in ("q", "quit", "exit"):
            say("已取消，未做任何改动。")
            return 1
        if raw == "":
            idx = 1                      # 直接回车 = 最新版
            break
        if raw.isdigit() and 1 <= int(raw) <= len(commits):
            idx = int(raw)
            break
        say(f"⚠️  请输入 1-{len(commits)} 之间的数字（或 q 退出）：")

    full, short, date, subject = commits[idx - 1]
    say()
    say(f"👉 已选：第 {idx} 个   {short}  {date}  {subject}")

    # ---- 本地有远程没有的提交？先确认 ----
    rc, ahead_out, _ = git("rev-list", "--count", f"{remote_ref}..HEAD")
    ahead = int(ahead_out.strip()) if rc == 0 and ahead_out.strip().isdigit() else 0
    if ahead > 0:
        say()
        say(f"⚠️  本地有 {ahead} 个提交不在 {remote_ref} 上。")
        say("    继续会把分支退回到你选的版本：这些本地提交不会消失（仍在 git reflog 里），")
        say("    但更稳妥的做法是先把它们 push 上去。")
        if (ask("    仍要继续吗？(y/N) > ") or "").lower() not in ("y", "yes"):
            say("已取消，未做任何改动。")
            return 1

    # ---- 工作区有未提交改动？自动 stash 留档 ----
    rc, status_out, _ = git("status", "--porcelain")
    stashed = ""
    if status_out.strip():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        stashed = f"sbs_update 自动留档 {stamp}"
        rc, _, err = git("stash", "push", "-u", "-m", stashed)
        if rc != 0:
            say("⚠️  自动备份本地改动失败：" + err.strip())
            say("    为避免丢数据已中止更新，请先手动处理（git status 看看）。")
            return 1
        say()
        say("📦 检测到工作区有未提交改动，已自动备份为 git stash，不会丢。")

    # ---- 切换版本 ----
    if local_sha != full:
        rc, _, err = git("reset", "--hard", full)
        if rc != 0:
            say("❌ 切换版本失败：" + err.strip())
            return 1
    else:
        say("ℹ️  当前已经就是这个版本，跳过切换。")

    restore_runners(runners)

    now_sha = git("rev-parse", "HEAD")[1].strip() or full

    say()
    say("🎉 已部署到本地：")
    say("   " + show_local_head())

    if stashed:
        say()
        say("📦 你切换前的本地改动存进了 git stash，需要时这样取回：")
        say("   git stash list      # 看看有哪些留档")
        say("   git stash pop       # 恢复最近一次留档")

    install_deps()

    if now_sha != remote_sha:
        say()
        say("ℹ️  想回到最新版：再运行一次本脚本，直接回车（= 第 1 项）即可。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
