#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""StoryBoardStudio 交互式更新器（macOS / Windows / Linux 通用）

由 update-sbs.command / update-sbs_win.bat 调用，也可以直接 `python sbs_update.py`。

流程：
  1. 自动选通道：SSH 可用就用 SSH，否则回退 HTTPS，并同步 origin 地址
  2. git fetch 远程，列出 origin/<当前分支> 上最近 20 个提交
  3. 输入序号 1-20 选择要部署到本地的版本（回车 = 1 = 最新版，q = 取消）
  4. 能快进就快进（merge --ff-only），选了历史版本才提示确认后回退
  5. 切完顺便把 *.command 的执行权限和隔离属性修好，再问要不要更新依赖

为什么这么改（对齐 AutoVideoStudio 的新版推送/更新脚本）：
  · 本地落后时不再一上来就 reset --hard，优先走快进，避免把本地分支历史冲掉
  · SSH 不通自动切 HTTPS，换台电脑没配密钥也能更新（旧版只会报「连接失败」）
  · 选完版本打印本次变动的文件清单，更新前/后版本号一目了然
  · 切换后自动修复 .command 权限 + 清 quarantine，双击就能跑，不用再手工 chmod

安全设计：
  · 切换前若工作区有未提交改动，自动 git stash 留档，并打印恢复命令
  · 本地有远程没有的提交、或要往回/横向切换时，额外确认一次才动 git 指针
  · 自愈：万一回退到「还没有这些脚本的旧版本」，把更新/推送脚本就地放回来

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

REPO_SSH = "git@github.com:gnrsbassoutlook/StoryBoardStudio.git"
REPO_HTTPS = "https://github.com/gnrsbassoutlook/StoryBoardStudio.git"

IS_MAC = sys.platform == "darwin"

# 运行器本体：更新时万一回退到「还没有这些文件的旧版本」，要把它们放回来，
# 否则用户下次就找不到启动/更新/推送脚本了。
RUNNERS = (
    "sbs_update.py",
    "update-sbs.command", "update-sbs_win.bat",
    "push-sbs.command", "push-sbs_win.bat",
    "launch_sbs_mac.command", "launch_sbs_win.bat",
)


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
    """跑 git，返回 (returncode, stdout, stderr)。输出一律按 UTF-8 解。"""
    try:
        p = subprocess.run(["git", *args], cwd=REPO, capture_output=True)
    except FileNotFoundError:
        return (127, "", "git command not found")
    return (p.returncode,
            p.stdout.decode("utf-8", "replace"),
            p.stderr.decode("utf-8", "replace"))


def say(msg: str = "") -> None:
    print(msg, flush=True)


def die(msg: str, code: int = 1):
    say(msg)
    sys.exit(code)


def ask(prompt: str):
    """读一行输入；Ctrl+C / Ctrl+D 返回 None 表示取消。"""
    try:
        return input(prompt).strip()
    except (EOFError, KeyboardInterrupt):
        say()
        return None


def get_branch() -> str:
    rc, out, _ = git("rev-parse", "--abbrev-ref", "HEAD")
    name = out.strip() if rc == 0 else ""
    return name if name and name != "HEAD" else DEFAULT_BRANCH     # detached 时兜底


def show_head(ref: str = "HEAD") -> str:
    rc, out, _ = git("log", "-1", "--date=format:%m-%d %H:%M",
                     "--pretty=format:%h  %ad  %s", ref)
    return out.strip() if rc == 0 and out.strip() else "(未知)"


def ensure_remote() -> None:
    """SSH 可用就用 SSH，否则回退 HTTPS；顺手把 origin 地址纠正过来。"""
    say()
    say("⏳ 检测远程通道 ...")
    rc, _, _ = git("ls-remote", REPO_SSH, "HEAD")
    remote = REPO_SSH if rc == 0 else REPO_HTTPS
    say("        SSH 可用" if rc == 0 else "        SSH 不可用, 自动回退 HTTPS")

    rc, cur, _ = git("remote", "get-url", "origin")
    cur = cur.strip()
    if rc != 0 or not cur:
        git("remote", "add", "origin", remote)
        say("        已添加 origin 地址")
    elif cur != remote:
        git("remote", "set-url", "origin", remote)
        say("        已更新 origin 地址")


def venv_python() -> str:
    """返回项目 venv 里的 python 路径；没建或不可用返回 ""。"""
    for rel in (("venv", "Scripts", "python.exe"), ("venv", "bin", "python"),
                (".venv", "Scripts", "python.exe"), (".venv", "bin", "python")):
        p = os.path.join(REPO, *rel)
        if os.path.exists(p):
            return p
    return ""


def install_deps() -> None:
    req = os.path.join(REPO, "requirements.txt")
    if not os.path.exists(req):
        return

    say()
    ans = ask("📦 是否同时更新 Python 依赖? (Y/n) > ")
    if ans and ans.lower() in ("n", "no"):
        return

    mirror = os.path.join(REPO, "requirement_国内源.txt")
    target = req
    if os.path.exists(mirror):
        m = ask("    是否使用国内镜像源版本 requirement_国内源.txt? (y/N) > ")
        if m and m.lower() in ("y", "yes"):
            target = mirror

    py = venv_python() or sys.executable
    say()
    say(f"📦 正在更新依赖 ({os.path.basename(target)}) ...")
    try:
        rc = subprocess.run([py, "-m", "pip", "install", "-q", "-r", target],
                            cwd=REPO).returncode
    except Exception as e:                                   # noqa: BLE001
        say(f"⚠️  依赖安装没能启动：{e}")
        return
    if rc == 0:
        say("✅ 依赖已就绪。")
    else:
        say("⚠️  依赖安装失败；不影响代码，双击启动脚本时会再试一次。")


def fix_permissions() -> None:
    """macOS 上把 *.command / *.sh 补回可执行权限，并清掉隔离属性。"""
    if not IS_MAC:
        return
    fixed = []
    for root, dirs, files in os.walk(REPO):
        dirs[:] = [d for d in dirs if d != ".git"]
        for name in files:
            if not name.endswith((".command", ".sh")):
                continue
            p = os.path.join(root, name)
            try:
                os.chmod(p, 0o755)
                subprocess.run(["xattr", "-d", "com.apple.quarantine", p],
                               capture_output=True)
                if name not in fixed:
                    fixed.append(name)
            except Exception:
                pass
    if fixed:
        say()
        say("🔑 已修复可执行权限：" + "、".join(fixed))


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
                if name.endswith((".command", ".sh")):
                    os.chmod(p, 0o755)
                fixed.append(name)
            except Exception:
                pass
    if fixed:
        say()
        say("🛠️  该版本里还没有这些脚本，已为你保留：" + "、".join(fixed))


def list_commits(ref: str) -> list:
    """读远程最近 MAX_LIST 条提交，返回 [full, short, date, author, subject]。"""
    rc, out, err = git("log", f"-n{MAX_LIST}",
                       "--date=format:%m-%d %H:%M",
                       "--pretty=format:%H%x1f%h%x1f%ad%x1f%an%x1f%s", ref)
    if rc != 0 or not out.strip():
        die("❌ 读取远程提交记录失败。\n" + err.strip())
    rows = []
    for ln in out.splitlines():
        parts = ln.split("\x1f")
        if len(parts) == 5:
            rows.append(parts)
    if not rows:
        die("❌ 远程没有任何提交记录。")
    return rows


def show_diff_summary(old: str, new: str) -> None:
    _, out, _ = git("diff", "--name-status", old, new)
    say("  --------------------------------------------------------------")
    if out.strip():
        say(out.rstrip())
    else:
        say("  (无文件变动)")
    say("  --------------------------------------------------------------")


def main() -> int:
    _fix_console_encoding()
    os.chdir(REPO)
    runners = snapshot_runners()

    say("=================================================")
    say("  🔄 StoryBoardStudio 更新器")
    say("=================================================")

    if not os.path.isdir(os.path.join(REPO, ".git")):
        die("❌ 当前目录不是 git 仓库，请先 git clone 本项目。")

    branch = get_branch()
    remote_ref = f"origin/{branch}"

    say()
    say(f"📌 当前分支：{branch}")

    ensure_remote()

    say()
    say("⏳ 正在获取远程最新提交 ...")
    rc, _, err = git("fetch", "--prune", "origin")
    if rc != 0:
        die("❌ 连接远程仓库失败，请检查网络（或 GitHub SSH 密钥）后重试。\n" + err.strip())

    rc, local_sha, _ = git("rev-parse", "HEAD")
    rc2, remote_sha, err2 = git("rev-parse", remote_ref)
    if rc != 0 or rc2 != 0:
        die(f"❌ 找不到远程分支 {remote_ref}，请确认仓库状态。\n{err2.strip()}")
    local_sha, remote_sha = local_sha.strip(), remote_sha.strip()

    say()
    say(f"📌 本地当前版本：{show_head()}")

    if local_sha == remote_sha:
        say()
        say(f"✅ 本地已与 {remote_ref} 一致（最新）。")
        say("    下面仍会列出历史提交，可以用来回退；不想动直接按 q 退出。")

    commits = list_commits(remote_ref)
    say()
    say(f"============== {remote_ref} 最近 {len(commits)} 个提交 (1 = 最新) ==============")
    for i, (full, short, date, author, subject) in enumerate(commits, 1):
        mark = "   ← 本地当前" if full == local_sha else ""
        say(f"  {i:>2}) {short}  {date}  {author}  {subject}{mark}")
    say("=" * 62)
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

    full, short, date, author, subject = commits[idx - 1]
    say()
    say(f"👉 已选：第 {idx} 个   {short}  {date}  {author}  {subject}")

    if full == local_sha:
        say()
        say("ℹ️  本地已经停在这个提交上，无需更新。")
        return 0

    OLD = show_head()

    # ---- 本地有远程没有的提交？或者要往回/横向切换？先确认 ----
    rc, ahead_out, _ = git("rev-list", "--count", f"{remote_ref}..HEAD")
    ahead = int(ahead_out.strip()) if rc == 0 and ahead_out.strip().isdigit() else 0
    rc_detail, is_anc, _ = git("merge-base", "--is-ancestor", "HEAD", full)
    need_confirm = (ahead > 0) or (rc_detail != 0)
    if need_confirm:
        say()
        if ahead > 0:
            say(f"⚠️  本地有 {ahead} 个提交不在 {remote_ref} 上。")
        if rc_detail != 0:
            say("⚠️  所选提交不在本地历史前方（回退或横向切换）。")
        say("    这些本地提交不会消失（仍在 git reflog 里），但更稳妥的做法是先把它们 push 上去。")
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

    # ---- 切换版本：能快进就快进，否则回退 ----
    if rc_detail == 0:
        say()
        say(f"⏩ 快进同步到 {short} ...")
        rc, _, err = git("merge", "--ff-only", full)
        if rc != 0:
            say("❌ 快进失败：" + err.strip())
            say("    本地没有丢东西，需要时手动执行：git merge --abort")
            return 1
    else:
        say()
        say(f"⏪ 回退/切换到 {short} ...")
        rc, _, err = git("reset", "--hard", full)
        if rc != 0:
            say("❌ 切换版本失败：" + err.strip())
            return 1

    restore_runners(runners)
    fix_permissions()

    NEW = show_head()
    say()
    say("🎉 已部署到本地：")
    say(f"   更新前：{OLD}")
    say(f"   更新后：{NEW}")
    say()
    say("   本次变动的文件：")
    show_diff_summary(local_sha, full)

    if stashed:
        say()
        say("📦 你切换前的本地改动存进了 git stash，需要时这样取回：")
        say("   git stash list      # 看看有哪些留档")
        say("   git stash pop       # 恢复最近一次留档")

    install_deps()

    now_rc, now_sha, _ = git("rev-parse", "HEAD")
    now_sha = now_sha.strip() if now_rc == 0 else full
    if now_sha != remote_sha:
        say()
        say("ℹ️  想回到最新版：再运行一次本脚本，直接回车（= 第 1 项）即可。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
