# StoryBoardStudio

> Excel 驱动的影视分镜 / 资产表可视化工作台。
> Python + Gradio + FastAPI，在浏览器里直接看分镜、图片、音频、视频，改完一键写回 Excel 源文件。

界面分两个 Tab：

| Tab | 内容 | 数据源 |
| --- | --- | --- |
| **📊 表格 A · 影视资产表** | X（人物）/ Y（场景）/ Z（道具）三个 sheet，含提示词、参考音频、声音特质、台词；右端固定「图片预览 / 声音预览」两列，单元格可直接编辑 | `path_a` 指向的资产 Excel |
| **🎬 表格 B · 分镜表** | A–G 列分镜，按镜头号自动匹配音频 / 图片 / 视频，小窗悬停播放、灯箱放大 | `path_b` 指向的分镜 Excel |

## ✨ 核心特性

- 📊 **Excel 原生驱动**：直接读写 `.xlsx`，多 Sheet 切换，sheet 顺序严格按源文件，不做排序
- 🖼️ **图片资产预览**：按编码精确匹配，表格 A 的缩略图逐行对齐行高、可点开灯箱
- 🔊 **音频资产预览**：多轨编码匹配，内置播放器；预览列宽度恒定，不随整表缩放而挤掉进度条
- 🎬 **视频资产预览**：按镜头号自动匹配（`mp4 / mov / webm / mkv`）
- 🎛️ **精细视图控制**：栏宽 / 行高缩放、适配宽度、布局记忆（存 localStorage）
- 📝 **提示词大字区**：点行在上方大字看整行文本
- 💾 **一键写回**：保存前自动备份原 Excel，写坏了大不了回滚
- 🌙 **深色沉浸主题**

## 💻 环境要求

- **操作系统**：macOS（Intel / Apple Silicon 都行）、Windows 10+、Linux
  - Mac 用 `.command` 脚本，Windows 用 `.bat` 脚本，Linux 直接用「手动启动」的命令
- **Python 3.10+**（开发验证于 3.13）
  - Windows 安装时**务必勾选 `Add Python to PATH`**，否则脚本找不到 `python`
- **浏览器**：**Safari 优先**（已针对 WebKit 的 sticky、Shift 组合键差异做过适配），Chrome / Edge 兼容
- **网络**：默认走清华 PyPI 镜像；海外网络可在 `requirements.txt` 里删掉 `--index-url` 一行改回官方源

## 🚀 安装与首次运行

### 🍎 macOS

#### 0. 从 GitHub 下载后先做两件事（**必做**）

从 zip 解压或拷贝过来的文件会带上 macOS 的隔离属性，同时脚本的可执行位会丢失 —— 不处理的话双击 `.command` 会被 Gatekeeper 拦下，报「无法打开，因为来自身份不明的开发者」。

```bash
cd /path/to/StoryBoardStudio

# ① 解除隔离属性（递归处理整个目录）
xattr -dr com.apple.quarantine .

#    只处理单个文件也行：
#    xattr -d com.apple.quarantine launch_sbs_mac.command
#    查看某个文件当前有没有隔离属性（没输出 = 干净）：
#    xattr -l launch_sbs_mac.command

# ② 恢复可执行权限（三个 .command 脚本都要）
chmod +x launch_sbs_mac.command push-sbs.command update-sbs.command

#    确认权限已生效：输出里要能看到 x
ls -l *.command
```

> 用 `git clone` 拉下来的通常不带隔离属性，但 `chmod +x` 仍建议跑一遍 —— 从别的磁盘 / 网盘拷贝过来时会丢权限位。

#### 1. 一键启动

双击 **`launch_sbs_mac.command`**：自动创建虚拟环境 → 安装依赖 → 启动服务 → 打开浏览器。
访问地址：<http://localhost:7861>

脚本会先自检端口 7861：**已经在运行就不会重复启动**，只帮你打开页面（重复启动会报 `[Errno 48] address already in use`，看着吓人其实只是「已经开着呢」）。

#### 2. 手动启动（等价命令）

```bash
cd /path/to/StoryBoardStudio
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python StoryBoardStudio.py
```

停止服务：在运行窗口按 `Ctrl+C`。
强杀残留进程：`pkill -f "StoryBoardStu[d]io.py"`（别写 `pkill -f StoryBoardStudio.py` —— 你自己的命令行里也含这个字符串，会连当前 shell 一起杀掉）。

### 🪟 Windows

#### 1. 一键启动

双击 **`launch_sbs.bat`**：自动创建虚拟环境 → 安装依赖 → 启动服务 → 打开浏览器。
访问地址：<http://localhost:7861>

脚本同样会自检 7861 端口，**已经在跑就不会重复启动**，只帮你打开页面。

> **Windows 不需要 `xattr` / `chmod`** —— 那两条只针对 macOS 的隔离属性与 Unix 可执行位。
> `.bat` 内容为 UTF-8，脚本开头会 `chcp 65001` 切到 UTF-8 代码页；若你的老 `cmd` 仍显示乱码，换 Windows Terminal 打开即可。

#### 2. 手动启动（等价命令）

```bat
cd /path\to\StoryBoardStudio
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python StoryBoardStudio.py
```

停止服务：在运行窗口按 `Ctrl+C`。
强杀残留进程：`taskkill /F /IM python.exe`（这会杀掉**所有** python 进程，确认没有别的任务在跑再用）。

### 🔁 依赖装太慢？（国内镜像源）

`requirements.txt` 顶部已经写好了清华镜像，所以 `pip install -r requirements.txt` **默认就走国内源**：

```
--index-url https://pypi.tuna.tsinghua.edu.cn/simple
```

- **镜像抽风 / 人在海外**：把那一行**删掉**，即自动回到官方 PyPI
- **想换源**：文件顶部注释里已列好阿里云 / 腾讯云 / 中科大三套现成的 `--index-url`，换一行即可
- **临时指定一次**：`pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/`

## 📁 目录结构

```
StoryBoardStudio/
├── StoryBoardStudio.py          # 主程序：第二部分（分镜表）+ 整体 UI + 服务入口
├── asset_tab.py                 # 第一部分（资产表）独立模块，自带 CSS/JS/索引/保存接口
├── requirements.txt             # Python 依赖（gradio / pandas / openpyxl / Pillow / fastapi / uvicorn）
├── README.md                    # 本文件
├── .gitignore                   # 忽略规则（venv / __pycache__ / _backups / .DS_Store …）
│
├── launch_sbs_mac.command       # 🍎 macOS 一键启动（建 venv → 装依赖 → 起服务 → 开浏览器）
├── update-sbs.command           # 🍎 macOS 拉取最新代码（先列最近 20 条提交，确认后再拉）
├── push-sbs.command             # 🍎 macOS 提交并推送到 GitHub
├── launch_sbs.bat               # 🪟 Windows 一键启动
├── update-sbs.bat               # 🪟 Windows 拉取最新代码
├── push-sbs.bat                 # 🪟 Windows 提交并推送
│
├── gradio_temp/                 # Gradio 本地运行目录
│   └── config.json              # ★ 本地配置（见下方「配置文件」）
├── venv/                        # 虚拟环境（★ 不在仓库里，首次运行由本机自动创建）
└── _backups/                    # 代码快照 / 资产目录未填时的 Excel 备份回退位置
```

### ⚠️ venv 不要跨机器拷贝（重要）

`venv/` **不进仓库**（已被 `.gitignore` 屏蔽），每台机器首次运行时由启动脚本自动创建。

千万别把一台机器上的 `venv/` 整包拷到另一台机器，**跨操作系统更是 100% 用不了**：

| 拷贝场景 | 结果 |
| --- | --- |
| 同一台机器原地保留 | ✅ 正常 |
| 另一台 Apple Silicon Mac（用户目录/路径不同） | ❌ venv 里脚本的 shebang 写死了绝对路径，直接失效 |
| Intel Mac | ❌ site-packages 里装的是 arm64 二进制 |
| Windows / Linux | ❌ Mac 上是 Mach-O `.so`；Windows 要 `.dll` / `.pyd`，且目录结构是 `Scripts\` 而非 `bin/` |

三个启动脚本都会**自动识别「有 venv 但不可用」**的情况（检测 `venv/bin/python` 或 `venv\Scripts\python.exe` 是否存在），改名留档为 `venv_broken_*` 后重建，一般不用你操心。想手动清理就整个删掉 `venv/` 再启动一次。

`asset_tab.py` 是**必需文件**，缺了主程序起不来（ImportError）。
在 `StoryBoardStudio.py` 里加功能时，**新逻辑请写进独立模块**（照 `asset_tab.py` 的样子：自带 CSS/JS/索引/接口，类名统一加前缀），主程序只保留 import + 挂载那几行，避免两部分的代码互相污染。

### 配置文件 `gradio_temp/config.json`

```json
{
  "path_a": "资产表 .xlsx 的绝对路径",
  "path_b": "分镜表 .xlsx 的绝对路径",
  "asset_dir": "素材根目录",
  "audio_dir": "音频目录（留空 → 资产目录/Audio → 资产目录）",
  "video_dir": "视频目录（留空 → 同资产目录）"
}
```

界面上改完路径点「🚀 加载/刷新工程」会自动写回这个文件，不需要手改。

### 运行时自动生成的目录

```
<资产目录>/excel_backup/          # 每次写回 Excel 前的原文件备份（见「Excel 同步与备份」）
```

### ✅ 该提交进仓库的 / ❌ 不该提交的

**该提交**：`StoryBoardStudio.py`、`asset_tab.py`、`requirements.txt`、`README.md`、`.gitignore`、三个 `.command` 脚本（macOS）、三个 `.bat` 脚本（Windows）

**不该提交**（已被 `.gitignore` 屏蔽）：

```
__pycache__/                     # Python 字节码缓存，删了会自动重建
*.pyc  *.pyc.*                   # py_compile 失败时留下的临时字节码
_backups/                        # 代码快照 + Excel 备份回退目录
gradio_temp/config.json.bak_*    # 配置的手工备份
.DS_Store                        # macOS 目录元数据
venv/                            # 虚拟环境（首次运行本机自动创建）
```

> `.DS_Store` 本次已通过 `git rm --cached .DS_Store` **退出跟踪**，不会再被提交。
> 如果你在别的机器上还有旧克隆，那台机器要自己跑一次 `git rm --cached .DS_Store` 才能同步。

## ⌨️ 快捷键

键盘上 `a s d f g h` 连成一排，手指不用挪窝。**焦点在输入框 / 下拉框里时不生效**（正常打字优先）。

| 按键 | 作用 |
| --- | --- |
| `a` / `s` | **当前这张表内切 sheet** —— 表格 A 走「◀ 上一表 / 下一表 ▶」，表格 B 走「上一场 / 下一场」 |
| `d` / `f` | **切 Tab 页面**：`d` → 表格 A，`f` → 表格 B |
| `g` / `h` | 横向缩放（整表栏宽）：`g` 收窄，`h` 拉宽 |
| `Shift` + `g` / `h` | 纵向缩放（行高）：`Shift+g` 收紧，`Shift+h` 拉开 |
| `r` | 恢复默认排版（对当前所在的 Tab 生效） |
| `Shift` + 滚轮 | 表格横向滚动 |

## 📖 功能说明

### 表格 A · 影视资产表

- 列头带 **Excel 真实列号前缀**：`A.资产编号`、`B.资产名称`、`C.文生图 / 图生图提示词`……
- 最右侧 **📷 图片预览 / 🔊 声音预览两列吸附在视口右边**，横向滚动时始终可见；这两列宽度恒定，**不参与整表缩放**
- 图片缩略图高度**逐行对齐到该行最高的单元格**；一条编号只显示它自己那张图
- 单元格可直接编辑，点 **💾 同步并保存到 Excel** 按绝对行号写回源文件（空行不错位）

**资产编码硬规则**（不符合的一律**不参与匹配**，只在页面顶部红字提示并给出建议写法）：

| 规则 | ✅ 合法 | ❌ 废片 |
| --- | --- | --- |
| `X`/`Y`/`Z` + **两位数字**（个位补 0） | `X01`、`Y02`、`Z01` | `X1`、`Y2`、`X001` |
| 后缀字母**必须小写** | `X01a` | `X01A`、`X01C` |
| 接描述**必须用英文句点** `.` | `X01.何功伟.jpg`、`X01a.何功伟狱中.png` | `X01何功伟行刑.jpg`、`X4A父亲.png` |
| 不认空格 / 短横 / 括号 | —— | `X1a 何功伟绝食状态-.jpg` |

匹配是**严格精确**的：`X01` 只找 `X01.*`，`X01a` 只找 `X01a.*`，**`X01 ≠ X1`**，不做前导零归一、不做模糊匹配。
衍生形态（`X01a`）的「参考音频」列只填**去掉后缀的裸基底编号**（`X01a` → `X01`）。

素材目录**递归扫描子文件夹**，`Audio/02_主角/X01.flac` 一样能认到。

### 表格 B · 分镜表

| 列 | 说明 |
| --- | --- |
| A. 镜头号 | 镜头唯一编号，用于音频 / 视频自动匹配 |
| B. 小标题 | 镜头简短描述 |
| C. 音频结果（模式1）/ 镜头时长（模式2） | 双模式兼容，见下 |
| D. 台词内容（模式1）/ 台词角色（模式2） | 台词文本 |
| E. 图片资产 | 图片编码，用竖线分隔多图（例：`X01` 竖线 `X01a` 竖线 `Z01`） |
| F. 生视频提示词 | 完整画面描述，多行 |
| G. 视频结果 | 按镜头号自动匹配 |

**两种生成模式**

- **模式 1 · 音频先行**：C 列填**镜头号**（如 `s01-01`），该镜头的语音时长就是视频时长依据；D 列填台词内容
- **模式 2 · 时长先行**：C 列填**数字时长**，D 列不参与；图片 + 音频一次性送进参考生成视频节点

### 视图操作

栏宽收窄 / 拉宽、行高收紧 / 拉开、适配宽度、记住排版（写入 localStorage）、恢复默认，以及表格 B 的列显示 / 隐藏。

## 💾 Excel 同步与备份

- 点「💾 同步并保存到 Excel」把网页改动写回源文件（两个 Tab 各自独立）
- **写前自动备份**到：

  ```
  <资产目录>/excel_backup/<原文件名>_<YYYYmmdd_HHMMSS>.xlsx
  ```

  资产目录没填或不存在时，退回 Excel 同目录的 `_backups/`
- 保存提示里会带上备份文件名，方便回溯

## ⚠️ 注意事项 / 常见问题

| 现象 | 原因与处理 |
| --- | --- |
| 双击 `.command` 打不开，提示「身份不明的开发者」 | 隔离属性没解除：`xattr -dr com.apple.quarantine .`（见「安装」第 0 步） |
| 双击 `.command` 一闪而过 / Permission denied | 缺可执行位：`chmod +x *.command` |
| 终端报 `[Errno 48] address already in use` | 7861 已被占用，说明服务本来就在跑 → 直接开 <http://localhost:7861>；要重启先在原窗口 `Ctrl+C` |
| 改了代码但页面没变 | 老实例还占着端口，新进程没起来 → `lsof -nP -iTCP:7861 -sTCP:LISTEN` 确认后重启 |
| 保存报「文件正在被占用」 | 用 Excel / WPS 打开着同一个文件，关掉再保存 |
| 图片 / 音频显示「无文件」 | 素材命名不符合上面的硬规则，页面顶部红字会给出建议改名（**不会自动改名**） |
| Safari 下 `Shift` 组合键不灵 | 已按 `e.shiftKey` 判定；若仍异常请确认 Safari 升级到较新版本 |
| **Win**：`.bat` 双击一闪而过 | 先在 `cmd` 里手动跑一遍看报错；多为没勾 `Add Python to PATH`，或公司策略禁止脚本执行 |
| **Win**：提示 `python` 不是内部或外部命令 | Python 没加进 PATH → 重装时勾选 `Add Python to PATH`，或改用 `py -3` 代替 `python` |
| **Win**：`.bat` 中文提示乱码 | 脚本已 `chcp 65001`；老 `cmd` 字体不全时换 Windows Terminal |
| **Win**：`netstat` / `taskkill` 提示拒绝访问 | 用「以管理员身份运行」重开终端 |
| **Win**：`pip` 卡住不动 | 走镜像仍慢就 `pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/`；公司代理需另配 `set HTTPS_PROXY=...` |
| 从 Mac 拷 `venv` 到 Windows，启动报错 | venv **不能跨系统**，删掉整个 `venv/` 目录后重新双击启动脚本（见「venv 不要跨机器拷贝」） |

## 🔄 更新与推送

**macOS**

```bash
./update-sbs.command    # 拉取远程最新代码（会先列出最近 20 条提交让你确认）
./push-sbs.command      # 提交并推送（工作区干净时会直接退出，不会产生空提交）
```

**Windows**：直接双击 `update-sbs.bat` / `push-sbs.bat`（或在 `cmd` 里运行同样名字）。

推送默认走 SSH（`git@github.com:gnrsbassoutlook/StoryBoardStudio.git`），需要本机已配置 GitHub SSH 密钥；想换 HTTPS 就编辑脚本里注释掉的那两行 —— macOS 在 `push-sbs.command`，Windows 在 `push-sbs.bat`。

---

> 🚧 持续迭代中。
