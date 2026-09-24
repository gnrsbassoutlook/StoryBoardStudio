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
  - 启动 / 更新 / 推送三个脚本按平台成对提供：
    - 🍎 macOS：`launch_sbs_mac.command`、`update-sbs.command`、`push-sbs.command`
    - 🪟 Windows：`launch_sbs_win.bat`、`update-sbs_win.bat`、`push-sbs_win.bat`
    - 🐧 Linux：直接用「手动启动」的命令
- **Python 3.10+**（开发验证于 3.13）
  - Windows 安装时**务必勾选 `Add Python to PATH`**，否则脚本找不到 `python`
- **浏览器**：**Safari 优先**（已针对 WebKit 的 sticky、Shift 组合键差异做过适配），Chrome / Edge 兼容
- **网络**：默认走**官方 PyPI 源**，无需任何额外配置；如果官方源拉得慢，按下方「依赖装太慢」一节临时换镜像即可

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

双击 **`launch_sbs_win.bat`**：自动创建虚拟环境 → 安装依赖 → 启动服务 → 打开浏览器。
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

### 🔁 依赖装太慢？国内镜像怎么切

**默认走官方 PyPI 源**（`https://pypi.org/simple`），不需要改任何文件、也不需要额外配置。

国内网络要长期加速时，仓库里**已经备好一份镜像版清单**，改个名就生效，**启动脚本一行都不用动**：

```bash
cd /path/to/StoryBoardStudio

# ①（可选）先把官方版留个底，方便随时换回来
mv requirements.txt requirements_官方源.txt.bak

# ② 镜像版顶上，改完就能直接用启动脚本了
mv requirement_国内源.txt requirements.txt
```

两个文件依赖完全一致，`requirement_国内源.txt` 只多一行：

```
--index-url https://pypi.tuna.tsinghua.edu.cn/simple
```

想换回官方源：

```bash
git checkout requirements.txt                                    # 有 git 就直接还原
mv requirements_官方源.txt.bak requirements.txt                   # 或者把刚才的备份改回来
```

> 只是临时跑一次、不想动文件也行：`pip install -r requirements.txt -i <镜像地址>`
>
> ⚠️ **平时别把 `--index-url` 留在 `requirements.txt` 里** —— 一旦那个源不通，整个安装会卡死在那儿，还不好排查（这正是当初踩的坑）。所以放在单独一份文件里按需切换，而不是默认写死。

## 📁 目录结构

```
StoryBoardStudio/
├── StoryBoardStudio.py          # 主程序：第二部分（分镜表）+ 整体 UI + 服务入口
├── asset_tab.py                 # 第一部分（资产表）独立模块，自带 CSS/JS/索引/保存接口
├── requirements.txt             # Python 依赖（官方 PyPI 源）★ 启动脚本固定读这个文件
├── requirement_国内源.txt        # 可选：国内镜像版依赖清单（重命名为 requirements.txt 即生效）
├── README.md                    # 本文件
├── .gitignore                   # 忽略规则（venv / __pycache__ / _backups / .DS_Store …）
│
├── launch_sbs_mac.command       # 🍎 macOS 一键启动（建 venv → 装依赖 → 起服务 → 开浏览器）
├── update-sbs.command           # 🍎 macOS 更新（薄壳，转调 sbs_update.py）
├── push-sbs.command             # 🍎 macOS 提交并推送到 GitHub
├── launch_sbs_win.bat           # 🪟 Windows 一键启动
├── update-sbs_win.bat           # 🪟 Windows 更新（薄壳，转调 sbs_update.py）
├── push-sbs_win.bat             # 🪟 Windows 提交并推送
├── sbs_update.py                # ★ 交互式更新器本体（Mac / Win 共用，列 20 条提交选版本）
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

界面上改完路径**失焦（点别处）就会自动写回**这个文件，不需要手改，也不必等点「🚀 加载/刷新工程」。

**刷新或重开网页时，顶部那 5 个路径框会自动回填上一次用的路径**，不用每次重敲：

- 打开 / 刷新页面 → 从 `config.json` 把 5 个路径读回输入框
- 手打、粘贴、以及点「📁 浏览…」选出来的路径 → 立即落盘（失焦或内容变化时）
- 内容跟文件里一致时不会重复写盘，所以 `config.json` 的修改时间只会因为你真的改了路径而变

> `gradio_temp/` 属于**运行时目录，已被 `.gitignore` 忽略**（`config.json` 里是本机路径，属于个人环境信息，不该进仓库）。所以换台机器克隆下来后，第一次打开是空的，填一次就记住了。

### 运行时自动生成的目录

```
<资产目录>/excel_backup/          # 每次写回 Excel 前的原文件备份（见「Excel 同步与备份」）
```

### ✅ 该提交进仓库的 / ❌ 不该提交的

**该提交**：`StoryBoardStudio.py`、`asset_tab.py`、`requirements.txt`、`requirement_国内源.txt`、`README.md`、`.gitignore`、三个 `.command` 脚本（macOS）、三个 `_win.bat` 脚本（Windows）

**不该提交**（已被 `.gitignore` 屏蔽）：

```
__pycache__/                     # Python 字节码缓存，删了会自动重建
*.pyc  *.pyc.*                   # py_compile 失败时留下的临时字节码
_backups/                        # 代码快照 + Excel 备份回退目录
gradio_temp/                     # 运行时目录：config.json 存的是本机 5 个路径（个人环境信息）
.DS_Store                        # macOS 目录元数据
venv/                            # 虚拟环境（首次运行本机自动创建）
```

> `.DS_Store` 已通过 `git rm --cached .DS_Store` **退出跟踪**，不会再被提交。
> 如果你在别的机器上还有旧克隆，那台机器要自己跑一次 `git rm --cached .DS_Store` 才能同步。
>
> `gradio_temp/config.json` 同样已 **`git rm --cached` 退出跟踪**（它存的是「你这台机器的路径」，
> 进仓库既会泄露本地目录结构，也会让别人克隆下来看到你的路径）。如果你有旧克隆，
> 那台机器也跑一次 `git rm --cached gradio_temp/config.json`；不跑也不影响使用，只是 `git status` 里会多一条删除。

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

**两个 Tab 各有一排「📐 视图微调」**：`◀ 栏宽收窄` / `▶ 栏宽拉宽` / `▲ 行高收紧` / `▼ 行高拉开` / `↔ 适配宽度` / `📌 记住排版`（写入 localStorage）/ `🔄 恢复默认`。表格 B 另有「✓ 隐藏/显示列」，可单独藏掉某几列。

- **横向缩放**改栏宽（整表所有列）；**纵向缩放**改行高
- 表格 B 的行高由 **B 列小标题 / D 列台词 / F 列提示词 / G 列视频占位**四个框决定，纵向缩放会一起改；**E 列图片、G 列视频画面跟「栏宽」走，不跟行高**，这样拉高一行只是多出空白，媒体不会被拉变形
- 表格 A 右端 **📷 图片预览 / 🔊 声音预览两列宽度恒定**，横向滚动时吸附在视口右边，**不参与整表缩放**
- 那几个文本框可以手动往下拖高；拖过之后它会被内联高度锁住，此时按 `Shift+g` / `Shift+h` 或 `r`，内联高度会被清掉、重新交回缩放接管

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
| `pip` 卡住不动 | 默认走官方源；国内网络慢就按「依赖装太慢」一节把 `requirement_国内源.txt` 改名为 `requirements.txt`；公司 / 校园网代理需另配 `set HTTPS_PROXY=...`（Win）或 `export HTTPS_PROXY=...`（macOS） |
| 从 Mac 拷 `venv` 到 Windows，启动报错 | venv **不能跨系统**，删掉整个 `venv/` 目录后重新双击启动脚本（见「venv 不要跨机器拷贝」） |

## 🔄 更新与推送

**macOS**

```bash
./update-sbs.command    # 交互式更新（见下方说明）
./push-sbs.command      # 提交并推送（工作区干净时会直接退出，不会产生空提交）
```

**Windows**：直接双击 `update-sbs_win.bat` / `push-sbs_win.bat`（或在 `cmd` 里运行同样名字）。

### 更新器怎么用

两个平台的更新脚本都是薄壳，真正的交互逻辑在 `sbs_update.py` 里（Mac / Win 行为完全一致，也不会被 `cmd.exe` 的中文乱码问题波及）：

1. 先连远程仓库，**本地已经是最新就直接提示「本地已是最新版本」，什么都不改就退出**；
2. 不是最新时，列出远程最近 **20 条提交**，每条带短 SHA、日期、提交信息，并标出「← 你现在的版本」；
3. 输入序号 `1`–`20` 选择要部署到本地的版本，**直接回车 = 1（最新版）**，输入 `q` 取消；输入别的会提示重新输入，不会退出；
4. 切换前若工作区有未提交改动，会自动 `git stash` 留档，并在结束时打印 `git stash pop` 的取回方法；
5. 若本地存在「不在远程上的提交」，会额外确认一次，避免把本地提交冲掉（`reset --hard` 不删数据，它们仍在 `git reflog` 里）；
6. 万一选中的是还没有 `sbs_update.py` 的早期版本，脚本会把运行器本体写回磁盘，不会让你失去启动 / 更新入口。

脚本只移动 git 指针，**不删任何文件**。想回到最新版：再跑一次，直接回车即可。

推送默认走 SSH（`git@github.com:gnrsbassoutlook/StoryBoardStudio.git`），需要本机已配置 GitHub SSH 密钥；想换 HTTPS 就编辑脚本里注释掉的那两行 —— macOS 在 `push-sbs.command`，Windows 在 `push-sbs_win.bat`。

---

> 🚧 持续迭代中。
