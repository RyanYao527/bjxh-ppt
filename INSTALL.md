# bjxh-ppt 离线安装指南（本地发给同事版）

> **目标用户**：北京兴华会计师事务所内部同事（Windows 系统）
> **分发方式**：U盘 / 内网文件服务器 / 邮件附件（zip 包）
> **耗时**：5–10 分钟

---

## 1. 拿到安装包

你会收到一个 `bjxh-ppt-v1.0.zip`。解压到任意目录，得到如下结构：

```
bjxh-ppt/
├── install.bat              ← 双击这个安装
├── uninstall.bat            ← 卸载用
├── INSTALL.md               ← 本文档
├── SKILL.md                 ← skill 规范（不必看）
├── README.md                ← 项目说明
├── LICENSE
├── template_spec.json
├── scripts/
│   ├── from_outline.py
│   ├── from_template.py
│   ├── qa.py
│   ├── dump_template_spec.py
│   └── preview_html.py
└── examples/
    ├── audit_demo.md
    ├── audit_demo.pptx
    ├── xxb_2025_2026_report.md
    ├── xxb_2025_2026_report.pptx
    └── xxb_2025_2026_report.preview.html
```

**注意**：解压路径里**不要**包含中文（虽然 Windows 一般能处理，但 python 偶尔会在中文路径上出问题）。

---

## 2. 准备 Python 环境（仅第一次需要）

skill 依赖 Python 3.11+ 和 `python-pptx` / `lxml` 两个包。

### 2.1 安装 Python

1. 打开 https://www.python.org/downloads/
2. 下载 **Python 3.11 或更新**（建议 3.12）
3. 双击安装包，**关键勾选**：
   - ☑ **Add Python to PATH**（最下面那个 checkbox，不勾后面会很难受）
   - ☑ Install launcher for all users
4. 点 Install Now，等 1–2 分钟

验证：打开 PowerShell 或 cmd，输入 `python --version`，应该看到 `Python 3.11.x` 或更高。

### 2.2 安装 skill 依赖

```cmd
pip install python-pptx lxml
```

如果公司有内部 PyPI 源，提前 `pip config set global.index-url <你的内网源>`。

---

## 3. 准备母版模板（必做）

skill 默认从以下路径读取母版：

```
C:\工作\04-总结与报告\2026年工作\2026合伙人大会\北京兴华模板.pptx
```

### 3.1 如果你的电脑上有这个路径

跳过本节，直接进 4。

### 3.2 如果没有（路径不一致 / 文件不在）

从内网文件服务器 / 共享盘 / 同事那里把 **北京兴华模板.pptx** 拷贝到本地。

**两种处理方式二选一**：

#### 方式 A：放到 skill 期望的路径（推荐，最省事）

在 PowerShell 里：
```powershell
$dir = "C:\工作\04-总结与报告\2026年工作\2026合伙人大会"
New-Item -ItemType Directory -Force -Path $dir
# 然后把 北京兴华模板.pptx 拷到这个目录
```

#### 方式 B：放到你自己的位置，编辑配置文件

将 `scripts/config.example.json` 复制为 `scripts/config.json`，修改其中的 `template_path`：

```json
{
    "template_path": "D:/templates/BJXH-master.pptx",
    "company": {
        "company_name": "北京兴华集团",
        "phone": "010-82250666",
        "fax": "010-82250851",
        "address": "北京市西城区裕民路18号北环中心27层"
    }
}
```

或者设置环境变量 `BJXH_TEMPLATE` 指向模板文件路径。

---

## 4. 运行 install.bat

回到解压后的 `bjxh-ppt/` 目录，**双击 `install.bat`**。

正常情况下会看到：
```
=== bjxh-ppt skill installer ===

Source : C:\Users\你的名字\Desktop\bjxh-ppt\
Target : C:\Users\你的名字\.claude\skills\bjxh-ppt\

[OK] Installed to: C:\Users\你的名字\.claude\skills\bjxh-ppt\
```

按任意键关闭窗口。

如果 install.bat 报 [WARN] 关于 Python 或 python-pptx，回到第 2 节补装。

---

## 5. 验证安装

### 5.1 重启 AI agent

**完全退出** Claude Code / OpenCode 然后重新打开（skill 是在 agent 启动时扫描的）。

### 5.2 命令行快速测试

打开 PowerShell：

```powershell
cd $env:USERPROFILE\.claude\skills\bjxh-ppt
python scripts\qa.py examples\audit_demo.pptx
```

预期输出：
```
PASS: audit_demo.pptx (5 slides) meets 北京兴华 PPT 规范
```

如果显示 FAIL，按 qa.py 输出的具体 issue 修复（一般是模板路径不对或 Python 包没装全）。

### 5.3 在 AI agent 里测试

在 Claude Code / OpenCode 里说一句话：

> "帮我按北京兴华规范做一份 PPT，主题是 XXX，3 章节，每章 3 页"

如果 skill 装好了，agent 会读 `SKILL.md` 并按规范生成。

如果 agent 完全没反应、不知道有 bjxh-ppt skill，可能是：
- agent 启动时还没扫到 → 重启 agent
- `~/.claude/skills/bjxh-ppt/SKILL.md` 的 YAML frontmatter 缺 `name` 或 `description` 字段 → 重新解压 install
- 文件被杀毒软件隔离了 → 关杀毒或加白名单

---

## 6. 日常使用

每次出 PPT 三步走：

```powershell
# 1. 写大纲（用 examples\audit_demo.md 风格）
notepad my_outline.md

# 2. 生成
python $env:USERPROFILE\.claude\skills\bjxh-ppt\scripts\from_outline.py my_outline.md my_output.pptx

# 3. 自检
python $env:USERPROFILE\.claude\skills\bjxh-ppt\scripts\qa.py my_output.pptx

# 4. PowerPoint 打开 my_output.pptx 微调
```

或者最省事的方式 — 直接在 Claude Code / OpenCode 里写大纲，让 agent 调 skill 生成。

---

## 7. 卸载

不想用了就双击 `uninstall.bat`（会保留一份 backup，要彻底删自己手动 rm）。

---

## 8. 常见问题

### Q1: install.bat 提示"python 不是内部或外部命令"
→ 没装 Python，或安装时没勾"Add Python to PATH"。重装 Python 并勾上。

### Q2: qa.py 报"TemplateNotFound: 北京兴华模板.pptx"
→ 模板路径不对。回到第 3 节处理。

### Q3: PowerPoint 打开生成的 .pptx 提示"修复"
→ 模板用了 111 个母版版式，python-pptx 写入时偶尔会触发 PowerPoint 的兼容性修复弹窗，点"是"即可，不影响内容。

### Q4: 生成的 .pptx 字体是 Arial / 默认英文，标题应该是微软雅黑
→ 模板里没指定东亚字体时 PowerPoint 会用系统默认。检查模板路径是否正确（用了错误路径会 fallback 到 python-pptx 默认）。

### Q5: agent 完全不知道有 bjxh-ppt skill
→ 99% 是没重启 agent。完全退出再打开。

### Q6: 装完之后发现哪个文件没装上
→ 重新双击 install.bat，会自动备份旧版再覆盖。

---

## 9. 文件大小提醒

整个安装包约 **40 MB**（主要是 examples/ 里两份示例 .pptx）。如果你只需要脚本不要示例，删掉 `examples/*.pptx` 和 `examples/*.preview.html` 再压缩，可以压到 ~100 KB。

---

## 10. 给 IT 同事的部署提示（如果走组策略 / SCCM 推送）

skill 本身在 `%USERPROFILE%\.claude\skills\bjxh-ppt\`，**不需要管理员权限**。可以用登录脚本 / 计划任务静默部署：

```cmd
:: 登录脚本里加一行（普通用户权限）
xcopy /E /I /Y /Q "\\fileserver\share\bjxh-ppt\*" "%USERPROFILE%\.claude\skills\bjxh-ppt\"
```

Python 和 pip 依赖需要另外走组策略或 chocolatey/scoop 推送。
