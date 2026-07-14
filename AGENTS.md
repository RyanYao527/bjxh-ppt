# bjxh-ppt

AI agent 入口文档。

## 项目概述

北京兴华会计师事务所 PPT 制作规范 — Markdown 大纲 → 品牌合规 .pptx。

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# 跑测试
python -m pytest tests/ -v

# 生成 PPT
python scripts/from_outline.py examples/audit_demo.md output.pptx <template.pptx> --no-chapter-covers

# QA 自检
python scripts/qa.py output.pptx
```

## 关键文件

- `SKILL.md` — 完整规范（中文，330 行）
- `README.md` — 项目说明（英文）
- `examples/` — 示例大纲 + .pptx
- `scripts/` — 工具脚本

## 注意事项

- 不要修改 LICENSE 声明的 BJXH 资产
- 模板文件不在仓库中，需通过 `BJXH_TEMPLATE` 环境变量或 `scripts/config.json` 配置路径
- 代码修改后跑 `python -m pytest tests/ -v` 确认无回归
