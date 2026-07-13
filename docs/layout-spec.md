# 北京兴华 PPT 版式实测规范

> 本文档是 [SKILL.md §6.2-6.3](../SKILL.md) 的补充，记录封底和封面版式的逐像素测量数据。
> 数据来源：`北京兴华模板.pptx`（2026 合伙人大会精简版，8 页样例，111 个版式）。
> 测量日期：2026-06-18。

---

## 封底（自定义版式）

> 模板第 111 号版式 + 模板示例 slide 7（1-based index 7，即 0-based index 6）的内联形状。
> 2026-06-18 v3 修订：标题位置上移 + 联系方式字号 16pt。

### layout 自带（不需代码补）

| 元素 | 类型 | 位置(in) | 尺寸(in) | 说明 |
|---|---|---|---|---|
| 主标题占位符 | BODY 文本占位符 idx=11 | (3.94, **3.01**) | 5.44 × 0.89 | **模板默认位置**（代码必须重写） |
| 页码 | SLIDE_NUMBER 占位符 idx=10 | (12.07, 6.78) | 0.35 × 0.29 | PowerPoint 自动填 |
| 装饰文字 | TEXT_BOX | (3.44, 3.68) | 6.02 × 0.45 | "诚信铸就品牌 专业创造价值" 14pt |
| 顶部装饰 | PICTURE（红白波浪） | (6.00, 0.00) | 7.33 × 3.01 | — |
| 底部装饰 | PICTURE（红色波浪） | (0, ~6.5) | 13.33 × 1.0 | — |

### slide 层必须补

| 元素 | 类型 | 位置(in) | 尺寸(in) | 说明 |
|---|---|---|---|---|
| 主标题位置覆盖 | 重写 spPr/xfrm | **(3.94, 2.50)** | 5.44 × 0.89 | 上移到 y=2.50in，与 slogan 保持 0.29in 间距 |
| Logo | PICTURE（拷贝自模板 media/image2.png） | (0.94, 0.82) | 1.98 × 0.51 | "北京兴华 XH GLOBAL" |
| 联系方式 | TEXT_BOX | (0.94, 5.99) | 8.0 × 0.9 | 3 行：电话/传真/地址（白字 16pt） |

### 关键规则

- **必须**用 `自定义版式`（不是 `主题-封底页`）
- **只**写 idx=11 placeholder；idx=10 由 PowerPoint 填
- **必须重写 spPr/xfrm**：标题从 (3.94, 3.01) → (3.94, 2.50)
  - 原因：默认位置 y=3.01 + 字号 40pt → 标题底边 ~3.90in，与 slogan 顶部 3.68in 重叠 0.22in
  - 修正：标题底边 = 2.50 + 0.89 = 3.39in，与 slogan 间距 0.29in
- **代码必须 add**：标题位置重写、联系方式文本框、logo 图片
- **不要**写装饰文字（layout 内置，写入会双重显示）
- 北京建筑背景由 layout 自动继承，**不要**在 slide 层重复 add

### 示例代码（`render.py` 中已实现）

```python
set_placeholder_text(
    find_placeholder(slide, 10),
    spec.title,
    size_pt=36, bold=True
)
```

---

## 封面（主题-封面）

> 模板第 39 号版式。2026-06-18 v3 修订：标题 66pt + 副标题 18pt 强制单行。

### layout 自带（不需代码补）

| 元素 | 类型 | 位置(in) | 尺寸(in) | 说明 |
|---|---|---|---|---|
| 灰色水印 | TEXT_BOX | (2.78, 2.20) | 8.03 × 1.01 | "BEIJING XINGHUA" 54pt Arial Black + softEdge |
| 主标题占位符 | BODY 文本占位符 idx=10 | (2.40, 2.62) | 8.54 × 1.18 | 默认 60pt bold |
| 副标题占位符 | BODY 文本占位符 idx=11 | (0.51, 6.10) | 4.04 × 1.39 | 默认 20pt |
| Logo | PICTURE | (0.94, 0.82) | 1.98 × 0.51 | "北京兴华 XH GLOBAL" |
| 装饰 | PICTURE × 多 | — | — | 顶部右上 + 底部红色波浪 |

### 关键规则（v3 修订）

1. **主标题 66pt bold**（不是 28/36/60pt），与水印视觉重叠：
   - 66pt 文字高 0.92in，水印 y=2.20-3.21in，**重叠 0.72in**
   - 28pt 文字仅 0.39in 高，完全脱离水印 → 不合格
2. **副标题 18pt**（不是 28pt），强制单行：
   - 18pt "BEIJING XINGHUA GROUP" ~3.5in 宽 < 框宽 4.04in → 单行 ✓
   - 28pt 文字 ~5.4in 宽 > 框宽 → 换行 2 行 ✗
3. **不要**写水印文字（layout 内置）
4. **不要** add logo（layout 已包含）
5. 标题和副标题用不同占位符（idx=10 vs idx=11）

### 示例代码（`render.py` 中已实现）

```python
if spec.kind == "cover":
    set_placeholder_text(
        find_placeholder(slide, 10), spec.title, size_pt=66, bold=True
    )
    sub_ph = find_placeholder(slide, 11)
    for para in sub_ph.text_frame.paragraphs:
        for run in para.runs:
            run.font.size = Pt(18)
```
