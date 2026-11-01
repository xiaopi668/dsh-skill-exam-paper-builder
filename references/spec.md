# exam_spec.json 字段参考

最小可运行例子见 `assets/demo_linux_exam.json`（Linux 卷，23 题 / 150 分 / 4 页）。

## meta（必填：subject）

| 键 | 默认 | 说明 |
|---|---|---|
| exam_title | 2026年普通高等学校招生全国统一考试 | 封面第一行，19 pt 黑体 |
| subject | — | 科目名，如 "Linux系统架构与工程实践" |
| paper_type | 综合卷 | 子标题括号内文字 |
| paper_code | EXAM-2026 | 页眉"科目代码" |
| card_type | A | 页眉"试卷类型"，也印在答题卡上 |
| card_type_label | 试卷类型 | 页眉左标签 |
| confidential | 绝密★启用前 | 页眉左文字 |
| duration_min | 120 | 用时 |
| total_score | — | 满分，封面与答题卡标题用 |
| question_count | — | 题数，封面用 |
| pages | 4 | 起始页数；导出 PDF 后自动校准 |
| file_stem / card_stem / key_stem | 试卷·答题卡·参考答案与评分标准 | 输出文件名 |
| key_header | 内部资料　评卷与讲评使用 | 答案卷左上角小字 |

## notice[] / card_notice[]

试卷注意事项、答题卡注意事项，逐条渲染；`notice[0]` 加粗（通常写 "注意事项："）。

## sections[]

| 键 | 说明 |
|---|---|
| type | `choice` 选择题（2×2 选项表）／`fill` 填空／`short` 简答／`essay` 综合 |
| heading | "选择题"、"填空题"……（脚本自动加"一、二、三、"） |
| instruction | "本题共 10 小题，每小题 5 分，共 50 分。……"，与 heading 用"："相连，整段加粗 |
| card_heading | 答题卡上的分区标题，缺省取 heading |
| card_score | 答题卡标题括号内分值说明 |
| points_each | 每题分值，答题卡生成 "（8 分）" 用 |
| questions[] | 见下 |

`questions[]`：

- `stem`（必填）：整段题干，含题号，如 `"16．简述 fork 与 execve 的分工……"`
- `options`（choice 必填）：4 项，不带 `A．` 也行
- `score`：答题卡角标，如 `"12 分"`
- `lines`：该题答题框行数，覆盖 `card.lines`

## answer_key

```jsonc
"answer_key": {
  "choice_heading": "一、选择题（每小题 5 分，共 50 分）",
  "choice": ["B", "A", "ABD"],          // 与题号顺序一致
  "choice_note": "【评分说明】…",
  "sections": [
    { "heading": "二、填空题（每小题 4 分，共 20 分）",
      "items": [ { "no": "11．", "text": "fd；cwd。", "points": "每空 2 分。" } ] },
    { "heading": "四、综合应用题（共 48 分）",
      "items": [ { "no": "20．（10 分）", "lead": "参考步骤：",
                   "code": "pvcreate /dev/sdd\n...", "text": "……", "points": "【评分】…" } ] }
  ]
}
```

`items[]` 渲染顺序：`no`（加粗）→ `lead` → `text` → `code`（灰底等宽块）→ `points`（9.5 pt 评分细则）。

## card

| 键 | 默认 | 说明 |
|---|---|---|
| theme | C00000 | 答题卡装饰框颜色（十六进制） |
| choice_per_row | 5 | 选择题每行题数 |
| choice_hint | 选择题填涂区… | 选择题区标题行文字 |
| tip | 请在各题目的答题区域内作答…… | 页脚红框提示 |
| absent_note | 缺考标记（由监考员填涂）… | 缺考标记行 |
| lines.fill/short/essay | 1/6/9 | 各题型默认答题行数 |
