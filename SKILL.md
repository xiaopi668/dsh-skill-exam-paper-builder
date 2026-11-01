---
name: exam-paper-builder
description: 一键生成中文试卷套件（试卷 + 高考式标准答题卡 + 参考答案与评分标准），从一份 JSON 配置渲染 DOCX 并导出 PDF。用于"出一套卷子""写份试卷""生成答题卡/答案"这类请求。
whenToUse: 用户要求出卷、组卷、生成试卷/答题卡/参考答案，或要把某学科题目排成高考或校内考试格式时加载本技能。
---

# 试卷一键生成（exam-paper-builder）

把**内容**与**版式**分开：模型只负责按 JSON 写题目，脚本负责把 4 页试卷、4 页标准答题卡、参考答案排版成可直接打印的 DOCX/PDF。

## 三步流程

1. **写配置**：在任务工作区新建 `exam_spec.json`（照 `assets/demo_linux_exam.json` 改内容）。题干文本自己写，编号写进题干（如 `"11．……"`），选项不要带 `A．` 前缀（脚本会自动加，带了也不重复）。
2. **跑脚本**：用 `load_workspace_dependencies` 返回的 Python 执行（系统 python 通常没有 python-docx）：

   ```bash
   "<python>" <skill>/scripts/build_exam.py --spec exam_spec.json --out exam_out
   "<python>" <skill>/scripts/build_exam.py --demo  --out exam_out     # 内置示例自检
   ```

   产物：`<file_stem>.docx/.pdf`、`答题卡.docx/.pdf`、`参考答案与评分标准.docx/.pdf`。
   只要 DOCX 加 `--no-pdf`；只出其中几份用 `--only paper,card,key`。
3. **核对再交付**：按 `references/layout.md` 的"渲染核对"一节抽页渲染（第 1 页必看），确认无空标题、无跨页断框、页码为"第 X 页（共 N 页）"，然后 `present` 文件。

## 一键闭环（推荐）

```bash
"<python>" <skill>/scripts/build_exam.py --spec exam_spec.json --out exam_out && \
"<node>" "<libreoffice-kit-cli>" render --input exam_out/<file_stem>.pdf --output-dir .preview --pages "1"
```

## 配置要点

| 字段 | 作用 |
|---|---|
| `meta.exam_title / subject / paper_type` | 封面标题三行 |
| `meta.paper_code / card_type` | 页眉"试卷类型：A ｜ 科目代码：…" |
| `meta.pages` | 起始页数提示；导出 PDF 后脚本会自动校准为真实页数 |
| `meta.total_score / question_count / duration_min` | 封面信息行，也用于答题卡标题 |
| `notice[]` | 试卷注意事项框（第一条自动加粗） |
| `sections[]` | 题型块，`type` 取 `choice`/`fill`/`short`/`essay`；`heading`+`instruction` 拼成"一、选择题：本题共…" |
| `sections[].questions[]` | `stem` 必填；`choice` 需 `options`（4 个，2×2 排版）；非选择题可给 `score`、`lines`（答题卡答题行数） |
| `answer_key.choice[]` | 选择题答案，自动排成 4 列答案表（超过 5 题分左右两栏） |
| `answer_key.sections[].items[]` | `no`/`text`/`points`，可带 `lead`+`code` 输出灰底代码块 |
| `card.theme / choice_per_row / lines` | 答题卡主题色（默认暗红 C00000）、选择题每行题数、各题型默认行数 |

完整字段见 `references/spec.md`。

## 版式约束（踩过的坑，别重犯）

- **不要用段落边框画答题横线**：Word/LO 会把相邻同色边框合并，内线消失。用"逐段底边 + 颜色微差"（脚本已实现 `_ruled`），或表格单元格。
- **表格宽度只设一次**：重复追加 `w:tblW`/`w:tblLayout` 会让 LibreOffice 误判表格尺寸，出现整页塌陷或空页（`table_width` 已复用现有节点）。
- **答题框必须 `cantSplit`**：否则一个框被拆到两页，扫描端无法识别。
- **页数自校准**：封面"本试卷共 N 页"必须等于实际页数，否则一眼假。脚本导出 PDF 后用 `/Type /Page` 计数回写并重建，最多三轮。
- **中文字体**：每个 run 都要设 `w:eastAsia`，只设 `font.name` 中文会掉成默认字体。

## 交付

`present` 试卷 DOCX + PDF、答题卡 DOCX、参考答案 DOCX（最多 4 个，其余在正文给路径）。若用户只想要内容不想要版式，跳过本技能直接写 Markdown。

## 本机已验证命令（照抄即可）

```bash
SKILL="$HOME/.dsh/skills/exam-paper-builder"
PY="/home/xiaopi/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3"   # 若无 python-docx，改用 load_workspace_dependencies 返回的路径
NODE="/usr/lib/deepseek-harness-desktop/resources/runtime/primary-runtime/dependencies/node/bin/node"
CLI="/usr/lib/deepseek-harness-desktop/resources/app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js"

# 1) 先跑内置示例确认环境
"$PY" "$SKILL/scripts/build_exam.py" --demo --out exam_out

# 2) 换成自己的题目
"$PY" "$SKILL/scripts/build_exam.py" --spec exam_spec.json --out exam_out

# 3) 抽页核对
"$NODE" "$CLI" render --input exam_out/Linux系统架构与工程实践试卷.pdf --output-dir .preview --pages "1,2"
```

产物命名由 `meta.file_stem / card_stem / key_stem` 决定；脚本退出码非 0 时先看 stderr 的 `TypeError/ImportError`，再按 `references/layout.md` 的症状表排查。
