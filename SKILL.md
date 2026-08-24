---
name: english-word-extractor
description: 从用户上传的图片中识别英文单词，标注词性（POS）及对应中文释义，并生成一个两列表格的 Excel（.xlsx）文件（单词在第一列，词性与释义在第二列，多词性纵向堆叠在第二列）。当用户上传包含英文单词的图片（如教材截图、单词表、练习题、屏幕截图），并要求"识别单词"、"提取单词"、"整理单词表"、"查词性释义"、"做单词表格"等时，应使用此技能。英文触发词如 extract words from image、identify English words in the picture 也适用。
---

# English Word Extractor（图片英文单词提取 → Excel 表格）

## Overview

根据用户上传的图片，识别图中出现的英文单词，为每个单词给出词性（如 n. / v. / adj. / adv. 等）及对应中文释义，最终生成一个格式规范的两列表格 Excel 文件（.xlsx）交付给用户。识别由多模态视觉能力完成，Excel 生成由本技能内置脚本 `scripts/build_word_table.py`（基于 openpyxl）完成。

## 触发条件

满足以下任一条件即触发本技能：

1. 用户上传一张图片，且消息中包含以下意图（中英文均适用）：
   - "识别/提取/找出图片里的（英文）单词"
   - "整理/生成单词表 / 词汇表"
   - "这个图片里有哪些单词，什么意思/什么词性"
   - "帮我把图片里的单词做成表格"
   - "extract / identify / list the English words in this image/picture/screenshot"
2. 图片内容明显是单词表、词汇书截图、英语练习题等以英文单词为主的材料，且用户希望得到词汇整理结果。

不触发：用户仅要求翻译图片中的句子/段落、描述图片内容、或处理与英文单词无关的图片。

## 执行流程

### Step 1：读取图片

使用多模态视觉能力仔细查看用户上传的图片。注意：

- 逐行扫描，不遗漏任何英文单词。
- 忽略图片中的中文释义原文（释义由本技能重新给出，保证准确性），但可参考其确认单词拼写。
- 若图片清晰度不足导致某个单词无法确认拼写，在回复中单独列出"待确认"项并说明，不要猜测编造，也不要写入 Excel。

### Step 2：识别与查证

对识别出的每个单词：

1. 去重：同一单词只保留一条（忽略大小写差异；若同一拼写有不同词性，合并到同一条目的多词性结构中）。
2. 排序：按单词在图片中出现的先后顺序排列（从上到下、从左到右）。
3. 确定词性及中文释义：给出该单词的常用词性；若该单词有多个常用词性，逐一列出，每个词性对应各自的中文释义。

### Step 3：生成 JSON 中间数据

将识别结果整理为 JSON 数组并保存为临时文件（如工作区下的 `_words.json`），结构如下：

```json
[
  {
    "word": "book",
    "entries": [
      {"pos": "n.", "meaning": "书；书本"},
      {"pos": "v.", "meaning": "预订；预约"}
    ]
  },
  {
    "word": "apple",
    "entries": [
      {"pos": "n.", "meaning": "苹果"}
    ]
  }
]
```

只有一个词性的单词，`entries` 只包含一个元素；没有任何词性信息时，可用 `{"pos": "", "meaning": "释义"}`。

### Step 4：运行脚本生成 Excel

执行本技能目录下的脚本（使用带 openpyxl 的 Python 环境，本机为 `C:\Users\Cheering\.workbuddy\binaries\python\envs\default\Scripts\python.exe`）：

```bash
"C:\Users\Cheering\.workbuddy\binaries\python\envs\default\Scripts\python.exe" \
  "C:\Users\Cheering\.workbuddy\skills\english-word-extractor\scripts\build_word_table.py" \
  <input.json> <output.xlsx>
```

输出文件命名建议 `单词表.xlsx`（或按图片主题命名，如 `Unit1单词表.xlsx`），保存在当前工作区目录下。

脚本会自动完成：

- 两列表格：第一列"单词"，第二列"词性与释义"（格式 `n. 苹果`）。
- 多词性合并：单词只写在第一行，第一列跨词性行纵向合并单元格；第二个及以后的词性释义依次放在第二列中、紧跟在第一个释义下方的单元格里。
- 表头加粗底色、边框、列宽、自动换行等排版。

若该 Python 环境缺少 openpyxl，先执行 `pip install openpyxl` 安装；若脚本执行失败，降级方案：在回复中输出 Markdown 两列表格（多词性时第一列留空模拟合并），并说明未能生成文件。

### Step 5：交付结果

1. 使用 present_files 向用户展示生成的 .xlsx 文件（必须执行，不得只口头描述）。
2. 回复中给出简短摘要：共识别 N 个单词、输出文件路径；有"待确认"单词时一并列出。
3. 单词数不超过 20 个时，可在回复中同时附上 Markdown 表格预览（两列格式与 Excel 一致，多词性时第一列留空）；超过 20 个时只给摘要，避免刷屏。
4. 图片中未发现任何英文单词时，如实告知，不生成空文件。

## 输出规范

- 释义中同一词性下的多个义项用中文分号"；"分隔。
- 词性缩写使用常规词典写法：n.（名词）、v.（动词）、vt.（及物动词）、vi.（不及物动词）、adj.（形容词）、adv.（副词）、prep.（介词）、conj.（连词）、pron.（代词）、num.（数词）、art.（冠词）、interj.（感叹词）。
- 用户明确要求"不要文件、直接看表格"时，可跳过 Step 4，仅输出 Markdown 表格。
- 用户要求 Word 或其他格式时，先按本流程生成表格数据，再转换为对应格式。
