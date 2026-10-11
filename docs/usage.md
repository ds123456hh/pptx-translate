# pptx_translate — PPT 保排版翻译脚本

把一个 `.pptx` 的**文字**翻译成中文，产出新的中文版 PPT，**版式、样式、图片、动画一律不动**。

```
输入 deck.pptx  ──►  逐段替换文字  ──►  deck.zh-CN.pptx
```

---

## 1. 快速开始

**推荐：本地 Ollama（零成本、无需密钥、内容不出本机）**

```powershell
python -m pptx_translate "D:\课件\1 -LLM PENTEST.pptx" `
    --provider ollama --ollama-model hy-mt-local
# 产出：D:\课件\1 -LLM PENTEST.zh-CN.pptx
```

部署步骤见 [local-ollama.md](local-ollama.md)。先跑冒烟测试更稳妥：

```powershell
python scripts\ollama_smoke_test.py --model hy-mt-local
```

**或者：远程 OpenAI 兼容接口**

```powershell
$env:DEEPSEEK_API_KEY = "sk-xxxxxxxx"
python -m pptx_translate "D:\课件\1 -LLM PENTEST.pptx"          # 默认就是 deepseek

# 任意兼容接口（Key 从环境变量读，不要写进命令）
python -m pptx_translate "D:\课件\1 -LLM PENTEST.pptx" `
    --provider openai `
    --base-url "https://api.example.com/v1/chat/completions" `
    --model "your-model"
```

API Key 按以下顺序从环境变量读取，第一个有值的生效：
`DSH_TRANSLATE_API_KEY` → `DEEPSEEK_API_KEY` → `TOKENHUB_API_KEY` → `SILICONFLOW_API_KEY` → `OPENAI_API_KEY`。

它做什么：

1. 打开 pptx，扫描全部可翻译文字（文本框、占位符、表格、组合形状、图表/图示文字）；
2. 逐条发给翻译模型，带缓存、并发、重试；
3. 把译文**就地写回原来的文本 run**，只换字，不换排版；
4. 保存新文件，并输出一份 JSON 报告（翻译了多少、跳过了什么、结构校验结果）。

## 2. 常用命令

```powershell
# 先看清单，不调用 API：多少页、多少段、多少字符，还列出前 40 条原文
python -m pptx_translate deck.pptx --dry-run

# 只导出所有待翻译字符串，交付给人工/其他工具翻译
python -m pptx_translate deck.pptx --dump-strings strings.json

# 用现成译文打补丁（不调用 API，适合人工校对后的稿子）
python -m pptx_translate deck.pptx --strings-json strings-zh.json

# 用 OpenAI / 通义 / Kimi 等任意兼容接口
python -m pptx_translate deck.pptx --provider openai `
    --base-url https://api.openai.com/v1 --model gpt-4o-mini

# 先预览哪些框会溢出（不改字号）
python -m pptx_translate deck.pptx --autofit-preview

# 中文变长导致溢出时，自动把正文缩小到放得下（标题保持原字号）
python -m pptx_translate deck.pptx --autofit no-titles

# 顺带翻译演讲者备注
python -m pptx_translate deck.pptx --notes

# 离线自检：不联网，验证整条流程（把每段用【】包起来）
python -m pptx_translate deck.pptx --provider mock
```

### 全部参数

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `input` | 必填 | 源 `.pptx` / `.pptm` |
| `-o, --output` | `<名字>.zh-CN.pptx` | 输出路径 |
| `--provider` | `deepseek` | `ollama` / `local` / `deepseek` / `openai` / `mock` |
| `--ollama-model` / `--ollama-host` | `hy-mt-local` / `http://localhost:11434/v1` | 本地 Ollama 服务，见 [local-ollama.md](local-ollama.md)，**无需密钥** |
| `--local-model` / `--local-device` / `--local-prompt-style` | `DALabCommunity/Haidass-Translate-143M` / `auto` / `auto` | 进程内 HF 模型（需 `pip install -e ".[local]"`） |
| `--base-url` | `https://api.deepseek.com/v1` | 接口地址，可直接给完整的 `/chat/completions` |
| `--api-key` | 读环境变量 | **建议用环境变量**，命令行传参会留在 shell 历史里 |
| `--model` | `deepseek-chat` | 模型名 |
| `--target-language` | `Simplified Chinese (简体中文)` | 目标语言 |
| `--batch-size` | `1` | 每次请求多少段。**默认 1**，原因见第 5 节 |
| `--concurrency` | `6` | 并发请求数 |
| `--timeout` / `--retries` | `120` / `4` | 单请求超时 / 单条重试次数 |
| `--keep-terms` | 内置术语表 | 提示词里要求保留英文的术语，逗号分隔 |
| `--glossary` | 内置 24 个专有名词 | 强制锁定的术语表（JSON 对象或 `源,译` 每行），会覆盖内置项 |
| `--no-glossary` | 关 | 不锁定任何术语 |
| `--strings-json` | 无 | 预翻译映射，给出后完全不调 API |
| `--cache` / `--no-cache` | `<输出>.cache.json` | 翻译记忆，中断可续跑 |
| `--autofit` | `off` | `off` / `no-titles` / `all` |
| `--autofit-preview` | 关 | 只报告会溢出的框，不改字号 |
| `--notes` | 关 | 同时翻译备注（默认不翻） |
| `--dry-run` | 关 | 只列清单 |
| `--report` | `<输出>.report.json` | 报告路径 |
| `--dump-strings` | 无 | 导出所有原文 |
| `-v` / `-q` | 中 | 详细 / 安静 |

## 3. 为什么排版不会变

**做法：就地替换文字节点，绝不重建形状。** 具体保障：

| 内容 | 处理方式 |
| --- | --- |
| 位置/尺寸/旋转、层级 | 完全不碰（几何校验 0 差异） |
| 字体名、字号、加粗、颜色、项目符号、缩进、行距 | 只改 run 的文字，`rPr` 原样保留 |
| 段落数、手动换行 `<a:br/>` | 保留骨架：译文一行对一行，绝不合并/新增/删除段 |
| 同一段里多种格式（加粗词、高亮词） | 译文按原 run 长度比例切分，且**每个 run 至少一个字**，不会出现空 run 丢格式 |
| 图片、媒体、母版、版式、主题 | 字节级不变（校验已确认 34 个媒体文件哈希一致） |
| 动画、切换、备注、超链接 | 保留 |
| 图表 / SmartArt 文字 | 直接改写 `chart*.xml` / `diagrams/data*.xml` 里的文本（`<a:t>` 与 `strCache` 里的 `<c:v>`），数值缓存不动 |
| 表情符号 😀🛡️ | 翻译前替换成 `[[E1]]` 占位符，译后还原，绝不会丢或被乱译 |

**表情和特殊符号**走占位符机制；**行内换行** `\x0b` 会被转成合法的 XML 换行，视觉一致。

## 4. 关于中文变长（`--autofit`）

中文字形约为 1em 宽，英文小写约 0.5em，所以同样内容中文可能更宽。默认 `--autofit off`：**只翻译，绝不动字号**。若某些固定尺寸文本框放不下：

```powershell
python -m pptx_translate deck.pptx --autofit-preview      # 先看有多少个框会溢出
python -m pptx_translate deck.pptx --autofit no-titles    # 再决定缩小（标题保持原样）
```

- 只缩小**固定尺寸**文本框里的显式字号，最少到原字号的 70%；
- 设置为「随文字调整形状」或「溢出时缩排文字」的框交给 PowerPoint 自己处理，不重复干预；
- 表格单元格可缩，但**不改行高列宽**（改表格几何就是改排版）。

## 5. 译文质量与"为什么默认一次只翻一段"

- 系统提示要求：术语表内的词保留英文（`LLM`、`prompt injection`、`OWASP`、`ATT&CK`…），其余写成通顺的简体中文；数字、版本号、CVE、URL、代码原样保留；
- 段落级翻译（不是逐句），保留原文的行结构；
- 自动保护表情符号与换行符；
- 同一段出现多次的字符串只翻译一次（去重），并写入缓存复用；
- **术语锁定是硬约束，不是"请求"**：命中的词在发出去之前就被替换成 `[[G1]]` 这样的占位符，模型根本没有机会翻译它，译文写回时再还原成英文。不依赖模型的服从度。

### 为什么 `--batch-size` 默认是 1

翻译专用模型（如 `hy-mt2-lite`）会**把相邻的两段合并成一句**。给 8 段，它可能只回 7 行；而它自己的编号也不可信——实测它会把自己发出来的行重新编号，把标记移到行尾，甚至丢掉标记。合并和丢弃在输出上无法区分，所以**任何"按行对位"的方案在批量下都会把译文错配到别的段落上**，而且错得很安静：整页文字看起来通顺，内容却是隔壁那一句的。

一次只发一段，就从结构上消除了这个歧义。代价只是请求数变多，而实测并不慢（446 条字符串、并发 6，约几分钟），加上缓存，反复重跑几乎不花时间。

> 如果你换成通用对话模型（DeepSeek、GPT 等），它对多段 JSON 输出的服从度很好，可以把 `--batch-size` 调回 10~20 提速。

### 实测结论（腾讯 tokenhub + hy-mt2-lite，77 页课件）

| 项目 | 结果 |
| --- | --- |
| 幻灯片 | 77 → 77 |
| 形状集合 / 几何 / 样式差异 | 0 / 0 / 0 |
| 图片与媒体 | 34 个文件字节一致 |
| 段落数 | 478 → 478，无合并、无新增 |
| 译文覆盖 | 446 条唯一字符串，最终仅剩 `LLM`、`GPT – 3.5` 保持英文（本就该保留） |

### 限流与失败重试

翻译接口会成片返回 `429 Too Many Requests`（尤其是并发高的时候）。脚本的处理是：

- 遇到 429 时等待时间自动拉长（6s → 45s），而不是固定短暂重试；
- **失败的条目绝不写入缓存**。这点很关键：如果把"没翻成功、保留原文"也当成译文缓存下来，报告会显示"全部翻译完成"，重跑也永远不会再试它。现在失败条目不进缓存，重跑同一条命令就只会重试这些条目（成功的走缓存，几乎不花时间）；
- 报告里的 `untranslated_strings` 和 `untranslated` 列表会如实告诉你哪些还留在英文。

如果失败较多，降低并发再跑一次即可：

```powershell
python -m pptx_translate deck.pptx --concurrency 2 --retries 6
```

### 保存前的自动清洗

译文写回 XML 前会做两件必要处理，否则会破坏文件或渲染：

1. **ASCII 引号 → 中文引号**。JSON 传输里一个裸 `"` 就会让整行解析失败；改成 `“ ”` 语义相同且永不破坏结构。
2. **换行符 → `_x000B_` 转义**。DrawingML 的手动换行约定是 `_x000B_`，模型输出的 `\v`、`\u000b`、`\\v` 等各种写法都会被归一成真正的换行，再按 PowerPoint 的合法形式写回。

自定义术语（覆盖内置表，作用于包含该词的字符串）：

```powershell
# glossary.csv
# prompt injection,提示注入
# LLM,大语言模型
python -m pptx_translate deck.pptx --glossary glossary.csv
```

## 6. 产物与校验

每次运行都会写一份报告 `<输出>.report.json`，含：页数、段落数、唯一字符串数、字符数、翻译成功数、图表改动数、自动缩放统计、以及**结构校验**结果：

```json
"verification": {
  "slides": 77, "pictures": 34, "media_identical": true,
  "parts_added": [], "parts_removed": []
}
```

`media_identical: true` 表示所有图片/媒体与源文件字节一致；`parts_added/removed` 为空表示包结构没有被增删。

配套自检脚本 `tools_verify.py` 做更严格的逐形状比对：

```powershell
python tools_verify.py 原文件.pptx 译后文件.pptx
# shape-set differences / geometry differences / style differences 应全为 0
```

## 7. 已知边界

- **图片里的文字**（截图、导出的图）无法翻译——需要用原图重做或 PS 覆盖。脚本会在报告里列出面积较大的图片及其所在页，方便你人工过一遍。本套课件 77 页里有 34 张图。
- **嵌入的 Excel/OLE 对象**（`ppt/embeddings/*.xlsx`）不翻译。
- 仅支持 `.pptx` / `.pptm`，不支持老的 `.ppt`（先另存为 pptx）。
- 智能引号、不间断空格等原样保留，便于排版一致。
- 译文由模型生成，涉及安全术语时建议人工抽检。

## 8. 环境要求

- Python 3.9+，仅依赖 `python-pptx` 与 `lxml`（已随 DSH 运行时提供）。
- 纯标准库发 HTTP 请求，无需 `openai` SDK。
- 文件结构：

```
pptx_translate.py            # 入口（等价于 python -m pptx_translate）
pptx_translate/
  __main__.py                # python -m 入口
  cli.py                     # 命令行、流程编排、校验、报告
  core.py                    # 提取/写回/占位符/翻译器/供应商
  fit.py                     # 溢出检测与保守缩放
tests/test_roundtrip.py      # 文字往返单测（9 项，含 <a:br> 与图表）
tools_verify.py              # 源/译结构样式比对
```

## 9. 出问题怎么办

| 现象 | 处理 |
| --- | --- |
| `no API key` | 设 `DEEPSEEK_API_KEY`，或加 `--api-key` |
| 某批失败 | 自动重试 4 次；仍失败的段落保留英文并在报告里计数，可用缓存续跑 |
| 想重译 | 删掉 `<输出>.cache.json`，或加 `--no-cache` |
| 中文被截断 | 先 `--autofit-preview` 查看，再 `--autofit no-titles` |
| 想省 token | 调大 `--batch-size`，或先 `--dump-strings` 人工翻译后用 `--strings-json` |
