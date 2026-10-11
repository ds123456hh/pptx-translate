# pptx-translate

把一个 `.pptx` 的**文字**翻译成中文（或任何语言），产出一份新的 PPT——**排版、字体、颜色、项目符号、图片、动画一律不动**。

```
输入 deck.pptx  ──►  就地替换文字节点  ──►  deck.zh-CN.pptx
```

这不是"重新生成一份 PPT"，而是把原文件当作 ZIP 包，只改写文本节点，因此：

| 项目 | 实测结果（77 页课件） |
| --- | --- |
| 幻灯片数 | 77 → 77 |
| 形状集合差异 | **0** |
| 几何（位置/尺寸）差异 | **0** |
| 样式（字体/字号/加粗/颜色）差异 | **0** |
| 图片与媒体 | 34 个文件**字节一致** |
| 段落数 | 478 → 478（不合并、不新增、不删除） |
| 动画/切换/母版/版式 | 保持不变 |

---

## 为什么排版不会变

**核心做法：就地替换文字节点，绝不重建形状。**

- 每个 run 的 `rPr`（字体、字号、加粗、颜色、超链接）原样保留，只改 `<a:t>` 里的字；
- 手动换行 `<a:br/>`、段落 `pPr`（项目符号、缩进、行距）保留骨架，译文**一行对一行**；
- 同一段里多种格式（加粗词、高亮词）按原 run 长度比例切分，且**每个 run 至少一个字**，不会出现空 run 导致格式丢失；
- 表情符号 😀🛡️ 先替换成 `[[E1]]` 占位符再送翻译，译后还原——模型物理上不可能丢掉它；
- 图表 / SmartArt 文字直接改写 `chart*.xml` / `diagrams/data*.xml`（`<a:t>` 与 `strCache` 里的 `<c:v>`），数值缓存不动。

详见 [docs/usage.md](docs/usage.md)。

---

## 安装

```bash
git clone https://github.com/OWNER/pptx-translate.git
cd pptx-translate
pip install -e .
```

只依赖 `python-pptx` 和 `lxml`。要跑**本地模型**（推荐，零成本、不外传内容）：

```bash
pip install -e ".[local]"      # 加装 torch(CPU) + transformers + huggingface_hub
```

---

## 快速开始

### 方式一：Ollama + GGUF（推荐：零成本、纯本地、内容不出机器）

不需要任何 Python 深度学习依赖——Ollama 自带推理后端，脚本只走 HTTP。

```bash
# 一次性准备：装 Ollama、拉 GGUF、注册模型
ollama create hy-mt-local -f D:\AI-tools\Models\Modelfile

# 之后每次翻译
pptx-translate deck.pptx --provider ollama --ollama-model hy-mt-local
# 产出 deck.zh-CN.pptx
```

默认模型是 [tencent/Hy-MT2-1.8B-GGUF](https://huggingface.co/tencent/Hy-MT2-1.8B-GGUF)
的 Q4_K_M 量化版（约 1.05GB，纯 CPU 可跑）。
**完整傻瓜式步骤见 [docs/local-ollama.md](docs/local-ollama.md)**，含国内镜像下载、
Modelfile、服务验证和冒烟测试。

```bash
# 先验证管线和本地服务都对（表情/换行/术语锁定都查）
python scripts/ollama_smoke_test.py --model hy-mt-local
```

### 方式二：进程内小模型（HuggingFace，需 torch）

不想装 Ollama、或需要批处理时用这条；代价是 torch 约 1GB 依赖：

```bash
pip install -e ".[local]"
export HF_ENDPOINT=https://hf-mirror.com     # 国内可选
pptx-translate deck.pptx --provider local
```

默认模型 [DALabCommunity/Haidass-Translate-143M](https://huggingface.co/DALabCommunity/Haidass-Translate-143M)（143M 参数）。

### 方式三：OpenAI 兼容接口（DeepSeek / OpenAI / 腾讯 tokenhub 等）

```bash
export DEEPSEEK_API_KEY=sk-xxxxxxxx
pptx-translate deck.pptx --provider deepseek

# 任意兼容接口
pptx-translate deck.pptx --provider openai \
    --base-url https://api.example.com/v1/chat/completions \
    --model your-model
```

API Key 按顺序读环境变量，第一个有值的生效：
`DSH_TRANSLATE_API_KEY` → `DEEPSEEK_API_KEY` → `TOKENHUB_API_KEY` → `SILICONFLOW_API_KEY` → `OPENAI_API_KEY`。
**不要把 Key 写进命令行或提交进仓库**——见 [.env.example](.env.example)。
指向 `localhost` 的本地服务会被自动识别为**无需密钥**。

### 方式四：离线 / 人工校对

```bash
pptx-translate deck.pptx --dry-run                    # 先看清单，不调模型
pptx-translate deck.pptx --dump-strings strings.json  # 导出全部原文
pptx-translate deck.pptx --strings-json zh.json       # 用现成译文打补丁
pptx-translate deck.pptx --provider mock              # 不联网，验证全流程
```

---

## 常用参数

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `--provider` | `deepseek` | `ollama` / `local` / `deepseek` / `openai` / `mock` |
| `--ollama-model` | `hy-mt-local` | Ollama 模型名 |
| `--ollama-host` | `http://localhost:11434/v1` | Ollama 地址 |
| `--local-model` | `DALabCommunity/Haidass-Translate-143M` | 进程内模型的 HF 仓库 |
| `--local-device` | `auto` | `auto` / `cpu` / `cuda` |
| `--local-prompt-style` | `auto` | `instruction` / `chat` / `direct` |
| `--base-url` / `--model` | — | 远程接口地址与模型名 |
| `--batch-size` | `1` | 见下文「为什么默认一次只翻一段」 |
| `--concurrency` | `6` | 并发请求数（本地 provider 自动降为 1） |
| `--glossary` | 内置 23 个专有名词 | 强制锁定的术语（JSON 或 `源,译` 每行） |
| `--autofit` | `off` | `off` / `no-titles` / `all`，中文变长时是否缩字号 |
| `--autofit-preview` | 关 | 只报告会溢出的框，不改字号 |
| `--notes` | 关 | 同时翻译演讲者备注 |
| `--cache` / `--no-cache` | `<输出>.cache.json` | 翻译记忆，失败不写入，可断点续跑 |

完整列表：`pptx-translate --help`，或 [docs/usage.md](docs/usage.md)。

---

## 术语锁定是硬约束

用户要求"术语保留英文"时，**不能只写在提示词里求模型听话**。本项目做法是：命中的术语在发出前就被替换成 `[[G1]]` 占位符，模型根本没机会翻译它，返回后再还原。

占位符编号是**全局稳定**的（由术语本身决定，不是按段落位置），否则同一个 `[[G9]]` 在不同段落会代表不同词，还原出来就是错的词。模型可能把 `[[G9]]` 改写成 `[[G9 ]]` 或漏掉一个括号，还原逻辑对此容错。

```bash
# 想让 "prompt injection" 统一译成中文？覆盖内置表即可
printf 'prompt injection,提示注入\nLLM,大语言模型\n' > glossary.csv
pptx-translate deck.pptx --glossary glossary.csv
```

内置表只收录**真正该保留英文的专有名词与缩写**（LLM、OWASP、MITRE ATT&CK、CVE…）。
像 `token`、`agent`、`prompt`、`penetration testing` 这类已有通用中文译法的词**故意不收**——强行留英文只会让句子更难读。

---

## 为什么 `--batch-size` 默认是 1

翻译专用模型（如 `hy-mt2-lite`）会**把相邻两段合并成一句**：给 8 段可能只回 7 行。而且它自己的编号不可信——实测会把自己发出的行重新编号、把标记挪到行尾、甚至丢掉标记。

合并和丢弃在输出上无法区分，所以**任何"按行对位"的批量方案都会把译文错配到隔壁段落**，而且错得很安静：整页读起来通顺，内容却是另一句的。

一次只发一段，从结构上消除歧义。代价只是请求数变多；配合翻译记忆，重跑几乎不花时间。

> 换成通用对话模型（DeepSeek、GPT 等）对多段输出服从度好，可把 `--batch-size` 调到 10~20 提速。

---

## 已知边界

- **图片里的文字翻不了**（截图、导出的图）。报告会列出面积较大的图片及页码，方便人工补。
- **嵌入的 OLE 对象**（`ppt/embeddings/*.xlsx`）不翻译。
- 只支持 `.pptx` / `.pptm`；老的 `.ppt` 请先另存为 pptx。
- 译文由模型生成，安全术语建议人工抽检。

---

## 校验

每次运行都会写一份报告 `<输出>.report.json`，含页数、段落数、翻译覆盖率、未翻译字符串清单与结构校验结果。

更严格的逐形状比对：

```bash
python scripts/verify_pptx.py 原文件.pptx 译后文件.pptx
# shape-set differences / geometry differences / style differences 应全为 0
```

单元测试（不需要网络、不需要模型）：

```bash
python tests/test_roundtrip.py     # 17 项
```

---

## 项目结构

```
pptx_translate/
  cli.py          命令行、流程编排、校验、报告
  core.py         提取/写回、占位符、术语锁定、翻译器、HTTP provider
  local.py        进程内 HuggingFace provider（可选，需 torch）
  fit.py          溢出检测与保守缩放
tests/
  test_roundtrip.py
scripts/
  install_ollama.ps1         静默安装 Ollama（winget，失败回退官方包）
  download_model.ps1         下载 Hy-MT2 1.8B GGUF（镜像回退 + 断点续传）
  ollama_smoke_test.py       验证本地服务与管线（表情/换行/术语）
  verify_pptx.py             源/译结构样式比对
  final_qa.py                交付前质检（占位符残留、段落对齐）
  render_with_font_fallback.mjs  用指定字体回退渲染页面，便于目视
docs/
  usage.md                    完整用法
  local-ollama.md             本地部署傻瓜式指南
examples/
  glossary.csv / strings-zh.json.example
```

## License

[MIT](LICENSE)
