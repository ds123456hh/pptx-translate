# 本地模型部署：Ollama + Hy-MT2 1.8B GGUF

零成本、纯本地、不受云 API 额度限制。全程不需要 Python 深度学习依赖——
Ollama 自带推理后端，脚本只通过 HTTP 调用它。

> 如果你想在自己机器上重跑一遍，`scripts/` 下有三个幂等脚本：
> `install_ollama.ps1`、`download_model.ps1`，以及本页末尾的验证脚本。

---

## 第一步：下载 GGUF 模型

目标文件（已核实仓库真实存在、文件名准确）：

```
仓库：tencent/Hy-MT2-1.8B-GGUF
文件：Hy-MT2-1.8B-Q4_K_M.gguf     约 1.05 GB
```

三选一：

**A. PowerShell 一键下载（推荐，带断点续传和镜像回退）**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File D:\AI-tools\scripts\download_model.ps1
```

脚本按 `hf-mirror.com` → `huggingface.co` 的顺序自动重试，中断后重跑会**接着下**，不会从头开始。

**B. 手动 curl**

```powershell
New-Item -ItemType Directory -Force D:\AI-tools\Models | Out-Null
curl.exe -L --retry 12 --retry-delay 5 --continue-at - `
  --output D:\AI-tools\Models\Hy-MT2-1.8B-Q4_K_M.gguf `
  https://hf-mirror.com/tencent/Hy-MT2-1.8B-GGUF/resolve/main/Hy-MT2-1.8B-Q4_K_M.gguf
```

**C. Hugging Face CLI**

```powershell
pip install -U "huggingface_hub[cli]"
$env:HF_ENDPOINT = "https://hf-mirror.com"
hf download tencent/Hy-MT2-1.8B-GGUF Hy-MT2-1.8B-Q4_K_M.gguf --local-dir D:\AI-tools\Models
```

> 备选仓库（主仓库拉不动时）：`zanish-labs/HY-MT1.5-1.8B-GGUF`（文件名
> `HY-MT1.5-1.8B-Q4_K_M.gguf`）。混元翻译 7B 的 GGUF 在
> `mradermacher/Hunyuan-MT-7B-GGUF`，更大但质量更高，CPU 也能跑。

---

## 第二步：安装 Ollama

```powershell
winget install --id Ollama.Ollama -e --silent
```

或一键脚本（winget 失败会自动改用官方安装包静默安装）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File D:\AI-tools\scripts\install_ollama.ps1
```

安装后**如果 `ollama` 命令找不到**，是 PATH 没刷新，开个新终端，或临时补上：

```powershell
$env:Path += ";$env:LOCALAPPDATA\Programs\Ollama"
ollama --version
```

---

## 第三步：注册为本地模型

**Modelfile**（`D:\AI-tools\Models\Modelfile`，内容只有一行）：

```
FROM ./Hy-MT2-1.8B-Q4_K_M.gguf
```

**创建模型**（必须在 Modelfile 所在目录执行，`FROM ./x.gguf` 是相对路径）：

```powershell
cd D:\AI-tools\Models
ollama create hy-mt-local -f Modelfile
ollama list          # 应看到 hy-mt-local
```

想固定生成参数可以扩展 Modelfile（可选）：

```
FROM ./Hy-MT2-1.8B-Q4_K_M.gguf
PARAMETER temperature 0.2
PARAMETER num_predict 512
```

---

## 第四步：启动服务并验证

```powershell
ollama serve          # 监听 11434；装了桌面版通常已自动在跑，端口占用属正常
```

开另一个终端验证三层：

```powershell
# 1) 服务活着
curl.exe -s http://localhost:11434/api/version

# 2) 模型在列表里
curl.exe -s http://localhost:11434/api/tags

# 3) 端到端翻译（OpenAI 兼容接口，无需密钥）
curl.exe -s http://localhost:11434/v1/chat/completions `
  -H "Content-Type: application/json" `
  -d "{\"model\":\"hy-mt-local\",\"messages\":[{\"role\":\"user\",\"content\":\"将下面的文本翻译成中文，只输出译文：Hello, translate this to Chinese\"}]}"
```

---

## 第五步：用真实 PPT 文本验证管线

上一步只证明模型能说话；这一步验证**表情、手动换行、术语锁定**都没坏：

```powershell
cd D:\AI-tools
python scripts\ollama_smoke_test.py --model hy-mt-local
```

输出应包含 `emoji preserved: True` 和 `soft break preserved: True`，最后打印 `RESULT: PASS`。

然后翻整份课件：

```powershell
python -m pptx_translate "D:\课件\1 -LLM PENTEST.pptx" `
    --provider ollama --ollama-model hy-mt-local
```

### 1.8B 小模型的三个已知毛病（脚本已自动处理）

实测这份 77 页课件时，1.8B 模型暴露了三个稳定出现的问题，都在管线里做了确定性修复，
不依赖模型"听话"：

| 毛病 | 实际现象 | 处理方式 |
| --- | --- | --- |
| **凭空编造占位符** | 原文没有 emoji，却把 `[[E1]]` 写进译文（因为提示词里举了这个例子） | 源码没有的 token 一律删除，并合并因此多出的标点 |
| **括号不配对** | `（ATT&CK）]`、`（ATLAS））。`、`（ATT&CK>）` | 按源文括号数量比对：删多余闭合符、把错字符还原成 `）`；**绝不凭空补括号** |
| **输出繁体字** | `嚴格範圍：無混亂`、`規則`、`誤解` | OpenCC 繁→简（`pip install opencc-python-reimplemented`）；没装则用内置小表 |

另外还会清掉模型给整句加的引号（`“输入：”` → `输入：`），但**源文本身有引号时不动**。

实测速度：**RTX 3050 上约 2.3 秒/段**，整份 446 段约 17 分钟；命中缓存后重跑只要 6 秒。

---

## 常见问题

| 现象 | 原因与解法 |
| --- | --- |
| `connection refused` 10061 | 服务没起来：`ollama serve`，或检查桌面版是否在运行 |
| `model 'hy-mt-local' not found` | 没 `ollama create`，或名字拼错：`ollama list` 对一下 |
| 下载到一半断了 | 直接重跑 `download_model.ps1`，`--continue-at -` 会续传 |
| `ollama create` 报找不到 gguf | 没有 `cd` 到 Modelfile 所在目录（`FROM ./` 是相对路径） |
| 翻译很慢 | 1.8B Q4_K_M 在纯 CPU 上约每段 1~3 秒；先掐表看冒烟测试的单段耗时 |
| 输出是英文没翻译 | 小模型对提示词敏感：试 `--local-prompt-style instruction`，或换 7B 版 |
| 术语没保住英文 | 术语锁定是硬约束，不应发生；若出现请把该段发我 |
| 中文变长导致溢出 | 先 `--autofit-preview` 看有多少框，再决定 `--autofit no-titles` |

## 为什么不用 PyTorch

原始计划是 `torch + transformers` 直接加载 143M 小模型。放弃的原因：

- torch 的 CPU wheel 就要 ~1GB 安装量，而 1.8B GGUF 在 Ollama 里只占 1.05GB **磁盘**、几百 MB 内存；
- Ollama 自带量化推理后端和 KV 缓存管理，纯 CPU 更快，也不需要 CUDA；
- 脚本与模型解耦：换模型只改 `--ollama-model`，不用重装 Python 依赖。

`--provider local`（torch 路线）代码仍保留，供显存充裕或需要批处理的场景；
需要时 `pip install -e ".[local]"` 即可，不影响 Ollama 路线。
