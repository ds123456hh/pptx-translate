# 商用交付评估与验收流程

## 当前结论

本项目可以作为 **PPT 翻译生产辅助工具**，目前不能把模型输出直接当作已验收的商用成品。程序保护大部分 Open XML 结构，并能发现部分漏译、数字和计量单位问题；它不能证明语义准确，也不能证明所有页面渲染后都没有遮挡。不同领域、模板和模型必须分别验收。

2026-11-06 使用默认本地模型 `hy-mt-local` 运行 `benchmarks/general_en_zh.json`：12 个跨领域短句中，6 个通过术语、字面量和质量规则检查。失败例包括 `net margin` →“净利润”（财务语境应为“净利率”）、`production yield` →“生产产量”（应审为“生产良率”）、`Median request latency` →“平均请求延迟”（丢失“中位”含义）。医学句的 `5 mg/kg` 被改写为中文单位；营销句 `2.4 million` →“240万”语义相同但触发了严格数字写法提示。这个小基准不是总体准确率估计，却足以证明当前本地模型不能单独承担专业译审。自动检查能提示单位变化，不能可靠识别所有语义误译。

可重复运行：

```bash
python scripts/evaluate_model.py --output benchmark-result.json
```

## 商用候选稿流程

1. **先界定范围。** 确认客户是否要求翻译备注、图中嵌入文字、图表、附件和嵌入对象；取得课件及译后成品的处理授权。默认只翻译可编辑文字，备注需 `--notes`。
2. **建立领域术语表。** 用 `--glossary` 明确固定译法，尤其是专业术语、产品名、缩写和单位。医学、法律、财务等内容需要有相应领域知识的审校人。不要把某个领域的术语表用在其他领域。
3. **生成候选稿。** 明确选择本地或远程 provider，并运行 `--qa-mode commercial`。该模式自动运行文字溢出预检；发现漏译、残留占位符、数字或常见计量单位异常、结构损坏等会返回状态码 3，并在报告中列出 `delivery_assessment.blockers`。
4. **核对报告与文件。** `python scripts/final_qa.py source.pptx output.pptx` 检查报告的输入/输出 SHA-256、包结构和残留占位符，避免把旧报告用于新文件。状态码 0 仅表示自动检查没有阻断项，报告仍为 `manual_review_required`。
5. **逐页渲染审校。** 在客户将使用的 PowerPoint 或兼容软件中打开原稿与译稿，逐页检查文本溢出、遮挡、独立答案框、表格边界、图表标签、图片内文字和动画。审校每个专业术语、否定词、条件句、数量、单位、日期与法律义务表述。记录修订和最终批准人。

示例：

```bash
pptx-translate deck.pptx --provider ollama --qa-mode commercial --glossary terms.csv
python scripts/final_qa.py deck.pptx deck.zh-CN.pptx
```

商用模式要求显式指定 `--provider`，避免默认接口在不知情时接收客户内容。`--provider deepseek`、`--provider openai` 会把待翻译文字发往指定接口。处理客户资料时，应按客户的数据处理要求选择 provider，并核实接口的留存与使用条款。缓存和质量报告会在本地存储原文/译文；按合同要求选择受控路径、设置访问权限与保留期限，必要时使用 `--no-cache`。

## 验收基线

| 维度 | 机器检查 | 交付前人工检查 |
| --- | --- | --- |
| 文件完整性 | 页数、媒体哈希、动画节点、包部件、报告文件哈希 | 用目标软件打开，实际播放关键动画 |
| 文本覆盖 | 可编辑文字的漏译、空译、残留标记 | 图片文字、扫描页、嵌入对象、备注范围 |
| 内容准确 | 数字、常见 SI 单位、序号、可疑新增英文 | 全文语义、术语、专名、法律和医疗含义 |
| 视觉排版 | 启发式溢出预检、自动行高表格提示 | 逐页检查换行、遮挡、答案框和图表 |

每次更换模型、提示词或重要代码后，需使用不同领域的授权样本做盲审：至少覆盖教育、医疗、法律、财务、工程、市场材料，以及长段落、表格、图表、图片文字、备注和混合语言页面。分别记录术语准确率、数字/单位保留、漏译数、逐页排版问题及人工修改时间。没有这组数据，不能声称“各种 PPT 都达标”。

## 已知边界

- 只支持 `.pptx` / `.pptm`；旧 `.ppt` 需要先转换。图片、截图、嵌入 OLE 附件中的文字不会自动翻译。报告中的图片候选只是面积筛选，不是 OCR。
- XML 几何尺寸不变，不代表视觉位置与换行不变。独立定位的填空答案、复杂动画和特殊字体尤其需要逐页检查。
- 默认本地模型体积较小，专业领域术语和复杂条件句可能误译；质量告警不能替代人工审校。
- 输出是候选稿，`delivery_assessment.client_ready` 始终为 `false`。客户交付批准应由独立审校流程记录。

## 许可证核对

项目源码为 MIT；当前 README 指向的 [Hy-MT2-1.8B-GGUF 模型](https://huggingface.co/tencent/Hy-MT2-1.8B-GGUF/blob/main/LICENSE.txt)与 [Haidass-Translate-143M 模型](https://huggingface.co/DALabCommunity/Haidass-Translate-143M)标注 Apache-2.0；[Ollama](https://github.com/ollama/ollama/blob/main/LICENSE)和 [python-pptx](https://github.com/scanny/python-pptx/blob/master/pyproject.toml)标注 MIT。分发产品时仍需保留适用的许可证/版权声明，并按实际使用的模型、字体、Office 软件及 API 合同再次核对。
