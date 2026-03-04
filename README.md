<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white" />
  <img src="https://img.shields.io/badge/Flask-Web_UI-green?logo=flask" />
  <img src="https://img.shields.io/badge/License-MIT-yellow" />
  <img src="https://img.shields.io/badge/Engines-5_Supported-orange" />
</p>

<h1 align="center">📊 ExcelTranslator Pro</h1>
<h3 align="center">智能 Excel 多语言翻译工具 — 一键翻译，完美保留格式</h3>

<p align="center">
  上传 Excel → 选择翻译引擎 → 自动翻译全部文本单元格<br/>
  支持 5 大翻译引擎 · 实时进度追踪 · 翻译缓存断点续传 · 双语对照版本生成
</p>

---

## ✨ 为什么选择 ExcelTranslator Pro？

你是否遇到过这些场景？

- 📄 收到一份几百行的俄语/日语/韩语 Excel 报表，需要翻译成中文
- 🔄 手动复制粘贴到翻译网站，再贴回去？太慢了，还容易搞乱格式
- 📐 翻译完发现公式被覆盖了、合并单元格错位了、日期变成乱码了

**ExcelTranslator Pro 一键解决所有问题：**

| 特性 | 说明 |
|------|------|
| 🚀 **5 大翻译引擎** | DeepSeek · OpenAI GPT · Claude · Google 翻译（免费） · 通义千问 |
| 📐 **格式零损失** | 公式不动、合并单元格不乱、样式完整保留 |
| 🧠 **智能跳过** | 自动识别并跳过公式、纯数字、日期、URL、邮箱、代码片段 |
| 🌐 **20+ 语种** | 中/英/日/韩/俄/法/德/西/阿拉伯语等全覆盖 |
| 📑 **双语对照** | 一键生成原文+译文并排对照版本，便于审核 |
| 💾 **翻译缓存** | 相同文本不重复翻译，支持中断后续传 |
| 📊 **实时进度** | 网页端 SSE 实时推送翻译进度，清晰掌握任务状态 |
| 🔍 **Dry-run 分析** | 先预览哪些单元格会被翻译，不花一分钱 API 费用 |

---

## 🖥️ 界面预览

运行后自动打开浏览器，全程可视化操作，无需任何命令行知识：

> **步骤 1** → 上传 Excel 文件（.xlsx / .xlsm）  
> **步骤 2** → 选择翻译引擎、源语言、目标语言  
> **步骤 3** → 实时查看翻译进度  
> **步骤 4** → 下载翻译结果 & 双语对照版本  

---

## 🚀 快速开始

### 1. 环境要求

- **Python 3.8+**（推荐 3.10+）
- 操作系统：Windows / macOS / Linux 均可

### 2. 安装

```bash
# 克隆项目
git clone https://github.com/你的用户名/ExcelTranslator-Pro.git
cd ExcelTranslator-Pro

# 安装依赖
pip install flask openpyxl tqdm
```

### 3. 配置 API Key（按需）

| 引擎 | 环境变量 | 说明 |
|------|---------|------|
| DeepSeek | `DEEPSEEK_API_KEY` | [获取 Key](https://platform.deepseek.com/) |
| OpenAI | `OPENAI_API_KEY` | [获取 Key](https://platform.openai.com/) |
| Claude | `ANTHROPIC_API_KEY` | [获取 Key](https://console.anthropic.com/) |
| Google 翻译 | 无需 Key | 免费引擎，直接使用 |
| 通义千问 | `DASHSCOPE_API_KEY` | [获取 Key](https://dashscope.aliyun.com/) |

你可以提前设置环境变量，也可以在网页界面中手动输入 API Key。

```bash
# 示例：设置环境变量（macOS/Linux）
export DEEPSEEK_API_KEY="你的key"

# 示例：设置环境变量（Windows PowerShell）
$env:DEEPSEEK_API_KEY="你的key"
```

### 4. 运行

```bash
python app.py
```

程序将自动在浏览器中打开 **http://localhost:8686** ，按界面提示操作即可。

> 💡 也可以直接运行 `python excel_translator_pro.py` 使用命令行版本，功能完全相同。

---

## 📁 项目结构

```
ExcelTranslator-Pro/
├── app.py                    # Flask 网页应用（运行此文件即可）
├── excel_translator_pro.py   # 核心翻译引擎（纯后端，也可独立运行）
├── TestFile.xlsx             # 测试用 Excel 文件
└── README.md                 # 本文件
```

---

## 📝 使用提示

- **首次使用？** 项目内附带了 `TestFile.xlsx` 测试文件，可直接上传体验完整流程
- **不确定格式？** 可先使用「仅分析（Dry-run）」功能预览，不消耗任何 API 额度
- **想省钱？** 选择 Google 翻译引擎，完全免费，无需 API Key
- **输入文件建议：** 确保 Excel 文件没有隐藏行、隐藏文字或折叠分组，否则可能影响翻译结果（程序不会修改原始文件，可放心尝试）

---

## ⚠️ 注意事项

- 程序**不会覆盖**你的原始 Excel 文件，翻译结果会保存为新文件
- 使用 LLM 引擎（DeepSeek / OpenAI / Claude / 通义千问）时需要对应的 API Key，会产生少量 API 调用费用
- Google 翻译引擎完全免费，适合预算有限的场景

---

## 📄 License

本项目基于 [MIT License](LICENSE) 开源，欢迎自由使用和二次开发。
