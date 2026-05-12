<div align="center">
  <h1>FormatX</h1>
  <p><b>专治『AI 复制粘贴综合征』的小小招式</b></p>
  <p>
    <a href="https://github.com/050715Xx/FormatX/blob/main/LICENSE">
      <img src="https://img.shields.io/badge/License-GPL_3.0-blue.svg" alt="License">
    </a>
    <img src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg" alt="Platform">
  </p>
</div>




##   首先说明,这并不是给专家用的复杂工具，而是为每一位受过格式折磨的大学生设计的小小妙招。它能看懂大模型的 LaTeX 语言，并将它们转化为原生 Word 公式。AI 负责思考，FormatX 负责体面。

FormatX 是一个微型的本地化自动化排版与可编辑公式解析工具。



* **作为助手，它是极简的**：不需要配置复杂的排版参数，不需要学习晦涩的 Word 样式逻辑。

* **作为引擎，它是严谨的**：清除脏文本、拆解软回车、识别层级结构、区分中英文字体，并能将纯文本代码转化为真实的数学公式。



我开这个项目的初衷是：让大学生们在忙碌的大学生活中能轻松一点。无论你是在准备考研高数、撰写实验报告，还是整理项目文档，FormatX 都能让排版不再占用您宝贵的创作时间。



## ✨ 功能特性



* **原生数学公式解析** — 像考研复习高等数学时，你可以把 Gemini、豆包、DeepSeek 等 AI 生成的包含 LaTeX 的解题过程直接复制进来。FormatX 会将它们识别并转化为 Word 中真正可以双击修改的原生公式（OMML），分数有横线，积分有上下限。

* **一键规范化排版** — 写实验报告或相关论文时，无需边写边调格式。直接丢入草稿，它会一键变成规范模板：严格的宋体/Times New Roman 中英分流、固定 20 磅行距、首行缩进 2 字符、标题自动加粗。

* **代码与正文隔离** — 计算机专业学生在写文档时，只需用 ``` 将代码包裹，程序便会自动赋予代码等宽字体，保留原有缩进，并确保它不会和普通正文格式互相干扰。

* **底层排版纠错** — 清理“Word 排版灾难”。自动修复因“软回车”（Shift+Enter）导致的字间距异常拉大问题，并自动剔除网页复制带来的零宽字符。

* **纯文本智能嗅探** — 即使你不懂 Markdown，只要写了“第一章”、“1.1”或“摘要”，正则引擎都会自动识别并将其提拔为对应的标题层级。

* **桌面与 Web 双模式** — 提供基于 CustomTkinter 的深色模式桌面 GUI，同时内置 Flask Web 方案，支持作为 PWA 部署在局域网内使用。



## 📸 截图



示例1:



*<img width="1695" height="779" alt="586db53729300dddc5c6973ab7416a61" src="https://github.com/user-attachments/assets/3de818f1-f43b-4948-baff-b4ef45f45954" />*





示例2:



*<img width="1749" height="768" alt="abc7939d185df498eef8061c39a07efb" src="https://github.com/user-attachments/assets/d2494d50-2720-4322-8fcc-d2952cf12350" />*





## 🚀 快速开始



### 下载安装



* **Windows**：从 [Releases](../../releases) 下载最新的 `FormatX.exe`。

    > **Windows SmartScreen 提示**：由于软件未经过昂贵的代码签名，首次运行时可能会被拦截。请点击 **“更多信息” → “仍要运行”**。这是开源软件的正常现象。

* **macOS / Linux**：目前建议直接通过 Python 源码运行。



### 首次运行



1. 下载 `.exe` 后直接双击打开，无需安装过程。

2. 在界面中点击“浏览”选择文件，或点击“文本输入”粘贴 AI 复制的解答。

3. 点击“一键开始排版”，完成后即可在同目录下获取格式化后的 `.docx` 文档。

---

### 🌟 共同进化

**FormatX** 还在不断进化中。
* 如果你觉得它还不错，欢迎点个 **Star** 收藏，这是对开发者最大的鼓励！
* 如果你有更好的想法，或者遇到了奇怪的 Bug，欢迎随时提 [Issue](../../issues)。

让我们一起，让排版变得更体面。



## 🏗️ 架构

```text

format_conversion.py   # 核心排版引擎（段落拆分、角色判定、字体约束、格式重写）

latex_to_omml.py       # 公式解析库（将 LaTeX token 流转化为 Word 底层 XML 树）

gui_main.py            # 桌面端 GUI（基于 CustomTkinter，多线程调度）

web_app.py             # 局域网 Web 服务（基于 Flask，支持 PWA 桌面化）

2.ico / 2.png          # UI 图标与视觉资产


