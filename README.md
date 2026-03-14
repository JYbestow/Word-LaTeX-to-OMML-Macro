# Word LaTeX to Equation Macro

苦于Gemini等ai输出的word文档无法自动将LaTeX转为OMML公式，只能一个一个手动转格式，故编写了一个自动将LaTeX自动转为OMML的宏

这是一个用于 Microsoft Word 的 VBA 宏工具。它可以自动扫描 Word 文档中由 `$` 包裹的 LaTeX 公式（例如 `$f(0) = R_1$`），并将其自动转换为 Word 原生的专业公式格式（OMML）。

## ✨ 功能特点
* **一键转换**：自动遍历全篇文档，批量转换所有符合格式的公式。
* **原生支持**：不依赖第三方插件，转换为 Word 原生公式对象，方便后续编辑。
* **非贪婪匹配**：精准识别每对 `$` 符号内部的内容，不会误吞普通文本。

## 🚀 如何安装和使用

1. 打开 Microsoft Word，按下 `Alt + F11` 打开 VBA 编辑器。
2. 在左侧的项目窗口中，右键点击 `Normal`（如果想全局应用）或你的当前文档，选择 **插入 (Insert) -> 模块 (Module)**。
3. 将本项目中的 `ConvertLatexToEquation.bas` 代码复制并粘贴到右侧的空白代码窗口中。
4. 关闭 VBA 编辑器。
5. 在 Word 文档中撰写带有 LaTeX 语法的公式，如 `The formula is $a^2 + b^2 = c^2$`。
6. 按下 `Alt + F8`，选择 `ConvertLatexToEquation`，点击 **运行 (Run)**。

## 💡 进阶技巧
为了更方便地使用，建议在 Word 中为该宏分配一个**快捷键**（如 `Ctrl + Shift + M`），或者将其添加到**快速访问工具栏**中。

## ⚠️ 注意事项
* 本宏依赖于 Word 内置的公式解析引擎 (UnicodeMath 解析逻辑)。它对基础的 LaTeX 语法（上下标、分数、根号、希腊字母等）支持极佳。
* 过于复杂的宏包级 LaTeX（如复杂的矩阵或特定环境）可能无法被 Word 原生引擎完美解析。

## 📄 License
This project is licensed under the MIT License.
