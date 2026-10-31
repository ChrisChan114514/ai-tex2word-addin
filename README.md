
# AI TeX 转 Word 公式

一个源码可用的 Word 加载项，将 AI 输出中的 TeX 定界公式转换成可编辑的 Word 原生公式（OMML）。使用本地 Pandoc/texmath 数学解析引擎，不依赖 MathType，转换过程无需联网。

## 功能

- `\(...\)` 转换为行内公式。
- `\[...\]` 转换为独立显示公式。
- 按定界符决定混排方式：行内公式与前后正文保持在同一段落，不新增回车或手动换行；独立公式才单独成行。正文原有的分段和手动换行会保留，行宽不足时由 Word 正常折行。
- 公式先在隐藏的临时文档中准备，再仅复制不含段落标记的公式范围；正文不导入 XML 段落或文档默认样式，保留原有段前/段后间距、行距、缩进与样式。
- 兼容 JSON/Markdown 转义后的 `\\(...\\)` 和 `\\[...\\]`。
- 支持常见 TeX 矩阵环境，例如 `\begin{bmatrix}...\end{bmatrix}`。
- 兼容 `\ `、`\,`、`\quad` 等常见 TeX 间距命令。
- 使用数学解析引擎处理嵌套分式、上下括注及中文标注、根式、积分求和、矩阵、分段与多行对齐；直接插入 OMML，不再调用 Word 的线性 `BuildUp` 解析器。
- 遇到未知命令、语法错误或解析警告时保留该公式原文，其余公式继续转换；点击“转换错误详情”查看原因。
- 支持转换选中内容或正文全文。
- 使用 `Ctrl+Alt+Shift+E` 快速转换选中内容；如该组合已被占用，加载项不会覆盖原命令。
- 使用 Word 原生公式，可继续编辑、复制和导出。
- 一次转换记录为一个撤销操作，可立即按 `Ctrl+Z` 恢复。
- 默认不转换 `$...$`，避免金额和普通美元符号被误识别。
- 点击按钮后直接转换，不弹出完成对话框。

## 安装

在 PowerShell 中运行：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\tools\Install-AITeXWordAddin.ps1
```

需先安装 [Pandoc](https://pandoc.org/installing.html)，本次验证版本为 **3.9.0.2**。安装脚本会记录可执行文件的绝对路径；也可以使用 `-PandocPath 'C:\路径\pandoc.exe'` 指定路径。Pandoc 更新或移动后重新运行安装脚本。Pandoc 为独立依赖，不随本仓库分发。

安装脚本会：

1. 从 `src/AITeXConverter.bas` 构建宏模板；
2. 嵌入 `src/customUI14.xml` Ribbon 界面；
3. 把生成文件放入 `dist`；
4. 安装到当前用户的 Word `STARTUP` 目录，同时安装 `AITeXEngine` 配套脚本及引擎路径配置；
5. 在未保存的临时文档中执行 Word 自动化测试，检查原生公式结构、失败保留与整次撤销。

VBA 模块使用纯 ASCII 文本，以避开部分中文 Windows/Office 组合在保存含中文 VBA 源码时发生的卡死；Ribbon 按钮仍为中文。构建脚本通过 `AddFromString` 写入模块，以避开新版 Word 中旧 `.bas` 导入器的兼容问题。

构建期间，脚本会在它启动的隐藏 Word 实例中临时卸载 MathType 6.9 加载项，避免 MathType 的受保护旧 VBA 工程导致 `SaveAs2` 死循环；不会删除 MathType，也不影响下次启动时加载。

如果安装时 Word 正在运行，请保存文档并重启 Word。重启后顶部会出现 **AI 公式** 选项卡。
如果 Word 锁住旧模板，安装脚本会保留已生成的 `dist` 构建，并提示先保存、关闭 Word 再重新安装。只构建和测试、不替换已安装模板时可使用 `-BuildOnly`。运行 `dist` 中的模板也需要同目录下的 `AITeXEngine` 配套文件。

## 使用

例如，文档中包含：

```text
小信号模型包含 \(C_{be},C_{bc},C_{ce}\)。

结电阻与电容为 \(R\_\pi,C\_\pi\)。

\[f_T=\frac{g_m}{2\pi(C_{be}+C_{bc})}\]
```

打开 **AI 公式**，选择“转换选中内容”或“转换全文”。

也可以先选中一段文字，再按 `Ctrl+Alt+Shift+E`。加载项在 Word 启动时检查快捷键，仅当该组合尚未分配时才注册，因此不会覆盖 Word、用户或其他加载项已有的快捷键。

## 兼容性

加载项面向 Windows 桌面版 Word 2016/2019/2021/2024 和 Microsoft 365；实际 Office 环境须运行测试验证。语法支持来自 [texmath](https://github.com/jgm/texmath) 的 LaTeX/AMS 数学解析及 OMML 输出，而非 Word 当前选中的 LaTeX/UnicodeMath 输入模式。

常用数学命令和嵌套结构由同一解析器处理，包括 `\underbrace`、`\overbrace`、`\text{中文}`、`\operatorname`、字体、重音、矩阵、`cases`、`aligned`、`gathered` 和 `array`。支持公式内部定义并使用的局部 `\newcommand`。输出适配层会将括注转换为可随操作数宽度拉伸的 OMML 分组字符结构。为延续 AI/Markdown 输入兼容性，数学区域的 `\_` 仍按下标 `_` 处理，目前不保留 `\_` 的字面下划线语义。

**支持范围是 LaTeX/AMS 数学语法，不是完整 TeX 排版系统。** 不执行外部文件输入、系统命令、任意宏包、TikZ、文档布局或自动公式编号；不读取正文中其他位置的宏定义。OMML 与 TeX 的字体、间距和排版能力也不完全相同，不能保证任意 TeX 都有逐像素相同的可编辑 Word 表达。解析器报错或警告时不会退回旧解析器，也不会静默插入原始 TeX 冒充公式。

公式正文通过结构化 Math AST 传给 Pandoc，不作为完整 LaTeX 文档执行。每次转换只启动一个隐藏 PowerShell 引擎进程；每个不同公式的 Pandoc 解析限时 15 秒，同批重复公式复用结果。转换与错误详情均保留在本机。

## 验证与扩展

```powershell
# 无需 Word，验证数学结构、Unicode、局部宏及错误拒绝
.\tools\Test-TexEngine.ps1

# 在真实 Word 中验证原生结构、选区、失败保留和一次撤销
.\tools\Test-AITeXWordAddin.ps1

# 检查正文混排、行内公式不新增换行、独立公式及撤销
.\tools\Test-FormulaLayout.ps1

# 检查段后 2 磅、直接格式/样式、混合段落格式及选区转换
.\tools\Test-FormulaFormatting.ps1

# 另存覆盖用例的 Word 文档与 Word 导出的 PDF，便于检查外观
.\tools\Test-AITeXWordAddin.ps1 -ReviewDirectory "$PWD\dist\review"
```

`tests/formulas.json` 是按语法类别组织的 44 个用例，其中 38 个有效公式、6 个错误/不支持输入。测试会检查分式、根式、上下标、矩阵行数、括注与对齐等 OMML 结构，以及中文和符号的保留；Word 测试还检查一次撤销能否完整还原原文。新增支持应扩展这套语法覆盖，并优先升级/修正解析引擎，避免添加命令替换补丁。该用例集是回归基线，不代表已经穷尽 TeX。

## 卸载

先关闭 Word，然后运行：

```powershell
.\tools\Uninstall-AITeXWordAddin.ps1
```

## 项目结构

```text
src/AITeXConverter.bas          VBA 转换源码
src/Convert-TexToOmml.ps1        本地 TeX → OMML 引擎桥接
src/customUI14.xml              Word Ribbon 定义
tools/Install-AITeXWordAddin.ps1 构建与安装脚本
tools/Build-VbaTemplate.ps1       隔离的 STA/VBA 构建脚本
tools/Test-AITeXWordAddin.ps1    自动化测试
tools/Test-TexEngine.ps1         不依赖 Word 的解析器测试
tools/Test-FormulaLayout.ps1     Word 正文与公式混排测试
tools/Test-FormulaFormatting.ps1 Word 段落格式保留测试
tests/formulas.json             按语法类别组织的覆盖用例
tests/layout.json               行内/独立公式与换行覆盖用例
tools/Uninstall-AITeXWordAddin.ps1 卸载脚本
dist/                           本机构建产物（不纳入版本控制）
```

## 许可证

本项目采用 **AI TeX Word Add-in Educational and Research License 1.0**：

- 个人学习、课堂教学、学术研究和非商业科研可以免费使用、修改及非商业再分发；
- 再分发时必须保留版权声明和完整许可证，并注明所做的实质性修改；
- 公司内部业务使用、收费产品或服务、商业咨询、商业分发、与商业产品捆绑，以及以商业转化为目的的赞助研发，均属于商业使用；
- 任何商业使用都必须事先取得版权所有者的书面许可或单独的商业许可证。

版权所有者及商业授权联系人：[Chen Wenxuan](https://github.com/ChrisChan114514)。

完整条款见 [LICENSE](LICENSE)。由于许可证限制商业用途，本项目属于 **source-available（源码可用）**，不属于 OSI 定义的开源软件。
