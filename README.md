# AI TeX 转 Word 公式

一个轻量、源码可用的 Word 加载项，将 AI 输出中的 TeX 定界公式转换成可编辑的 Word 原生公式（OMML）。不依赖 MathType。

## 功能

- `\(...\)` 转换为行内公式。
- `\[...\]` 转换为独立显示公式。
- 支持转换选中内容或正文全文。
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

安装脚本会：

1. 从 `src/AITeXConverter.bas` 构建宏模板；
2. 嵌入 `src/customUI14.xml` Ribbon 界面；
3. 把生成文件放入 `dist`；
4. 安装到当前用户的 Word `STARTUP` 目录；
5. 在未保存的临时文档中执行自动化测试。

VBA 模块使用纯 ASCII 文本，以避开部分中文 Windows/Office 组合在保存含中文 VBA 源码时发生的卡死；Ribbon 按钮仍为中文。构建脚本通过 `AddFromString` 写入模块，以避开新版 Word 中旧 `.bas` 导入器的兼容问题。

构建期间，脚本会在它启动的隐藏 Word 实例中临时卸载 MathType 6.9 加载项，避免 MathType 的受保护旧 VBA 工程导致 `SaveAs2` 死循环；不会删除 MathType，也不影响下次启动时加载。

如果安装时 Word 正在运行，请保存文档并重启 Word。重启后顶部会出现 **AI 公式** 选项卡。

## 使用

例如，文档中包含：

```text
小信号模型包含 \(C_{be},C_{bc},C_{ce}\)。

\[f_T=\frac{g_m}{2\pi(C_{be}+C_{bc})}\]
```

打开 **AI 公式**，选择“转换选中内容”或“转换全文”。

## 兼容性

加载项面向 Word 2016/2019/2021/2024 和 Microsoft 365 桌面版。实际支持的 TeX 命令取决于 Word 自带的公式解析器；常见上下标、分式、根号、希腊字母、积分和求和均可使用。复杂自定义宏以及部分 `\begin{...}` 环境可能需要先改写为 Word 支持的 LaTeX。

## 卸载

先关闭 Word，然后运行：

```powershell
.\tools\Uninstall-AITeXWordAddin.ps1
```

## 项目结构

```text
src/AITeXConverter.bas          VBA 转换源码
src/customUI14.xml              Word Ribbon 定义
tools/Install-AITeXWordAddin.ps1 构建与安装脚本
tools/Build-VbaTemplate.ps1       隔离的 STA/VBA 构建脚本
tools/Test-AITeXWordAddin.ps1    自动化测试
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
