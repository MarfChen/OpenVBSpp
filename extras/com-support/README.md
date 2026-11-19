# com-support — OpenVBS++ 的 Windows COM 支持

这是 OpenVBS++ 的 **COM / ActiveX 自动化**支持包。它把 `pywin32` 封装成
一个 over 包，于是 VBScript 脚本里的 `CreateObject("Scripting.FileSystemObject")`
和 `CreateObject("Excel.Application")` 都能直接用，**脚本里不需要任何 import**。

## 为什么是 over 包

- 依赖声明在 `over.json` 里：`"python": {"pywin32": ">=306"}`。
  `over install` 会交给 pip 解析并安装 `pywin32` 及其所有传递依赖。
- 它出现在 `over list` 里，可以 `over info com-support` 查看，也可以
  `over remove com-support` 卸载 —— 包括 `pywin32` 本身。
- `"platforms": ["win32"]` 保证非 Windows 平台上这条 python 依赖会被跳过，
  不会在 Linux/macOS 上装 `pywin32`。

## 三种使用方式

1. **零配置（默认）**：OpenVBSPP 自带这份包的副本
   （`OpenVBSPP/vendor/com-support/`），Windows 上 `CreateObject` 开箱即用。
2. **装进项目**：`openvbs over install --with-com`，或
   `openvbs over add com-support file:extras/com-support`。
   装进去之后 `packages/com-support/` 优先于内置副本，便于钉版本或替换实现。
3. **关闭**：设置环境变量 `OPENVBS_COM=0`，或卸载包并把
   `OVPP_COM_PATH` 指向空目录。

## 脚本里怎么用

```vbscript
Set shell = CreateObject("WScript.Shell")
echo shell.ExpandEnvironmentStrings("%USERPROFILE%")
echo shell.SpecialFolders("Desktop")

Set excel = CreateObject("Excel.Application")
excel.Visible = False
Set book = excel.Workbooks.Add()
book.Sheets(1).Cells(1, 1).Value = "Hello"
book.SaveAs "C:\temp\demo.xlsx"
excel.Quit
```

`CreateObject` 拿到的对象支持：

- `obj.Prop` 读属性、`obj.Prop = x` 写属性；
- `obj.Method(arg)` 调用方法，也支持带参数的属性（`shell.SpecialFolders("Desktop")`）；
- 作为 `For Each` 的集合使用（`obj.Count`、`obj.Item(i)`、`For Each item In obj`）；
- `Set x = Nothing` 释放引用（`Quit` 由脚本自己调用）。

> 注意：OpenVBS++ 的输出函数是 `echo`，不是 `WScript.Echo`。

## 环境变量

| 变量 | 作用 |
| --- | --- |
| `OPENVBS_COM=0` | 彻底关闭 COM 支持（`CreateObject` 报错说明原因） |
| `OVPP_COM_PATH` | 指定 com-support 目录，优先级最高 |

## 已知限制

- 只支持后期绑定（`IDispatch`）；不支持 `CreateObject` 的远程激活参数。
- 事件（`WScript.CreateObject(..., "Event_")`）尚未实现。
- 32/64 位必须一致：`pywin32` 位数要和解释器一致，否则 `DispatchEx` 会报
  `Invalid class string`。
