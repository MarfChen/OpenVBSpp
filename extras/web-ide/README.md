# web-ide（over 包）

把 OpenVBS++ 自带的简易网页 IDE 打成一个 `over` 包，用包管理器安装和分发。

## 结构

```
web-ide/
├── over.json     # 包清单：name / version / entry / description
└── main.vbs      # 入口：调用运行时桥 `Ide.Start <port>`
```

## 本机直接运行

```powershell
.\venv\Scripts\python.exe cli.py extras\web-ide\main.vbs
# 或指定端口
.\venv\Scripts\python.exe cli.py extras\web-ide\main.vbs 9000
```

启动后浏览器打开 `http://127.0.0.1:7860/`。

## 打包并登记进本地 registry

```powershell
# 1) 打包成 .ovppkg
.\venv\Scripts\python.exe cli.py over pack extras\web-ide

# 2) 登记进 ~/.openvbspp/over/registry.json（无需服务器）
.\venv\Scripts\python.exe cli.py over publish extras\web-ide-1.0.0.ovppkg

# 3) 搜索 + 安装
.\venv\Scripts\python.exe cli.py over search web-ide
.\venv\Scripts\python.exe cli.py over add web-ide "file:<仓库路径>\extras\web-ide-1.0.0.ovppkg"
.\venv\Scripts\python.exe cli.py over install
```

## 运行时桥

`main.vbs` 用的是 `OpenVBSPP/ovpprt/ide_bridge.py` 暴露的对象：

| 调用 | 作用 |
|---|---|
| `Ide.Url(port)` | 返回 `http://127.0.0.1:<port>/` |
| `Ide.Open(port)` | 用默认浏览器打开（非阻塞） |
| `Ide.Start(port)` | 启动服务器（阻塞，Ctrl+C 结束） |

这样 IDE 就和其它包一样，用 `over` 安装、用 VBS 入口文件启动，不必记 `openvbs ide` 这个专用子命令。
