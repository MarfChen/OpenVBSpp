# Pull Request / 合并请求

## 做了什么 / What this PR does

<!-- 一两句话说明。关联 issue 用 "Fixes #123"。 -->

## 类型 / Type

- [ ] 缺陷修复 / bug fix
- [ ] 新语法或新内建 / new syntax or builtin（**必须**先确认不破坏超集）
- [ ] 性能优化 / performance
- [ ] 文档 / docs
- [ ] 工程化 / tooling、CI、打包

## 提交前自查 / Checklist

- [ ] `python -m unittest discover -s tests` 全绿（当前应是 485 个用例）
- [ ] 新增行为有对应的测试；修复缺陷的 PR 附了能复现该缺陷的测试
- [ ] 没有改动任何**合法 VBScript 程序的现有语义**（超集铁律）
- [ ] 改了 `OpenVBSPP/vendor/com-support/` 的话，`extras/com-support/` 里的同名文件同步了（两者必须逐字节一致）
- [ ] 改了 `ovpprt/` 里的模块的话，`packaging/openvbs.spec` 的 `HIDDEN_IMPORTS` 也补了
- [ ] 版本号五处一致（`pyproject.toml`、`OpenVBSPP/__init__.py`、`ast_nodes/__init__.py`、`ovpprt/__init__.py`、`plugins/__init__.py`），如果改了就说明原因
- [ ] 文件保持 UTF-8（勿用 PowerShell `Get-Content/Set-Content` 批量改源码，会把编码转成 GBK）

## 怎么验证的 / How it was verified

<!-- 命令、截图、基准数据都可以。 -->
