# OpenVBS++ demo project

Six short scripts that show the headline features, plus a runner and recording notes.
Everything here is normal OpenVBS++ source — you can open any file in your editor and
run it with `openvbs <file>` (or `python -m OpenVBSPP <file>`).

## What each step shows

| # | Script | Shows |
|---|---|---|
| 1 | `01-legacy.vbs` | VBScript written for `cscript.exe` runs unchanged: `Dim`, `For`, `Function`, `&` |
| 2 | `02-modern.vbs` | `Option Typing`, classes, `Inherits`, `Try/Catch/Finally` with `When` |
| 3 | `03-python.vbs` | `<script src="python-embed" />` → `#python ... #endpython` blocks and `COP` |
| 4 | `04-com.vbs` | Real COM on Windows: `WScript.Shell`, WMI through `WbemScripting.SWbemLocator` |
| 5 | `05-modules.vbs` | `<namespace src="lib/utils.vbs" alias="u">` isolation vs `<script src>` merging |
| 6 | `openvbs over list` | Package manager: bundled `com-support` and `python-embed` are visible |

Total runtime is about 1.1 s, so a single screen recording of the whole run fits a
short GIF.

## Run it

```powershell
# Windows PowerShell 5.1 or PowerShell 7
powershell -NoProfile -ExecutionPolicy Bypass -File examples\demo\run-demo.ps1
```

The runner picks `venv\Scripts\python.exe` when it exists and falls back to `python`
on `PATH`, so it also works from a clean checkout. Section 4 prints a fallback line
instead of failing when COM is unavailable (for example on Linux).

## Recording the GIF

[ScreenToGif](https://www.screentogif.com/) is the least fuss on Windows (portable
build available, no install):

1. Run the demo once **before** recording — the first run pays cold-start cost and
   the timings would look worse than they are.
2. ScreenToGif → **Recorder**: region ≈ `1000×700` around the console window,
   **15 fps** is plenty for text, enable *Capture cursor*.
3. Run the demo again, press <kbd>F7</kbd>/<kbd>F8</kbd> to stop, then in the editor
   use *Reduce frame count* (removes identical frames) and crop the margins.
4. Save as GIF (or MP4 if the target platform accepts it — much smaller and sharper).
   Aim for ≤ 15 s and ≤ 5 MB; a text GIF of six 3-line sections compresses well.

Terminal settings that read better in a GIF: a dark theme, 16–18 pt monospace font,
and `chcp 65001` plus `PYTHONIOENCODING=utf-8` so the Chinese captions in section 4
render correctly.

Optional extra scenes:

* End with `openvbs ide` and a browser window showing the built-in web IDE
  (run a snippet and click *Run*).
* Show `over install` in a scratch project if you want the dependency-graph story.

Where to put the result: `docs/demo.gif`, referenced from `README.md` right below the
badges, and attached to the GitHub Release. Keep the raw `.cast`/`.mp4` out of the
repository.

### Reproducible terminal capture (alternative)

If you prefer text over pixels, [asciinema](https://asciinema.org/) records a
replayable cast and `agg` renders it to GIF deterministically:

```bash
asciinema rec demo.cast --command "powershell -NoProfile -File examples/demo/run-demo.ps1"
agg demo.cast docs/demo.gif
```

## 中文说明

`examples\demo\` 是给演示录像准备的最小项目：六个脚本分别展示「旧 VBScript 原样运行 /
类与 Try-Catch / Python 互操作 / 真实 COM / 多文件命名空间 / 包管理器」，总耗时约 1.1 秒。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File examples\demo\run-demo.ps1
```

录制建议：先跑一遍预热；用 ScreenToGif 录制约 `1000×700` 区域、15 fps、开启光标捕捉；
录完删重复帧、裁剪留白，导出 GIF（或 MP4）。成品放 `docs/demo.gif`，在 README 徽章下方引用，
并附到 Release。终端用深色主题、16–18pt 等宽字体，先 `chcp 65001` 并设置
`PYTHONIOENCODING=utf-8`，这样第 4 段的中文 Windows 版本名才显示正常。
