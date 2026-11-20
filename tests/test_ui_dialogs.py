"""
Copyright 2027 Marf Chen(MarfChen)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Tests for the `ui-support` over package and the MsgBox / InputBox dispatcher.

中文：`ui-support` 内置包与 `MsgBox` / `InputBox` 分发层的测试。
"""

import os
import unittest
from unittest import mock

from OpenVBSPP.ovpprt import ui
from OpenVBSPP.runtime import vbs_run


def out(source):
    """Run a snippet and return its output lines.

    中文：运行一段代码并返回输出行。
    """
    import contextlib
    import io

    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer):
        vbs_run(source, filename="<test>")
    return buffer.getvalue().splitlines()


def setUpModule():
    """Force the console path for the whole module.

    中文：整个模块强制走控制台路径，避免测试卡在弹窗上。
    """
    global _SAVED_UI_ENV
    _SAVED_UI_ENV = os.environ.get("OPENVBS_UI")
    os.environ["OPENVBS_UI"] = "console"
    ui.reset_cache()


def tearDownModule():
    """Restore the environment and the dispatcher state.

    中文：恢复环境变量与分发层状态。
    """
    if _SAVED_UI_ENV is None:
        os.environ.pop("OPENVBS_UI", None)
    else:
        os.environ["OPENVBS_UI"] = _SAVED_UI_ENV
    ui.reset_cache()


class TestDialogSpec(unittest.TestCase):
    """The buttons value is split into tk_messageBox options.

    中文：把 buttons 值拆成 tk_messageBox 选项。
    """

    def module(self):
        directory, _origin = ui.resolve()
        self.assertIsNotNone(directory, "the bundled ui-support package is missing")
        return ui._load(directory)

    def test_button_sets(self):
        """All six VBScript button sets map onto their Tcl type.

        中文：六种 VBScript 按钮组合都映射到对应的 Tcl 类型。
        """
        module = self.module()
        expected = {0: "ok", 1: "okcancel", 2: "abortretryignore",
                    3: "yesnocancel", 4: "yesno", 5: "retrycancel"}
        for value, name in expected.items():
            self.assertEqual(module.dialog_spec(value)["type"], name)

    def test_icons(self):
        """The four icon bits map onto Tcl icons.

        中文：四个图标位映射到 Tcl 图标。
        """
        module = self.module()
        expected = {16: "error", 32: "question", 48: "warning", 64: "info"}
        for value, name in expected.items():
            self.assertEqual(module.dialog_spec(value)["icon"], name)
        self.assertEqual(module.dialog_spec(0)["icon"], "")

    def test_default_button(self):
        """vbDefaultButtonN picks the matching button, clamped to the last one.

        中文：vbDefaultButtonN 选中对应按钮，越界时取最后一个。
        """
        module = self.module()
        self.assertEqual(module.dialog_spec(4)["default"], "yes")
        self.assertEqual(module.dialog_spec(4 + 256)["default"], "no")
        self.assertEqual(module.dialog_spec(1 + 256)["default"], "cancel")
        self.assertEqual(module.dialog_spec(1 + 768)["default"], "cancel")
        self.assertEqual(module.dialog_spec(0)["default"], "ok")

    def test_modal_bits_make_it_topmost(self):
        """vbApplicationModal and vbSystemModal request a topmost dialog.

        中文：vbApplicationModal 与 vbSystemModal 让对话框置顶。
        """
        module = self.module()
        self.assertFalse(module.dialog_spec(0)["topmost"])
        self.assertTrue(module.dialog_spec(4096)["topmost"])
        self.assertTrue(module.dialog_spec(8192)["topmost"])

    def test_unknown_button_set_falls_back_to_ok(self):
        """An out-of-range button value behaves like vbOKOnly, as in VBScript.

        中文：按钮值越界时按 VBScript 的行为退回 vbOKOnly。
        """
        module = self.module()
        self.assertEqual(module.dialog_spec(7)["type"], "ok")
        self.assertEqual(module.dialog_spec(None)["type"], "ok")

    def test_return_codes(self):
        """The Tcl result strings map back onto the VBScript codes 1-7.

        中文：Tcl 返回字符串映射回 VBScript 的 1-7。
        """
        module = self.module()
        expected = {"ok": 1, "cancel": 2, "abort": 3, "retry": 4,
                    "ignore": 5, "yes": 6, "no": 7}
        self.assertEqual(module.RESULT_CODES, expected)


class TestActivation(unittest.TestCase):
    """Dialogs are opt-in for interactive command-line runs only.

    中文：只有交互式命令行运行才启用对话框。
    """

    def setUp(self):
        os.environ["OPENVBS_UI"] = "console"
        ui.reset_cache()

    def tearDown(self):
        os.environ["OPENVBS_UI"] = "console"
        ui.reset_cache()

    def test_console_mode_disables_dialogs(self):
        """OPENVBS_UI=console always keeps the console behaviour.

        中文：OPENVBS_UI=console 始终走控制台行为。
        """
        ui.mark_cli_entry()
        enabled, reason = ui.enabled()
        self.assertFalse(enabled)
        self.assertIn("OPENVBS_UI", reason)

    def test_library_entry_is_never_gui(self):
        """Without the CLI marker even an explicit gui request falls back.

        中文：没有 CLI 标记时，即使显式要求 gui 也退回控制台。
        """
        os.environ["OPENVBS_UI"] = "auto"
        ui.reset_cache()
        enabled, reason = ui.enabled()
        self.assertFalse(enabled)
        self.assertIn("command line", reason)

    def test_gui_mode_requires_a_terminal_in_auto(self):
        """auto mode needs stdin/stdout to be a terminal.

        中文：auto 模式要求 stdin 与 stdout 是终端。
        """
        os.environ["OPENVBS_UI"] = "auto"
        ui.reset_cache()
        ui.mark_cli_entry()
        with mock.patch.object(ui, "_is_a_terminal", return_value=False):
            enabled, reason = ui.enabled()
        self.assertFalse(enabled)
        self.assertIn("terminal", reason)

    def test_ci_blocks_auto_mode(self):
        """A CI run never waits for a click.

        中文：CI 运行永远不会等待点击。
        """
        os.environ["OPENVBS_UI"] = "auto"
        os.environ["CI"] = "true"
        ui.reset_cache()
        ui.mark_cli_entry()
        try:
            with mock.patch.object(ui, "_is_a_terminal", return_value=True):
                enabled, reason = ui.enabled()
        finally:
            os.environ.pop("CI", None)
        self.assertFalse(enabled)
        self.assertIn("CI", reason)

    def test_console_builtins_still_print(self):
        """Console mode keeps the historical MsgBox / InputBox behaviour.

        中文：控制台模式保持 MsgBox / InputBox 的历史行为。
        """
        self.assertEqual(out('x = MsgBox("hello")\necho x'), ["hello", "1"])
        with mock.patch("builtins.input", return_value="typed") as reader:
            self.assertEqual(out('echo InputBox("prompt")'), ["prompt", "typed"])
        reader.assert_called_once_with()


class TestBundledPackage(unittest.TestCase):
    """The bundled package is present, loadable and describes itself.

    中文：内置包存在、能加载、并能自我描述。
    """

    def test_manifest(self):
        """over.json declares the feature and its capability.

        中文：over.json 声明了包类型与能力。
        """
        import json

        directory, origin = ui.resolve()
        self.assertEqual(origin, "bundled")
        manifest = json.loads((directory / "over.json").read_text(encoding="utf-8"))
        self.assertEqual(manifest["name"], "ui-support")
        self.assertEqual(manifest["kind"], "feature")
        self.assertEqual(manifest["capabilities"], ["ui"])

    def test_module_loads_and_reports_tk(self):
        """The module loads and reports its Tk version.

        中文：模块能加载并报告 Tk 版本。
        """
        directory, _origin = ui.resolve()
        module = ui._load(directory)
        self.assertEqual(module.VERSION, "1.0.0")
        text = module.describe()
        self.assertIn("ui-support 1.0.0", text)
        if module.is_available():
            self.assertIn("Tk", text)

    def test_describe_names_the_mode(self):
        """The dispatcher reports console mode while it is forced.

        中文：强制控制台时，分发层如实报告。
        """
        self.assertIn("console", ui.describe())

    def test_extras_mirror_matches_vendor(self):
        """`extras/ui-support` stays byte-identical to the bundled copy.

        中文：`extras/ui-support` 与内置副本逐字节一致。
        """
        from pathlib import Path

        root = Path(__file__).resolve().parent.parent
        vendor = root / "OpenVBSPP" / "vendor" / "ui-support"
        extras = root / "extras" / "ui-support"
        self.assertTrue(extras.is_dir(), "extras/ui-support 应当存在")
        for name in ("ui_support.py", "over.json"):
            with self.subTest(name=name):
                self.assertEqual((vendor / name).read_bytes(),
                                 (extras / name).read_bytes())


if __name__ == "__main__":
    unittest.main()
