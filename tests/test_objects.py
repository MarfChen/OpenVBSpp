"""Object model tests (Class / Property / Me / CreateObject / Dictionary).

Copyright 2026 Marf Chen(MarfChen)

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
对象模型（Class / Property / Me / CreateObject / Dictionary）测试。

运行：
    python -m unittest tests.test_objects -v
"""

import contextlib
import io
import os
import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import (                     # noqa: E402
    OVPPError, OVPPObjectNotSupportedError, OVPPSyntaxError,
)
from OpenVBSPP.runtime import vbs_run                   # noqa: E402


def out(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<objects>")
    return buf.getvalue().splitlines()


class TestClassBasics(unittest.TestCase):
    def test_fields_and_new(self):
        code = ("Class Foo\n Public x\nEnd Class\n"
                "Set f = New Foo\nf.x = 5\necho f.x")
        self.assertEqual(out(code), ["5"])

    def test_method(self):
        code = ("Class Foo\n Public Function Twice(n)\n  Twice = n * 2\n"
                " End Function\nEnd Class\nSet f = New Foo\necho f.Twice(4)")
        self.assertEqual(out(code), ["8"])

    def test_sub_method_statement_call(self):
        code = ("Class Foo\n Public x\n Public Sub Set_(v)\n  x = v\n End Sub\n"
                "End Class\nSet f = New Foo\nf.Set_ 3\necho f.x")
        self.assertEqual(out(code), ["3"])

    def test_me(self):
        code = ("Class Foo\n Public x\n Public Sub Set_(v)\n  Me.x = v\n End Sub\n"
                "End Class\nSet f = New Foo\nf.Set_ 9\necho f.x")
        self.assertEqual(out(code), ["9"])

    def test_method_calls_sibling_method(self):
        code = ("Class Foo\n Public Function A()\n  A = B() + 1\n End Function\n"
                " Private Function B()\n  B = 41\n End Function\nEnd Class\n"
                "Set f = New Foo\necho f.A()")
        self.assertEqual(out(code), ["42"])

    def test_class_initialize(self):
        code = ("Class Foo\n Public x\n Private Sub Class_Initialize()\n  x = 7\n"
                " End Sub\nEnd Class\nSet f = New Foo\necho f.x")
        self.assertEqual(out(code), ["7"])

    def test_class_terminate_runs(self):
        code = ("Class Foo\n Private Sub Class_Terminate()\n  echo \"bye\"\n End Sub\n"
                "End Class\nSet f = New Foo\nSet f = Nothing")
        # __del__ relies on reference counting, so CPython runs it immediately
        # 中文：__del__ 依赖引用计数，CPython 下会立即触发
        self.assertEqual(out(code), ["bye"])

    def test_field_array(self):
        code = ("Class Bag\n Public items(2)\n Public Sub Fill()\n  Dim i\n"
                "  For i = 0 To 2\n   items(i) = i * 10\n  Next\n End Sub\nEnd Class\n"
                "Set b = New Bag\nb.Fill\necho b.items(2)")
        self.assertEqual(out(code), ["20"])

    def test_field_array_element_assignment(self):
        code = ("Class Bag\n Public items(2)\nEnd Class\nSet b = New Bag\n"
                "b.items(1) = 9\necho b.items(1)")
        self.assertEqual(out(code), ["9"])

    def test_object_method_updates_module_variable(self):
        code = ("total = 0\nClass Adder\n Public Sub Add1()\n  total = total + 1\n"
                " End Sub\nEnd Class\nSet a = New Adder\na.Add1\na.Add1\necho total")
        self.assertEqual(out(code), ["2"])


class TestChainedMemberCalls(unittest.TestCase):
    def test_call_on_returned_object(self):
        code = (
            "Class Child\n"
            " Public Sub Run()\n"
            "  echo 42\n"
            " End Sub\n"
            "End Class\n"
            "Class Parent\n"
            " Public Function GetChild(n)\n"
            "  Set GetChild = New Child\n"
            " End Function\n"
            "End Class\n"
            "Set p = New Parent\n"
            "p.GetChild(1).Run"
        )
        self.assertEqual(out(code), ["42"])

    def test_explicit_call_on_returned_object(self):
        code = (
            "Class Child\n"
            " Public Sub Run()\n"
            "  echo 42\n"
            " End Sub\n"
            "End Class\n"
            "Class Parent\n"
            " Public Function GetChild(n)\n"
            "  Set GetChild = New Child\n"
            " End Function\n"
            "End Class\n"
            "Set p = New Parent\n"
            "Call p.GetChild(1).Run()"
        )
        self.assertEqual(out(code), ["42"])

    def test_chained_call_with_arguments(self):
        code = (
            "Class Child\n"
            " Public Sub Run(n)\n"
            "  echo n\n"
            " End Sub\n"
            "End Class\n"
            "Class Parent\n"
            " Public Function GetChild(n)\n"
            "  Set GetChild = New Child\n"
            " End Function\n"
            "End Class\n"
            "Set p = New Parent\n"
            "p.GetChild(1).Run 42"
        )
        self.assertEqual(out(code), ["42"])

    def test_chained_member_assignment(self):
        code = (
            "Class Child\n"
            " Public value\n"
            "End Class\n"
            "Class Parent\n"
            " Public item\n"
            " Public Function GetChild(n)\n"
            "  Set GetChild = item\n"
            " End Function\n"
            "End Class\n"
            "Set p = New Parent\n"
            "Set p.item = New Child\n"
            "p.GetChild(1).value = 42\n"
            "echo p.item.value"
        )
        self.assertEqual(out(code), ["42"])


class TestProperties(unittest.TestCase):
    def test_get_and_let(self):
        code = ("Class Foo\n Private n\n Public Property Get Value()\n  Value = n\n"
                " End Property\n Public Property Let Value(v)\n  n = v\n End Property\n"
                "End Class\nSet f = New Foo\nf.Value = 7\necho f.Value")
        self.assertEqual(out(code), ["7"])

    def test_property_uses_method(self):
        code = ("Class Foo\n Private n\n Public Property Get Double_()\n"
                "  Double_ = n * 2\n End Property\n Public Sub Set_(v)\n  n = v\n"
                " End Sub\nEnd Class\nSet f = New Foo\nf.Set_ 5\necho f.Double_")
        self.assertEqual(out(code), ["10"])

    def test_property_inside_with(self):
        code = ("Class Box\n Public value\nEnd Class\nSet b = New Box\n"
                "With b\n .value = 42\nEnd With\necho b.value")
        self.assertEqual(out(code), ["42"])


class TestObjectIdentity(unittest.TestCase):
    def test_typename(self):
        self.assertEqual(out("Class Foo\nEnd Class\nSet f = New Foo\necho TypeName(f)"),
                         ["Foo"])

    def test_is_nothing(self):
        self.assertEqual(out("Set f = Nothing\necho f Is Nothing"), ["True"])

    def test_isobject(self):
        self.assertEqual(out("Class Foo\nEnd Class\nSet f = New Foo\necho IsObject(f)"),
                         ["True"])

    def test_createobject_script_class(self):
        code = ("Class Foo\n Public x\nEnd Class\nSet f = CreateObject(\"Foo\")\n"
                "f.x = 3\necho f.x")
        self.assertEqual(out(code), ["3"])


class TestMethodByRef(unittest.TestCase):
    def test_byref_scalar_writeback(self):
        code = ("Class Box\n Public Sub Fill(ByRef a, ByVal n)\n  a = n * 2\n End Sub\n"
                "End Class\nSet b = New Box\nx = 1\nb.Fill x, 5\necho x")
        self.assertEqual(out(code), ["10"])

    def test_byval_does_not_write_back(self):
        code = ("Class Box\n Public Sub Bump(ByVal n)\n  n = n + 100\n End Sub\n"
                "End Class\nSet b = New Box\nx = 1\nb.Bump x\necho x")
        self.assertEqual(out(code), ["1"])

    def test_byref_array_writeback(self):
        code = ("Class Filler\n Public Sub Put(ByRef arr)\n  arr(0) = 99\n End Sub\n"
                "End Class\nDim a(2)\nSet f = New Filler\nf.Put a\necho a(0)")
        self.assertEqual(out(code), ["99"])

    def test_method_byref_in_expression(self):
        code = ("Class Box\n Public Function Bump(ByRef a)\n  a = a + 1\n  Bump = a * 10\n"
                " End Function\nEnd Class\nSet b = New Box\nx = 1\ny = b.Bump(x)\n"
                "echo x & \",\" & y")
        self.assertEqual(out(code), ["2,20"])


class TestInheritance(unittest.TestCase):
    def test_inherited_method_and_field(self):
        code = ("Class Person\n Public name\n Public Function Greet()\n"
                "  Greet = \"hi \" & name\n End Function\nEnd Class\n"
                "Class Student Inherits Person\n Public school\n"
                " Public Function Info()\n  Info = Greet() & \" @\" & school\n"
                " End Function\nEnd Class\n"
                "Set s = New Student\ns.name = \"ann\"\ns.school = \"MIT\"\n"
                "echo s.Info()\necho TypeName(s)")
        self.assertEqual(out(code), ["hi ann @MIT", "Student"])

    def test_method_override(self):
        code = ("Class Base\n Public Function Who()\n  Who = \"base\"\n End Function\n"
                "End Class\nClass Child Inherits Base\n"
                " Public Function Who()\n  Who = \"child\"\n End Function\nEnd Class\n"
                "Set c = New Child\necho c.Who()")
        self.assertEqual(out(code), ["child"])

    def test_subclass_writes_inherited_field(self):
        code = ("Class Base\n Public x\nEnd Class\n"
                "Class Child Inherits Base\n Public Sub SetX(v)\n  x = v\n End Sub\n"
                "End Class\nSet c = New Child\nc.SetX 7\necho c.x")
        self.assertEqual(out(code), ["7"])

    def test_unknown_base_class(self):
        with self.assertRaises(OVPPError):
            out("Class Foo Inherits Missing\nEnd Class")


class TestDictionary(unittest.TestCase):
    def setUp(self):
        self.prelude = ('Set d = CreateObject("Scripting.Dictionary")\n')

    def run_dict(self, body):
        return out(self.prelude + body)

    def test_add_and_count(self):
        self.assertEqual(self.run_dict('d.Add "a", 1\nd.Add "b", 2\necho d.Count'), ["2"])

    def test_item_via_default_member(self):
        self.assertEqual(self.run_dict('d.Add "a", 5\necho d("a")'), ["5"])

    def test_item_assignment(self):
        self.assertEqual(self.run_dict('d("x") = 9\necho d("x")'), ["9"])

    def test_exists(self):
        self.assertEqual(self.run_dict('d.Add "a", 1\necho d.Exists("a")\necho d.Exists("z")'),
                         ["True", "False"])

    def test_remove_and_removeall(self):
        code = ('d.Add "a", 1\nd.Add "b", 2\nd.Remove "a"\necho d.Count\n'
                'd.RemoveAll\necho d.Count')
        self.assertEqual(self.run_dict(code), ["1", "0"])

    def test_keys_and_items(self):
        code = ('d.Add "a", 1\nd.Add "b", 2\necho Join(d.Keys, ",")\necho Join(d.Items, ",")')
        self.assertEqual(self.run_dict(code), ["a,b", "1,2"])

    def test_for_each_iterates_keys(self):
        code = ('d.Add "x", 1\nd.Add "y", 2\nFor Each k In d\n echo k\nNext')
        self.assertEqual(self.run_dict(code), ["x", "y"])


class TestObjectErrors(unittest.TestCase):
    def test_new_unknown_class(self):
        with self.assertRaises(OVPPObjectNotSupportedError):
            out("Set f = New Missing")

    def test_unknown_member(self):
        code = "Class Foo\nEnd Class\nSet f = New Foo\necho f.nope"
        with self.assertRaises((OVPPError, AttributeError)):
            out(code)

    def test_createobject_unknown_progid(self):
        # COM disabled, so the same error appears even with pywin32 installed (no real Excel)
        # 中文：关掉 COM 支持，保证在装了 pywin32 的机器上也是同一条错误（不真的去起 Excel）
        with mock.patch.dict(os.environ, {"OPENVBS_COM": "0"}):
            with self.assertRaises(OVPPObjectNotSupportedError) as ctx:
                out('Set o = CreateObject("Excel.Application")')
        self.assertIn("COM", str(ctx.exception))

    def test_assign_to_me_rejected(self):
        code = ("Class Foo\n Public Sub Bad()\n  Me = 1\n End Sub\nEnd Class")
        with self.assertRaises(OVPPSyntaxError):
            out(code)

    def test_class_inside_procedure_rejected(self):
        with self.assertRaises(OVPPSyntaxError):
            out("Sub P()\n Class Foo\n End Class\nEnd Sub")


class TestVisibility(unittest.TestCase):
    def test_private_field_not_readable_outside(self):
        code = ("Class Foo\n Private n\nEnd Class\nSet f = New Foo\necho f.n")
        with self.assertRaises((OVPPError, AttributeError)):
            out(code)

    def test_private_field_not_writable_outside(self):
        code = "Class Foo\n Private n\nEnd Class\nSet f = New Foo\nf.n = 1"
        with self.assertRaises(OVPPError):
            out(code)

    def test_unknown_member_write_rejected(self):
        code = "Class Foo\n Public x\nEnd Class\nSet f = New Foo\nf.y = 1"
        with self.assertRaises(OVPPError):
            out(code)

    def test_private_method_not_callable_outside(self):
        code = ("Class Foo\n Private Function Secret()\n  Secret = 1\n End Function\n"
                "End Class\nSet f = New Foo\necho f.Secret()")
        with self.assertRaises((OVPPError, AttributeError)):
            out(code)

    def test_public_field_still_works(self):
        code = "Class Foo\n Public x\nEnd Class\nSet f = New Foo\nf.x = 2\necho f.x"
        self.assertEqual(out(code), ["2"])


class TestDefaultMember(unittest.TestCase):
    def test_default_member_is_read_by_let(self):
        code = ("Class Foo\n Public value\n"
                " Public Property Get Item()\n  Item = value\n End Property\n"
                " Attribute Item.VB_UserMemId = 0\nEnd Class\n"
                "Set f = New Foo\nf.value = 42\nx = f\necho x")
        self.assertEqual(out(code), ["42"])

    def test_set_does_not_unwrap(self):
        code = ("Class Foo\n Public value\n"
                " Public Property Get Item()\n  Item = value\n End Property\n"
                " Attribute Item.VB_UserMemId = 0\nEnd Class\n"
                "Set f = New Foo\nf.value = 42\nSet g = f\necho g.value")
        self.assertEqual(out(code), ["42"])


class TestGetObject(unittest.TestCase):
    def test_getobject_script_class(self):
        code = ("Class Foo\n Public x\nEnd Class\nSet f = GetObject(\"Foo\")\n"
                "f.x = 8\necho f.x")
        self.assertEqual(out(code), ["8"])

    def test_getobject_unknown_progid(self):
        # as above: COM is off, so a running local Excel is never contacted
        # 中文：同上：关掉 COM，避免真的去连本机已运行的 Excel
        with mock.patch.dict(os.environ, {"OPENVBS_COM": "0"}):
            with self.assertRaises(OVPPObjectNotSupportedError):
                out('Set o = GetObject("Excel.Application")')


if __name__ == "__main__":
    unittest.main(verbosity=2)
