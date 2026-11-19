'''
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
'''

"""Shared（类静态成员）与 Implements（接口契约）。"""

import contextlib
import io
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP.exceptions import OVPPError        # noqa: E402
from OpenVBSPP.runtime import vbs_run             # noqa: E402


def out(code):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        vbs_run(code, filename="<oop>")
    return buf.getvalue().splitlines()


class TestSharedMembers(unittest.TestCase):
    def test_shared_field_is_class_level(self):
        code = ("Class Counter\n Shared total\nEnd Class\n"
                "Counter.total = 5\necho Counter.total")
        self.assertEqual(out(code), ["5"])

    def test_shared_field_visible_to_instances(self):
        code = ("Class Counter\n Shared total\n Public Sub Bump()\n"
                "  total = total + 1\n End Sub\nEnd Class\n"
                "Counter.total = 10\nSet c = New Counter\nc.Bump\n"
                "echo Counter.total")
        self.assertEqual(out(code), ["11"])

    def test_shared_method_needs_no_instance(self):
        code = ("Class MathX\n Shared Function Double(n)\n  Double = n * 2\n"
                " End Function\nEnd Class\necho MathX.Double(21)")
        self.assertEqual(out(code), ["42"])

    def test_instance_field_is_not_shared(self):
        code = ("Class A\n Public x\nEnd Class\nSet a = New A\nSet b = New A\n"
                "a.x = 1\nb.x = 2\necho a.x & b.x")
        self.assertEqual(out(code), ["12"])


class TestPropertySet(unittest.TestCase):
    CODE = """Class C
    Private stored
    Public Property Get V()
        V = "get:" & stored
    End Property
    Public Property Let V(x)
        stored = "let:" & x
    End Property
    Public Property Set V(y)
        stored = "set:" & TypeName(y)
    End Property
End Class
"""

    def test_let_and_set_are_separate(self):
        code = (self.CODE + "Set c = New C\n"
                "c.V = 5\necho c.V\n"
                "Set c.V = 5\necho c.V")
        self.assertEqual(out(code), ["get:let:5", "get:set:Integer"])

    def test_write_only_property_can_be_assigned(self):
        code = ("Class W\n Public Property Let Only(x)\n  echo \"w \" & x\n"
                " End Property\nEnd Class\nSet w = New W\nw.Only = 3")
        self.assertEqual(out(code), ["w 3"])

    def test_write_only_property_read_raises(self):
        code = ("Class W\n Public Property Let Only(x)\n End Property\n"
                "End Class\nSet w = New W\necho w.Only")
        with self.assertRaises(OVPPError):
            out(code)

    def test_get_still_works_alone(self):
        code = ("Class R\n Public Property Get X()\n  X = 9\n End Property\n"
                "End Class\nSet r = New R\necho r.X")
        self.assertEqual(out(code), ["9"])


class TestImplements(unittest.TestCase):
    SHAPE = ("Class IShape\n Public Function Area()\n End Function\n"
             "End Class\n")

    def test_satisfied_interface_compiles(self):
        code = (self.SHAPE +
                "Class Square Implements IShape\n Public Function Area()\n"
                "  Area = 4\n End Function\nEnd Class\n"
                "Set s = New Square\necho s.Area()")
        self.assertEqual(out(code), ["4"])

    def test_missing_method_is_a_compile_error(self):
        code = (self.SHAPE +
                "Class Square Implements IShape\nEnd Class")
        with self.assertRaises(OVPPError) as ctx:
            out(code)
        self.assertIn("does not implement", str(ctx.exception))

    def test_wrong_arity_is_a_compile_error(self):
        code = (self.SHAPE +
                "Class Square Implements IShape\n Public Function Area(n)\n"
                "  Area = n\n End Function\nEnd Class")
        with self.assertRaises(OVPPError):
            out(code)

    def test_unknown_interface_is_a_compile_error(self):
        with self.assertRaises(OVPPError):
            out("Class A Implements Missing\nEnd Class")

    def test_interface_can_be_inherited(self):
        code = (self.SHAPE +
                "Class Base\n Public Function Area()\n  Area = 1\n"
                " End Function\nEnd Class\n"
                "Class Child Inherits Base Implements IShape\nEnd Class\n"
                "Set c = New Child\necho c.Area()")
        self.assertEqual(out(code), ["1"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
