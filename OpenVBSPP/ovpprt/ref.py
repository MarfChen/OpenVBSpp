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

ByRef parameters.

VBScript passes procedure arguments by reference by default; Python only passes by
value, so the compiler boxes a ByRef argument in a `Ref`:

    create:  _r = Ref(x)
    inside:  every read and write of x becomes x.v
    after:   x = _r.v            (write the value back)

That is what makes `Sub Swap(ByRef a, ByRef b)` work.

Read-only ByRef parameters are passed by value when the compiler can prove they are
never written, but calls into `<namespace>` modules and `AddressOf` delegates have
no visible callee signature, so a procedure entry always accepts both conventions
(see `unwrap`).

中文：VBS 的过程参数默认按引用传递，Python 只有按值传递，所以编译器把 ByRef 参数
boxes it in a Ref: it creates `_r = Ref(x)`, reads and writes `x.v` inside the procedure, then writes back, so
中文：包进 Ref 盒子：创建 `_r = Ref(x)`、过程内读写 `x.v`、调用后写回，这样
`Sub Swap(ByRef a, ByRef b)` works. A read-only ByRef parameter passes by value (the compiler skips the Ref
中文：`Sub Swap(ByRef a, ByRef b)` 才能生效。只读的 ByRef 参数会按值传（编译器跳过 Ref
and the write-back); a `<namespace>` module or an `AddressOf` delegate cannot see the callee's signature, so every procedure entry
中文：与回写）；但跨 `<namespace>` 模块与 `AddressOf` 委托看不到被调方签名，所以过程入口
Every procedure entry unpacks with `unwrap()` and accepts either convention.
中文：入口用 `unwrap()` 统一拆包，两种调用约定都能接受。
'''

from OpenVBSPP.ovpprt.empty import EMPTY


class Ref:
    """A writable box that models a ByRef argument.

    中文：可写盒子，用来模拟 ByRef 参数。
    """

    __slots__ = ("v",)

    def __init__(self, value=EMPTY):
        self.v = value

    def __repr__(self):
        return f"Ref({self.v!r})"


def unwrap(value):
    """Return the boxed value when the argument is a `Ref`, otherwise the value itself.

    中文：参数是 Ref 时取出里面的真实值，否则原样返回。
    """
    return value.v if type(value) is Ref else value
