'''
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

VBScript `Nothing`: the empty object reference, tested with `x Is Nothing`.

中文：VBS 的 Nothing —— 空对象引用，用 `x Is Nothing` 判断。
'''


class NothingType:
    """The `Nothing` sentinel (an empty object reference).

    中文：Nothing 哨兵（空对象引用）。
    """

    __slots__ = ()

    def __str__(self):
        return "Nothing"

    def __repr__(self):
        return "Nothing"

    def __bool__(self):
        return False

    def __hash__(self):
        return hash("Nothing")

    def __eq__(self, other):
        # VBScripts compares objects with Is; this only has to make `Nothing Is Nothing` true.
        # 中文：VBS 里对象比较要用 Is，这里只保证 Nothing Is Nothing 成立。
        return other is self

    def __ne__(self, other):
        return other is not self


NOTHING = NothingType()
