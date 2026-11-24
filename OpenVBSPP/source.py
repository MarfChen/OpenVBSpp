"""Source file loading with BOM detection.

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
"""

import codecs
from pathlib import Path


def read_source(path):
    """Read UTF-8 or BOM-marked UTF-16 LE source files."""
    data = Path(path).read_bytes()

    if (
        data.startswith(codecs.BOM_UTF16_LE)
        and not data.startswith(codecs.BOM_UTF32_LE)
    ):
        encoding = "utf-16"
    else:
        encoding = "utf-8-sig"

    return data.decode(encoding).replace("\r\n", "\n").replace("\r", "\n")
