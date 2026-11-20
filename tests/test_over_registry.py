# -*- coding: utf-8 -*-
"""Tests for the over local registry (publish-less distribution).

中文：为 over 的本地 registry（publish-less 分发）补测试。"""
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from OpenVBSPP import over                       # noqa: E402


class RegistryCase(unittest.TestCase):
    """Build a packable package directory plus a standalone registry file.

    中文：准备一个可打包的包目录 + 一个独立的 registry 文件。"""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.registry = self.dir / "registry.json"
        self.pkg_dir = self.dir / "hello"
        self.pkg_dir.mkdir()
        (self.pkg_dir / "over.json").write_text(json.dumps({
            "name": "hello",
            "version": "1.2.3",
            "entry": "main.vbs",
            "description": "示例包",
        }), encoding="utf-8")
        (self.pkg_dir / "main.vbs").write_text('echo "hi"\n', encoding="utf-8")

    def tearDown(self):
        self._tmp.cleanup()

    def packed(self):
        return Path(over.pack(self.pkg_dir))

    def registry_data(self):
        return json.loads(self.registry.read_text(encoding="utf-8"))


class TestRegister(RegistryCase):
    def test_register_from_package_dir(self):
        info = over.register(self.pkg_dir, registry=self.registry)
        self.assertEqual(info["name"], "hello")
        self.assertEqual(info["version"], "1.2.3")
        self.assertEqual(len(info["hash"]), 64)
        self.assertGreater(info["size"], 0)

    def test_register_from_ovppkg(self):
        pkg = self.packed()
        info = over.register(pkg, registry=self.registry)
        self.assertEqual(info["version"], "1.2.3")

    def test_registry_structure(self):
        over.register(self.pkg_dir, registry=self.registry)
        data = self.registry_data()
        entry = data["packages"]["hello"]
        self.assertEqual(entry["name"], "hello")
        version = entry["versions"]["1.2.3"]
        self.assertEqual(version["entry"], "main.vbs")
        self.assertEqual(version["description"], "示例包")
        self.assertEqual(len(version["sha256"]), 64)
        self.assertIn("dependencies", version)

    def test_hash_matches_file(self):
        pkg = self.packed()
        info = over.register(pkg, registry=self.registry)
        digest, size = over._pkg_fingerprint(pkg)
        self.assertEqual(info["hash"], digest)
        self.assertEqual(info["size"], size)

    def test_same_version_is_overwritten(self):
        over.register(self.pkg_dir, registry=self.registry)
        over.register(self.pkg_dir, registry=self.registry)
        versions = self.registry_data()["packages"]["hello"]["versions"]
        self.assertEqual(list(versions), ["1.2.3"])

    def test_two_versions_coexist(self):
        over.register(self.pkg_dir, registry=self.registry)
        (self.pkg_dir / "over.json").write_text(json.dumps({
            "name": "hello", "version": "2.0.0", "entry": "main.vbs",
        }), encoding="utf-8")
        over.register(self.pkg_dir, registry=self.registry)
        versions = self.registry_data()["packages"]["hello"]["versions"]
        self.assertEqual(sorted(versions), ["1.2.3", "2.0.0"])

    def test_rejects_unknown_file(self):
        junk = self.dir / "nope.txt"
        junk.write_text("x", encoding="utf-8")
        with self.assertRaises(over.OverError):
            over.register(junk, registry=self.registry)


class TestIndexDirectory(RegistryCase):
    def test_indexes_all_packages(self):
        self.packed()
        other = self.dir / "world"
        other.mkdir()
        (other / "over.json").write_text(json.dumps({
            "name": "world", "version": "0.1.0", "entry": "main.vbs",
        }), encoding="utf-8")
        (other / "main.vbs").write_text('echo "w"\n', encoding="utf-8")
        over.pack(other)

        results = over.index_directory(self.dir, registry=self.registry)
        names = sorted(item["name"] for item in results)
        self.assertEqual(names, ["hello", "world"])
        self.assertEqual(sorted(self.registry_data()["packages"]), ["hello", "world"])

    def test_empty_directory(self):
        empty = self.dir / "empty"
        empty.mkdir()
        self.assertEqual(over.index_directory(empty, registry=self.registry), [])

    def test_not_a_directory(self):
        with self.assertRaises(over.OverError):
            over.index_directory(self.dir / "missing", registry=self.registry)


class TestRegistryVersions(RegistryCase):
    def test_sorted_versions(self):
        for version in ("1.0.0", "1.10.0", "1.2.0"):
            (self.pkg_dir / "over.json").write_text(json.dumps({
                "name": "hello", "version": version, "entry": "main.vbs",
            }), encoding="utf-8")
            over.register(self.pkg_dir, registry=self.registry)
        self.assertEqual(over.registry_versions("hello", registry=self.registry),
                         ["1.0.0", "1.2.0", "1.10.0"])

    def test_unknown_package(self):
        self.assertEqual(over.registry_versions("nope", registry=self.registry), [])


class TestSearchUsesRegistry(RegistryCase):
    def test_new_structure_is_searchable(self):
        over.register(self.pkg_dir, registry=self.registry)
        data = json.loads(self.registry.read_text(encoding="utf-8"))
        results = over.search("hello", registry=data["packages"])
        self.assertTrue(results)
        self.assertEqual(results[0]["name"], "hello")
        self.assertEqual(results[0]["version"], "1.2.3")


if __name__ == "__main__":
    unittest.main(verbosity=2)
