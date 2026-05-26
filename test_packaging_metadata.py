"""Package metadata smoke tests."""
import importlib
import os
import sys
import unittest

try:
    import tomllib
except ImportError:  # pragma: no cover - Python < 3.11
    tomllib = None


sys.path.insert(0, os.path.dirname(__file__))


RUNTIME_MODULES = [
    "Spreadsheet_LLM_Encoder",
    "paper_serializers",
    "chain_of_spreadsheet",
    "llm_backend",
    "tokenizer",
    "evaluation",
    "evaluation_metadata",
    "baselines",
]


class TestPackagingMetadata(unittest.TestCase):

    def test_runtime_modules_import(self):
        for module_name in RUNTIME_MODULES:
            with self.subTest(module=module_name):
                self.assertIsNotNone(importlib.import_module(module_name))

    def test_runtime_modules_declared_in_pyproject(self):
        if tomllib is None:
            self.skipTest("tomllib unavailable")
        pyproject_path = os.path.join(os.path.dirname(__file__), "pyproject.toml")
        with open(pyproject_path, "rb") as fh:
            pyproject = tomllib.load(fh)

        declared = set(pyproject["tool"]["setuptools"]["py-modules"])
        missing = sorted(set(RUNTIME_MODULES) - declared)
        self.assertEqual(missing, [])

    def test_optional_dependency_groups_declared(self):
        if tomllib is None:
            self.skipTest("tomllib unavailable")
        pyproject_path = os.path.join(os.path.dirname(__file__), "pyproject.toml")
        with open(pyproject_path, "rb") as fh:
            pyproject = tomllib.load(fh)

        optional = pyproject["project"]["optional-dependencies"]
        for group in ("tokenizer", "openai", "xlsb", "finetune", "qlora", "baselines", "all"):
            with self.subTest(group=group):
                self.assertIn(group, optional)
                self.assertGreater(len(optional[group]), 0)


if __name__ == "__main__":
    unittest.main()
