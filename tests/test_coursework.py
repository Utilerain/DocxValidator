# -*- coding: utf-8 -*-
"""Эталонная курсовая: docs/test_coursework.docx проходит все правила без нарушений."""
import tempfile
import unittest
from pathlib import Path

from main import DocxValidator, default_rules

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / 'docs' / 'test_coursework.docx'


def build_and_validate():
    import sys
    sys.path.insert(0, str(ROOT / 'docs'))
    try:
        from make_coursework import build
    finally:
        sys.path.pop(0)
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / 'coursework.docx'
        build(str(path))
        return DocxValidator(default_rules()).validate(str(path))


class CourseworkTests(unittest.TestCase):
    def test_committed_sample_passes_all_rules(self):
        self.assertTrue(SAMPLE.exists(),
                        'docs/test_coursework.docx не найден (python docs/make_coursework.py)')
        report = DocxValidator(default_rules()).validate(str(SAMPLE))
        self.assertEqual([], report)

    def test_rebuilt_sample_passes_all_rules(self):
        self.assertEqual([], build_and_validate())


if __name__ == '__main__':
    unittest.main()
