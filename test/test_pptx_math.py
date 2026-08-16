import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pptx_math
from pptx_math import _resolve_mathxsl_path


class ResolveMathXslPathTests(unittest.TestCase):
    def test_prefers_configured_stylesheet(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            configured = root / "custom.xsl"
            fallback = root / "mml2omml.xsl"
            configured.touch()
            fallback.touch()

            self.assertEqual(
                _resolve_mathxsl_path(configured, fallback),
                configured,
            )

    def test_uses_install_directory_fallback_when_configured_path_is_missing(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            configured = root / "missing.xsl"
            fallback = root / "mml2omml.xsl"
            fallback.touch()

            self.assertEqual(
                _resolve_mathxsl_path(configured, fallback),
                fallback,
            )

    def test_uses_install_directory_fallback_when_option_is_not_set(self):
        with tempfile.TemporaryDirectory() as directory:
            fallback = Path(directory) / "mml2omml.xsl"
            fallback.touch()

            self.assertEqual(
                _resolve_mathxsl_path(None, fallback),
                fallback,
            )

    def test_default_fallback_is_beside_pptx_math_module(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            fallback = root / "mml2omml.xsl"
            fallback.touch()

            with mock.patch.object(pptx_math, "__file__", str(root / "pptx_math.py")):
                self.assertEqual(_resolve_mathxsl_path(None), fallback)

    def test_reports_configured_and_fallback_paths_when_neither_exists(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            configured = root / "missing.xsl"
            fallback = root / "mml2omml.xsl"

            with self.assertRaises(FileNotFoundError) as caught:
                _resolve_mathxsl_path(configured, fallback)

            message = str(caught.exception)
            self.assertIn(str(configured), message)
            self.assertIn(str(fallback), message)


if __name__ == "__main__":
    unittest.main()
