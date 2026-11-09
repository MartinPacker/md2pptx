"""End-to-end tests for media lines that start with white space (issue #245).

A graphic such as ``![](pic.png)`` must be placed on the slide even when the
line starts with a few spaces. Four or more spaces still make a code block.
"""

import os
import shutil
import subprocess
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

REPO = os.path.join(os.path.dirname(__file__), "..")
SCRIPT = os.path.join(REPO, "md2pptx")
PICTURE = os.path.join(REPO, "test", "Battery W2M.png")


def build(tmp_path, markdown):
    shutil.copy(PICTURE, tmp_path / "pic.png")
    (tmp_path / "in.md").write_text(markdown)
    subprocess.run(
        [sys.executable, SCRIPT, "in.md", "out.pptx"],
        cwd=tmp_path,
        check=True,
        capture_output=True,
    )
    return Presentation(str(tmp_path / "out.pptx"))


def picture_count(slide):
    return sum(1 for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE)


def test_unindented_graphic_is_placed(tmp_path):
    prs = build(tmp_path, "# Slide\n\n![](pic.png)\n")
    assert picture_count(prs.slides[0]) == 1


def test_graphic_after_leading_spaces_is_placed(tmp_path):
    prs = build(tmp_path, "# Slide\n\n  ![](pic.png)\n")
    slide = prs.slides[0]
    assert picture_count(slide) == 1
    assert not slide.has_notes_slide or "pic.png" not in slide.notes_slide.notes_text_frame.text


def test_clickable_graphic_after_leading_tab_is_placed(tmp_path):
    prs = build(tmp_path, "# Slide\n\n\t[![](pic.png)](https://example.com)\n")
    assert picture_count(prs.slides[0]) == 1


def test_graphic_after_leading_spaces_in_side_by_side_row_is_placed(tmp_path):
    prs = build(tmp_path, "# Slide\n\n ![](pic.png) ![](pic.png)\n")
    assert picture_count(prs.slides[0]) == 2


def test_graphic_indented_as_code_stays_code(tmp_path):
    prs = build(tmp_path, "# Slide\n\n    ![](pic.png)\n")
    assert picture_count(prs.slides[0]) == 0
