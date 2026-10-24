"""Tests for rectangle.Rectangle geometry."""

from rectangle import Rectangle


def test_right_and_bottom_derived_from_left_top_and_size():
    r = Rectangle(10, 20, 30, 40)
    assert (r.top, r.left, r.height, r.width) == (10, 20, 30, 40)
    assert r.right == 60
    assert r.bottom == 40


def test_zero_rectangle():
    r = Rectangle(0, 0, 0, 0)
    assert r.right == 0
    assert r.bottom == 0


def test_float_dimensions():
    r = Rectangle(1.5, 2.5, 0.5, 0.25)
    assert r.right == 2.75
    assert r.bottom == 2.0


def test_negative_offsets():
    r = Rectangle(-5, -10, 20, 30)
    assert r.right == 20
    assert r.bottom == 15
