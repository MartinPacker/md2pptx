"""Tests for processingOptions.ProcessingOptions.

Option names are stored lower-cased; the current value is a stack whose
bottom element (the default) can never be popped off.
"""

import pytest

from processingOptions import ProcessingOptions


@pytest.fixture
def options():
    return ProcessingOptions()


def test_set_option_values_sets_all_three_layers(options):
    options.setOptionValues("FontSize", 12)
    assert options.getDefaultOption("fontsize") == 12
    assert options.getPresentationOption("fontsize") == 12
    assert options.getCurrentOption("fontsize") == 12


def test_option_names_are_case_insensitive(options):
    options.setOptionValues("FontSize", 12)
    assert options.getCurrentOption("FONTSIZE") == 12
    assert options.getCurrentOption("fontsize") == 12


def test_missing_option_raises_key_error(options):
    with pytest.raises(KeyError):
        options.getDefaultOption("missing")
    with pytest.raises(KeyError):
        options.getCurrentOption("missing")


def test_current_option_stack_push_and_pop(options):
    options.setOptionValues("FontSize", 12)
    options.setCurrentOption("fontsize", 14)
    options.setCurrentOption("fontsize", 16)
    assert options.getCurrentOption("fontsize") == 16
    options.popCurrentOption("fontsize")
    assert options.getCurrentOption("fontsize") == 14
    options.popCurrentOption("fontsize")
    assert options.getCurrentOption("fontsize") == 12


def test_pop_never_removes_the_default(options):
    options.setOptionValues("FontSize", 12)
    options.popCurrentOption("fontsize")
    options.popCurrentOption("fontsize")
    assert options.getCurrentOption("fontsize") == 12


def test_set_option_values_array(options):
    options.setOptionValuesArray([("A", 1), ("B", 2)])
    assert options.getCurrentOption("a") == 1
    assert options.getCurrentOption("b") == 2


def test_set_option_values_does_not_overwrite_existing_default(options):
    options.setDefaultOption("A", 1)
    options.setOptionValues("A", 2)
    assert options.getDefaultOption("a") == 1
    assert options.getPresentationOption("a") == 2
    assert options.getCurrentOption("a") == 2


def test_dynamically_set_plain_and_typed_values(options):
    options.dynamicallySetOption("Name", "abc", "")
    assert options.getCurrentOption("name") == "abc"
    options.dynamicallySetOption("Width", "3.5", "float")
    assert options.getCurrentOption("width") == 3.5
    options.dynamicallySetOption("Count", "42", "int")
    assert options.getCurrentOption("count") == 42


def test_dynamically_set_default_and_pres_keywords(options):
    options.setOptionValues("A", 1)
    options.setCurrentOption("a", 9)
    options.dynamicallySetOption("A", "default", "")
    assert options.getCurrentOption("a") == 1
    options.setCurrentOption("a", 9)
    options.dynamicallySetOption("A", "pres", "")
    assert options.getCurrentOption("a") == 1


def test_dynamically_set_pop_keyword(options):
    options.setOptionValues("A", 1)
    options.setCurrentOption("a", 5)
    options.dynamicallySetOption("A", "pop", "")
    assert options.getCurrentOption("a") == 1
    options.setCurrentOption("a", 5)
    options.dynamicallySetOption("A", "prev", "")
    assert options.getCurrentOption("a") == 1


def test_dynamically_set_marks_option_as_changed(options):
    options.dynamicallySetOption("Width", "3.5", "float")
    assert options.dynamicallyChangedOptions["width"] is True
    assert options.hideMetadataStyle is False
