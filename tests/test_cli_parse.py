"""Hand-rolled argv parser for md.py."""
import pytest

from md import Args, parse_args


@pytest.mark.parametrize("argv, want", [
    (["showcase"],                        Args("showcase", "live")),
    (["showcase", "-build"],              Args("showcase", "build")),
    (["showcase", "-b"],                  Args("showcase", "build")),
    (["showcase", "-diff"],               Args("showcase", "diff")),
    (["showcase", "-d"],                  Args("showcase", "diff")),
    (["showcase", "-open"],               Args("showcase", "open")),
    (["showcase", "-o"],                  Args("showcase", "open")),
    (["showcase", "-import", "a.docx"],   Args("showcase", "import", "a.docx")),
    (["showcase", "-i", "a.docx"],        Args("showcase", "import", "a.docx")),
    (["-new", "foo"],                     Args(None, "new", "foo")),
])
def test_happy_paths(argv, want):
    assert parse_args(argv) == want


@pytest.mark.parametrize("argv", [
    [],                             # no args
    ["-new"],                       # -new with no name
    ["-new", "a", "b"],             # -new with extra
    ["proj", "-bogus"],             # unknown flag
    ["proj", "-i"],                 # -import missing file
    ["proj", "-build", "extra"],    # -build takes no extra
])
def test_invalid(argv):
    with pytest.raises(ValueError):
        parse_args(argv)
