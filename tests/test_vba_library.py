"""
Unit tests for VBA library function emulators in vipermonkey.core.vba_library.
"""

import vipermonkey
from vipermonkey.core import vba_library


def _ctx():
    """Create a minimal context for testing."""
    return vipermonkey.Context()


def _eval(func_cls, params):
    """Helper: instantiate a VbaLibraryFunc class and eval it with a fresh context."""
    return func_cls().eval(_ctx(), params)


# ---------------------------------------------------------------------------
# Chr / ChrW / ChrB
# ---------------------------------------------------------------------------

class TestChr:
    def test_basic_ascii(self):
        assert _eval(vba_library.Chr, [65]) == "A"

    def test_lowercase(self):
        assert _eval(vba_library.Chr, [97]) == "a"

    def test_zero(self):
        assert _eval(vba_library.Chr, [0]) == "\x00"

    def test_float_param(self):
        # Float should be rounded to nearest int
        assert _eval(vba_library.Chr, [65.4]) == "A"
        assert _eval(vba_library.Chr, [65.6]) == "B"

    def test_no_params(self):
        assert _eval(vba_library.Chr, None) == "NULL"
        assert _eval(vba_library.Chr, []) == "NULL"

    def test_chrw(self):
        assert _eval(vba_library.ChrW, [65]) == "A"

    def test_chrb(self):
        assert _eval(vba_library.ChrB, [65]) == "A"

    def test_negative_param(self):
        # Negative values are made positive
        assert _eval(vba_library.Chr, [-65]) == "A"


# ---------------------------------------------------------------------------
# Len / LenB
# ---------------------------------------------------------------------------

class TestLen:
    def test_string(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Len().eval(ctx, ["hello"]) == 5

    def test_empty_string(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Len().eval(ctx, [""]) == 0

    def test_integer_param(self):
        assert _eval(vba_library.Len, [123]) == 3

    def test_lenb_basic(self):
        assert _eval(vba_library.LenB, ["hello"]) == 5

    def test_lenb_no_params(self):
        assert _eval(vba_library.LenB, None) == 0
        assert _eval(vba_library.LenB, []) == 0


# ---------------------------------------------------------------------------
# Mid
# ---------------------------------------------------------------------------

class TestMid:
    def test_basic(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Mid().eval(ctx, ["hello world", 7]) == "world"

    def test_with_length(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Mid().eval(ctx, ["hello world", 1, 5]) == "hello"

    def test_start_beyond_length(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Mid().eval(ctx, ["hello", 10]) == ""

    def test_start_zero(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Mid().eval(ctx, ["hello", 0]) == "NULL"

    def test_no_params(self):
        assert _eval(vba_library.Mid, None) == ""

    def test_null_string(self):
        assert _eval(vba_library.Mid, ["NULL", 1]) == "\x00"


# ---------------------------------------------------------------------------
# Left / Right
# ---------------------------------------------------------------------------

class TestLeftRight:
    def test_left_basic(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Left().eval(ctx, ["hello world", 5]) == "hello"

    def test_left_exceeds_length(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Left().eval(ctx, ["hi", 10]) == "hi"

    def test_left_zero(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Left().eval(ctx, ["hello", 0]) == ""

    def test_left_no_params(self):
        assert _eval(vba_library.Left, None) == "NULL"
        assert _eval(vba_library.Left, ["hello"]) == "NULL"

    def test_right_basic(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Right().eval(ctx, ["hello world", 5]) == "world"

    def test_right_exceeds_length(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Right().eval(ctx, ["hi", 10]) == "hi"

    def test_right_zero(self):
        ctx = _ctx()
        ctx.is_vbscript = True
        assert vba_library.Right().eval(ctx, ["hello", 0]) == ""

    def test_right_no_params(self):
        assert _eval(vba_library.Right, None) == "NULL"


# ---------------------------------------------------------------------------
# Trim / RTrim / LTrim
# ---------------------------------------------------------------------------

class TestTrim:
    def test_trim(self):
        assert _eval(vba_library.Trim, ["  hello  "]) == "hello"

    def test_trim_no_whitespace(self):
        assert _eval(vba_library.Trim, ["hello"]) == "hello"

    def test_trim_empty(self):
        assert _eval(vba_library.Trim, [""]) == ""

    def test_trim_no_params(self):
        assert _eval(vba_library.Trim, None) == ""
        assert _eval(vba_library.Trim, []) == ""

    def test_rtrim(self):
        assert _eval(vba_library.RTrim, ["hello   "]) == "hello"

    def test_rtrim_no_params(self):
        assert _eval(vba_library.RTrim, None) == "NULL"

    def test_rtrim_integer(self):
        assert _eval(vba_library.RTrim, [42]) == "42"

    def test_ltrim(self):
        assert _eval(vba_library.LTrim, ["   hello"]) == "hello"

    def test_ltrim_no_params(self):
        assert _eval(vba_library.LTrim, None) == ""

    def test_ltrim_integer(self):
        assert _eval(vba_library.LTrim, [42]) == "42"


# ---------------------------------------------------------------------------
# UCase / LCase
# ---------------------------------------------------------------------------

class TestCase:
    def test_ucase(self):
        assert _eval(vba_library.UCase, ["hello"]) == "HELLO"

    def test_lcase(self):
        assert _eval(vba_library.LCase, ["HELLO"]) == "hello"

    def test_ucase_mixed(self):
        assert _eval(vba_library.UCase, ["Hello World"]) == "HELLO WORLD"

    def test_lcase_mixed(self):
        assert _eval(vba_library.LCase, ["Hello World"]) == "hello world"


# ---------------------------------------------------------------------------
# AscW
# ---------------------------------------------------------------------------

class TestAscW:
    def test_basic(self):
        assert _eval(vba_library.AscW, ["A"]) == 65

    def test_string_takes_first_char(self):
        assert _eval(vba_library.AscW, ["ABC"]) == 65

    def test_null_string(self):
        assert _eval(vba_library.AscW, ["NULL"]) == 0

    def test_integer_param(self):
        assert _eval(vba_library.AscW, [65]) == 65

    def test_no_params(self):
        assert _eval(vba_library.AscW, None) == "NULL"

    def test_empty_string(self):
        assert _eval(vba_library.AscW, [""]) == 0


# ---------------------------------------------------------------------------
# Replace
# ---------------------------------------------------------------------------

class TestReplace:
    def test_basic(self):
        assert _eval(vba_library.Replace, ["hello world", "world", "mars"]) == "hello mars"

    def test_multiple_occurrences(self):
        assert _eval(vba_library.Replace, ["aaa", "a", "b"]) == "bbb"

    def test_no_match(self):
        assert _eval(vba_library.Replace, ["hello", "xyz", "abc"]) == "hello"

    def test_empty_find(self):
        assert _eval(vba_library.Replace, ["hello", "", "x"]) == "hello"

    def test_no_params(self):
        assert _eval(vba_library.Replace, None) == ""

    def test_insufficient_params(self):
        assert _eval(vba_library.Replace, ["hello"]) == "hello"

    def test_two_params(self):
        assert _eval(vba_library.Replace, ["hello", "l"]) == "hello"


# ---------------------------------------------------------------------------
# InStr
# ---------------------------------------------------------------------------

class TestInStr:
    def test_basic_found(self):
        assert _eval(vba_library.InStr, ["hello world", "world"]) == 7

    def test_not_found(self):
        assert _eval(vba_library.InStr, ["hello", "xyz"]) == 0

    def test_empty_haystack(self):
        assert _eval(vba_library.InStr, ["", "a"]) == 0

    def test_empty_needle(self):
        assert _eval(vba_library.InStr, ["hello", ""]) == 0

    def test_with_start_position(self):
        # When first param is int, it's the start position (1-based)
        assert _eval(vba_library.InStr, [7, "hello world", "world"]) == 7

    def test_no_params(self):
        assert _eval(vba_library.InStr, None) == "NULL"
        assert _eval(vba_library.InStr, ["hello"]) == "NULL"


# ---------------------------------------------------------------------------
# Split
# ---------------------------------------------------------------------------

class TestSplit:
    def test_default_separator(self):
        assert _eval(vba_library.Split, ["hello world"]) == ["hello", "world"]

    def test_custom_separator(self):
        assert _eval(vba_library.Split, ["a,b,c", ","]) == ["a", "b", "c"]

    def test_no_params(self):
        assert _eval(vba_library.Split, None) == "NULL"
        assert _eval(vba_library.Split, []) == "NULL"


# ---------------------------------------------------------------------------
# Join
# ---------------------------------------------------------------------------

class TestJoin:
    def test_no_params(self):
        assert _eval(vba_library.Join, None) == "NULL"


# ---------------------------------------------------------------------------
# Int / CInt
# ---------------------------------------------------------------------------

class TestInt:
    def test_basic(self):
        assert _eval(vba_library.Int, [42]) == 42

    def test_string_int(self):
        assert _eval(vba_library.Int, ["42"]) == 42

    def test_hex_string(self):
        assert _eval(vba_library.Int, ["&HFF"]) == 255

    def test_no_params(self):
        assert _eval(vba_library.Int, None) == "NULL"
        assert _eval(vba_library.Int, []) == "NULL"

    def test_cint_alias(self):
        assert _eval(vba_library.CInt, [42]) == 42


# ---------------------------------------------------------------------------
# Math functions: Sgn, Sqr, Abs, Fix
# ---------------------------------------------------------------------------

class TestMathFunctions:
    def test_sgn_positive(self):
        assert _eval(vba_library.Sgn, [5]) == 1

    def test_sgn_negative(self):
        assert _eval(vba_library.Sgn, [-5]) == -1

    def test_sgn_zero(self):
        assert _eval(vba_library.Sgn, [0]) == 0

    def test_sgn_no_params(self):
        assert _eval(vba_library.Sgn, None) == "NULL"

    def test_sqr(self):
        assert _eval(vba_library.Sqr, [9]) == 3.0

    def test_sqr_no_params(self):
        assert _eval(vba_library.Sqr, None) == "NULL"

    def test_abs_positive(self):
        assert _eval(vba_library.Abs, [5]) == 5

    def test_abs_negative(self):
        assert _eval(vba_library.Abs, [-5]) == 5

    def test_abs_zero(self):
        assert _eval(vba_library.Abs, [0]) == 0

    def test_abs_no_params(self):
        assert _eval(vba_library.Abs, None) == "NULL"

    def test_fix(self):
        assert _eval(vba_library.Fix, [3.7]) == 3

    def test_fix_negative(self):
        assert _eval(vba_library.Fix, [-3.7]) == -4

    def test_fix_no_params(self):
        assert _eval(vba_library.Fix, None) == "NULL"


# ---------------------------------------------------------------------------
# Hex
# ---------------------------------------------------------------------------

class TestHex:
    def test_no_params(self):
        assert _eval(vba_library.Hex, None) == "NULL"
        assert _eval(vba_library.Hex, []) == "NULL"


# ---------------------------------------------------------------------------
# CBool
# ---------------------------------------------------------------------------

class TestCBool:
    def test_true(self):
        assert _eval(vba_library.CBool, [True]) == 1

    def test_one(self):
        assert _eval(vba_library.CBool, [1]) == 1

    def test_false(self):
        assert _eval(vba_library.CBool, [False]) == 0

    def test_zero(self):
        assert _eval(vba_library.CBool, [0]) == 0

    def test_no_params(self):
        assert _eval(vba_library.CBool, None) == "NULL"


# ---------------------------------------------------------------------------
# Space
# ---------------------------------------------------------------------------

class TestSpace:
    def test_basic(self):
        assert _eval(vba_library.Space, [5]) == "     "

    def test_zero(self):
        assert _eval(vba_library.Space, [0]) == ""


# ---------------------------------------------------------------------------
# StrReverse
# ---------------------------------------------------------------------------

class TestStrReverse:
    def test_basic(self):
        assert _eval(vba_library.StrReverse, ["hello"]) == "olleh"

    def test_empty(self):
        assert _eval(vba_library.StrReverse, [""]) == ""

    def test_no_params(self):
        assert _eval(vba_library.StrReverse, None) == "NULL"


# ---------------------------------------------------------------------------
# IsEmpty / IsNull
# ---------------------------------------------------------------------------

class TestIsEmpty:
    def test_empty_string(self):
        assert _eval(vba_library.IsEmpty, [""]) is True

    def test_nonempty_string(self):
        assert _eval(vba_library.IsEmpty, ["hello"]) is False

    def test_no_params(self):
        assert _eval(vba_library.IsEmpty, None) is True


class TestIsNull:
    def test_none(self):
        assert _eval(vba_library.IsNull, [None]) is True

    def test_non_null(self):
        assert _eval(vba_library.IsNull, ["hello"]) is False


# ---------------------------------------------------------------------------
# Switch
# ---------------------------------------------------------------------------

class TestSwitch:
    def test_first_true(self):
        assert _eval(vba_library.Switch, [True, "yes", False, "no"]) == "yes"

    def test_second_true(self):
        assert _eval(vba_library.Switch, [False, "no", True, "yes"]) == "yes"

    def test_no_true(self):
        result = _eval(vba_library.Switch, [False, "no", False, "nope"])
        assert result == "NULL"


# ---------------------------------------------------------------------------
# Hour (stub)
# ---------------------------------------------------------------------------

class TestHour:
    def test_returns_13(self):
        assert _eval(vba_library.Hour, [None]) == 13


# ---------------------------------------------------------------------------
# Integration: test via vipermonkey.eval
# ---------------------------------------------------------------------------

class TestViaEval:
    def test_chr(self):
        assert vipermonkey.eval('Chr(65)') == "A"
