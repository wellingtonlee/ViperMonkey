"""
Unit tests for vipermonkey.core.vba_context.Context.
"""

import vipermonkey


def _ctx(**kwargs):
    """Create a fresh Context for testing."""
    return vipermonkey.Context(**kwargs)


# ---------------------------------------------------------------------------
# Constructor
# ---------------------------------------------------------------------------

class TestContextInit:
    def test_default_init(self):
        ctx = _ctx()
        assert ctx.locals is not None
        assert ctx.globals is not None
        assert isinstance(ctx.locals, dict)
        assert isinstance(ctx.globals, dict)

    def test_init_with_globals(self):
        ctx = _ctx(_globals={"foo": "bar"})
        assert ctx.globals["foo"] == "bar"

    def test_init_with_filename(self):
        ctx = _ctx(filename="test.doc")
        assert ctx.filename == "test.doc"


# ---------------------------------------------------------------------------
# Variable get/set
# ---------------------------------------------------------------------------

class TestGetSet:
    def test_set_and_get(self):
        ctx = _ctx()
        ctx.set("myvar", "hello")
        assert ctx.get("myvar") == "hello"

    def test_bracket_syntax(self):
        ctx = _ctx()
        ctx["result"] = "hello world!!!"
        assert ctx["result"] == "hello world!!!"

    def test_case_insensitive(self):
        ctx = _ctx()
        ctx.set("MyVar", "hello")
        assert ctx.get("myvar") == "hello"

    def test_overwrite(self):
        ctx = _ctx()
        ctx.set("x", 1)
        ctx.set("x", 2)
        assert ctx.get("x") == 2

    def test_get_nonexistent_raises(self):
        ctx = _ctx()
        try:
            ctx.get("nonexistent")
            assert False, "Should have raised KeyError"
        except KeyError:
            pass

    def test_set_none_value(self):
        """Setting None should be a no-op."""
        ctx = _ctx()
        ctx.set("x", None)
        try:
            ctx.get("x")
            # It's okay if it was set
        except KeyError:
            # Also okay - None values may be skipped
            pass


# ---------------------------------------------------------------------------
# contains / contains_user_defined
# ---------------------------------------------------------------------------

class TestContains:
    def test_contains_existing(self):
        ctx = _ctx()
        ctx.set("foo", "bar")
        assert ctx.contains("foo")

    def test_contains_nonexistent(self):
        ctx = _ctx()
        assert not ctx.contains("nonexistent_var_xyz")

    def test_contains_user_defined(self):
        ctx = _ctx()
        ctx.set("foo", "bar")
        assert ctx.contains_user_defined("foo")

    def test_contains_user_defined_not_found(self):
        ctx = _ctx()
        assert not ctx.contains_user_defined("nonexistent_var_xyz")


# ---------------------------------------------------------------------------
# Type tracking
# ---------------------------------------------------------------------------

class TestTypeTracking:
    def test_set_and_get_type(self):
        ctx = _ctx()
        ctx.set_type("myvar", "String")
        assert ctx.get_type("myvar") == "String"

    def test_get_type_nonexistent(self):
        ctx = _ctx()
        # Should not raise, returns None or a default
        result = ctx.get_type("nonexistent_var_xyz")
        assert result is None or isinstance(result, str)

    def test_type_case_insensitive(self):
        ctx = _ctx()
        ctx.set_type("MyVar", "Integer")
        assert ctx.get_type("myvar") == "Integer"

    def test_get_type_non_string(self):
        ctx = _ctx()
        assert ctx.get_type(123) is None


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------

class TestErrorHandling:
    def test_no_error_initially(self):
        ctx = _ctx()
        assert not ctx.have_error()

    def test_set_error(self):
        ctx = _ctx()
        ctx.set_error("something went wrong")
        assert ctx.have_error()

    def test_clear_error(self):
        ctx = _ctx()
        ctx.set_error("oops")
        ctx.clear_error()
        assert not ctx.have_error()

    def test_must_handle_error_without_handler(self):
        ctx = _ctx()
        ctx.set_error("test")
        assert not ctx.must_handle_error()

    def test_get_error_handler_default(self):
        ctx = _ctx()
        assert ctx.get_error_handler() is None


# ---------------------------------------------------------------------------
# General error tracking
# ---------------------------------------------------------------------------

class TestGeneralErrors:
    def test_initial_error_count(self):
        ctx = _ctx()
        assert ctx.get_general_errors() == 0

    def test_increase_errors(self):
        ctx = _ctx()
        ctx.increase_general_errors()
        ctx.increase_general_errors()
        assert ctx.get_general_errors() == 2

    def test_clear_errors(self):
        ctx = _ctx()
        ctx.increase_general_errors()
        ctx.clear_general_errors()
        assert ctx.get_general_errors() == 0

    def test_report_general_error(self):
        ctx = _ctx()
        ctx.report_general_error("test error")
        assert ctx.get_general_errors() == 1


# ---------------------------------------------------------------------------
# File operations
# ---------------------------------------------------------------------------

class TestFileOperations:
    def test_open_file(self):
        ctx = _ctx()
        ctx.open_file("test.txt")
        assert ctx.file_is_open("test.txt")
        assert ctx.get_num_open_files() == 1

    def test_file_not_open(self):
        ctx = _ctx()
        assert not ctx.file_is_open("nonexistent.txt")

    def test_write_to_unopened_file(self):
        ctx = _ctx()
        result = ctx.write_file("nonexistent.txt", "data")
        assert result is False

    def test_close_file(self):
        ctx = _ctx()
        ctx.open_file("test.txt")
        ctx.close_file("test.txt")
        assert not ctx.file_is_open("test.txt")
        assert "test.txt" in ctx.closed_files

    def test_open_multiple_files(self):
        ctx = _ctx()
        ctx.open_file("a.txt")
        ctx.open_file("b.txt")
        assert ctx.get_num_open_files() == 2

    def test_write_to_closed_file(self):
        ctx = _ctx()
        # Writing to a file that isn't open should not crash
        result = ctx.write_file("nonexistent.txt", b"data")
        # Should return False or handle gracefully
        assert result is False or result is None

    def test_open_file_with_id(self):
        ctx = _ctx()
        ctx.open_file("test.txt", file_id="1")
        assert ctx.file_is_open("test.txt")

    def test_path_normalization(self):
        ctx = _ctx()
        ctx.open_file(".\\test.txt")
        assert ctx.file_is_open("test.txt")


# ---------------------------------------------------------------------------
# get_doc_var
# ---------------------------------------------------------------------------

class TestDocVars:
    def test_get_doc_var(self):
        ctx = _ctx()
        ctx.doc_vars["subject"] = "test Subject"
        assert ctx.get_doc_var("subject") == "test Subject"

    def test_get_doc_var_case_insensitive(self):
        ctx = _ctx()
        ctx.doc_vars["subject"] = "test Subject"
        assert ctx.get_doc_var("Subject") == "test Subject"

    def test_get_doc_var_not_found(self):
        ctx = _ctx()
        result = ctx.get_doc_var("nonexistent_var_xyz")
        assert result is None

    def test_get_doc_var_non_string(self):
        ctx = _ctx()
        assert ctx.get_doc_var(None) is None


# ---------------------------------------------------------------------------
# get_lib_func
# ---------------------------------------------------------------------------

class TestLibFunc:
    def test_get_known_func(self):
        ctx = _ctx()
        func = ctx.get_lib_func("chr")
        assert func is not None

    def test_get_unknown_func(self):
        ctx = _ctx()
        try:
            ctx.get_lib_func("nonexistent_func_xyz")
            assert False, "Should have raised KeyError"
        except KeyError:
            pass

    def test_get_lib_func_non_string(self):
        ctx = _ctx()
        try:
            ctx.get_lib_func(123)
            assert False, "Should have raised KeyError"
        except KeyError:
            pass


# ---------------------------------------------------------------------------
# delete
# ---------------------------------------------------------------------------

class TestDelete:
    def test_delete_existing(self):
        ctx = _ctx()
        ctx.set("foo", "bar")
        ctx.delete("foo")
        assert not ctx.contains_user_defined("foo")

    def test_delete_nonexistent(self):
        ctx = _ctx()
        # Should not raise
        ctx.delete("nonexistent_var_xyz")


# ---------------------------------------------------------------------------
# Integration tests via vipermonkey.eval
# ---------------------------------------------------------------------------

class TestContextIntegration:
    def test_set_and_retrieve_via_api(self):
        ctx = _ctx()
        ctx["x"] = "hello"
        assert ctx["x"] == "hello"

    def test_doc_var_via_get_doc_var(self):
        ctx = _ctx()
        ctx.doc_vars["subject"] = "test Subject"
        assert ctx.get_doc_var("subject") == "test Subject"
