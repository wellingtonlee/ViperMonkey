# Changelog

## Unreleased

### Modernization & Python 3 Compatibility
- Remove all Python 2 compatibility shims (`from __future__`, `unicode`,
  `basestring`, `.iteritems()`, `raw_input`, `optparse`)
- Replace 91 bare `except:` clauses with specific exception types across 14
  files
- Convert all implicit relative imports to explicit relative imports (30+
  files)
- Fix Python 2-only syntax (`print` statements, `exec`, `collections.Iterable`,
  `str.decode()`)
- Add `python_requires>=3.6` to `setup.py`

### Dependency Cleanup
- Unpin `pyparsing` (now `>=2.2.0,<3`) and `unidecode`
- Sync `requirements.txt` with `setup.py`; remove `xlrd2` duplicate
- Add `long_description_content_type` to `setup.py`; remove deprecated
  `setup_requires`

### Logging
- Convert 269 debug `print()` calls to `log.debug()`/`log.error()` in
  `strip_lines.py`, `read_ole_fields.py`, `__init__.py`, and `statements.py`
- Add debug logging to silent `except-pass` blocks in `utils.py`

### Code Quality
- Add `pyproject.toml` with ruff linter config (target Python 3.6+)
- Fix non-Pythonic comparisons: `== None` -> `is None`, `== True` ->
  `is True`, `type(x) == type(y)` -> `isinstance()`
- Clean up stale and duplicate TODO comments across 5 files

### Testing
- Add 100+ unit tests for `vba_library.py` (Chr, Len, Mid, Left, Right, Trim,
  UCase, LCase, AscW, Replace, InStr, Split, Int, math functions, etc.)
- Add 48 unit tests for `vba_context.py` (get/set, contains, type tracking,
  error handling, file operations, doc vars, lib funcs, delete)

### Developer Tooling
- Add `Makefile` with `install-dev`, `test`, `lint`, `lint-fix`, `clean`
  targets
- Add `.pre-commit-config.yaml` with ruff and pre-commit-hooks
- Add `extras_require` dev dependencies (`pytest`, `ruff`, `pre-commit`) to
  `setup.py`

### Refactoring
- Decompose monolithic `vba_library.py` (5609 lines, 170+ classes) into a
  Python package with 10 categorized submodules:
  - `_string` — string manipulation (Mid, Left, Right, Replace, InStr, etc.)
  - `_math` — numeric/math functions (Int, Fix, Abs, Rnd, Log, etc.)
  - `_typecheck` — type inspection (IsNumeric, IsNull, IsArray, VarType, etc.)
  - `_datetime` — date/time functions (Date, DateAdd, DateDiff, Timer, etc.)
  - `_fileio` — file I/O operations (Open, Close, FreeFile, Dir, etc.)
  - `_system` — system/environment (Shell, Environ, CreateObject, etc.)
  - `_excel` — Excel-specific functions (Cells, Range, Sheets, etc.)
  - `_array` — array operations (Array, UBound, LBound, Items, etc.)
  - `_network` — network functions (URLDownloadToFile, Navigate, etc.)
  - `_misc` — miscellaneous (MsgBox, IIf, Switch, Debug, etc.)
  - `_common` — shared imports and helper utilities
- All existing import patterns preserved via `__init__.py` re-exports
- Zero test regressions: 149 pass, 22 pre-existing failures unchanged
