"""
ViperMonkey: VBA Library (package)

VBA library function emulators, organized into submodules by category.
All classes and helpers are re-exported here for backward compatibility.
"""

__version__ = '0.02'

import logging

# Re-export module-level state and helpers so that existing code like
# vba_library.run_function() and vba_library.member_access() still works.
from ._common import (
    member_access,
    get_raw_shellcode_data,
    run_external_function,
    run_function,
    _read_cell,
    VBA_LIBRARY,
    log,
)

# Module-level mutable state — external code does e.g.
# vba_library.var_names = self.params, which sets the package attribute.
from ._common import var_names, parse_cache, ticks  # noqa: F401

# Import all classes from submodules
from ._string import *  # noqa: F401,F403
from ._string import _Chr  # underscore-prefixed, not covered by wildcard
from ._math import *  # noqa: F401,F403
from ._typecheck import *  # noqa: F401,F403
from ._datetime import *  # noqa: F401,F403
from ._fileio import *  # noqa: F401,F403
from ._system import *  # noqa: F401,F403
from ._excel import *  # noqa: F401,F403
from ._array import *  # noqa: F401,F403
from ._network import *  # noqa: F401,F403
from ._misc import *  # noqa: F401,F403

# Also re-export everything from _common that submodules might have pulled in
from ._common import *  # noqa: F401,F403

# --- REGISTRATION -----------------------------------------------------------
# Register all VBA library function classes into VBA_LIBRARY dict.

for _class in (MsgBox, Shell, Len, Mid, MidB, Left, Right,
               BuiltInDocumentProperties, Array, UBound, LBound, Trim,
               StrConv, Split, Int, Item, StrReverse, InStr, Replace,
               Sgn, Sqr, Base64Decode, Abs, Fix, Hex, String, CByte, Atn,
               Dir, RGB, Log, Cos, Exp, Sin, Str, Val, CInt, Pmt, Day, Round,
               UCase, Randomize, CBool, CDate, CStr, CSng, Tan, Rnd, Oct,
               Environ, IIf, CleanString, Base64DecodeString, CLng, Close, Put, Run, InStrRev,
               LCase, RTrim, LTrim, AscW, AscB, CurDir, LenB, CreateObject,
               CheckSpelling, Specialfolders, StrComp, Space, Year, Variable,
               Exec, CDbl, Print, OpenTextFile, CreateTextFile, Write, Minute, Second, WinExec,
               CallByName, ReadText, Variables, Timer, Open, CVErr, WriteLine,
               URLDownloadToFile, FollowHyperlink, Join, VarType, DriveExists, Navigate,
               KeyString, CVar, IsNumeric, Assert, Sleep, Cells, Shapes,
               Format, Range, Switch, WeekDay, ShellExecute, OpenTextFile, GetTickCount,
               Month, ExecQuery, ExpandEnvironmentStrings, Execute, Eval, ExecuteGlobal,
               Unescape, FolderExists, IsArray, FileExists, Debug, GetExtensionName,
               AddCode, StrPtr, International, ExecuteStatement, InlineShapes,
               RegWrite, QBColor, LoadXML, SaveToFile, InternetGetConnectedState, InternetOpenA,
               FreeFile, GetByteCount_2, GetBytes_4, TransformFinalBlock, Add, Raise, Echo,
               AddFromString, Not, PrivateProfileString, GetCursorPos, CreateElement,
               IsObject, NumPut, GetLocale, URLDownloadToFile, URLDownloadToFileA,
               URLDownloadToFileW, SaveAs, Quit, Exists, RegRead, Kill, RmDir, EOF,
               MonthName, GetSpecialFolder, IsEmpty, Date, DeleteFile, MoveFile, DateAdd,
               Error, LanguageID, MultiByteToWideChar, IsNull, SetStringValue, TypeName,
               VarType, Send, CreateShortcut, Popup, MakeSureDirectoryPathExists,
               GetSaveAsFilename, ChDir, ExecuteExcel4Macro, VarPtr, WriteText, FileCopy,
               WriteProcessMemory, RunShell, CopyHere, GetFolder, Hour, _Chr, SaveAs2,
               Chr, CopyFile, GetFile, Paragraphs, UsedRange, CountA, SpecialCells,
               RandBetween, Items, Count, GetParentFolderName, WriteByte, ChrB, ChrW,
               RtlMoveMemory, OnTime, AddItem, Rows, DatePart, FileLen, Sheets, Choose,
               Worksheets, Value, IsObject, Filter, GetRef, BuildPath, CreateFolder,
               Arguments, DateDiff, SetRequestHeader, SetOption, SetTimeouts):
    name = _class.__name__.lower()
    VBA_LIBRARY[name] = _class()

if (log.getEffectiveLevel() == logging.DEBUG):
    log.debug('VBA Library contains: %s' % ', '.join(VBA_LIBRARY.keys()))

# --- VBA CONSTANTS ----------------------------------------------------------

# TODO: 6.1.1 Predefined Enums => complete the library here

for name, value in (

        # 6.1.1.12 VbMsgBoxStyle
        ('vbAbortRetryIgnore', 2),
        ('vbApplicationModal', 0),
        ('vbCritical', 16),
        ('vbDefaultButton1', 0),
        ('vbDefaultButton2', 256),
        ('vbDefaultButton3', 512),
        ('vbDefaultButton4', 768),
        ('vbExclamation', 48),
        ('vbInformation', 64),
        ('vbMsgBoxHelpButton', 16384),
        ('vbMsgBoxRight', 524288),
        ('vbMsgBoxRtlReading', 1048576),
        ('vbMsgBoxSetForeground', 65536),
        ('vbOKCancel', 1),
        ('vbOKOnly', 0),
        ('vbQuestion', 32),
        ('vbRetryCancel', 5),
        ('vbSystemModal', 4096),
        ('vbYesNo', 4),
        ('vbYesNoCancel', 3),

        # 6.1.2.2 Constants Module
        ('vbBack', '\n'),
        ('vbCr', '\r'),
        # From OleVBA the EOL character is just '\n'.
        ('vbCrLf', '\r\n'),
        #('vbCrLf', '\n'),
        ('vbFormFeed', '\f'),
        ('vbLf', '\n'),
        ('vbNewLine', '\r\n'),
        ('vbNullChar', '\x00'),
        ('vbTab', '\t'),
        ('vbVerticalTab', '\v'),
        ('vbNullString', ''),
        ('vbObjectError', -2147221504),

        # Shell Constants
        ('vbHide', 0),
        ('vbNormalFocus', 1),
        ('vbMinimizedFocus.', 2),
        ('vbMaximizedFocus', 3),
        ('vbNormalNoFocus', 4),
        ('vbMinimizedNoFocus', 6),
):
    VBA_LIBRARY[name.lower()] = value
