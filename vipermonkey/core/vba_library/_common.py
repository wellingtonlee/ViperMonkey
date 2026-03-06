"""
ViperMonkey: VBA Library — shared imports, state, and helper functions.

All submodules import from this file to get common dependencies.
"""

# --- Standard library imports ------------------------------------------------

import logging
from datetime import datetime
from datetime import date
import time
import array
import math
import base64
import re
from hashlib import sha256
import os
import random
from ..from_unicode_str import *
import decimal
from ..curses_ascii import isprint
import sys
import traceback

from pyparsing import *

# --- Project-internal imports ------------------------------------------------

from .. import vb_str
from ..vba_context import VBA_LIBRARY
from ..vba_object import coerce_to_int
from ..vba_object import eval_arg
from ..vba_object import VbaLibraryFunc
from ..vba_object import VBA_Object
from ..vba_object import excel_col_letter_to_index
from .. import expressions
from .. import excel
from .. import modules
from .. import strip_lines
from ..vba_object import _eval_python
from .. import utils
from ..excel import *

from ..logger import log

# === MODULE-LEVEL STATE =====================================================

# Track the unresolved arguments to the current call.
var_names = None

# Cache for parsed expressions (used by Execute class).
parse_cache = {}

# Global tick counter for GetTickCount.
ticks = 100000

# === HELPER FUNCTIONS =======================================================

def member_access(var, field, globals_calling_scope=None):
    """
    Read a field from an object. Used in Python JIT code.
    """

    # Were we given the globals in the calling scope?
    if (globals_calling_scope is None):
        globals_calling_scope = {}

    # Reading a field from a dict?
    field = str(field)
    field_l = field.lower()
    if (isinstance(var, dict)):
        if (field_l in var):

            # Regular member access.
            return var[field_l]

        # Accessing text field?
        elif ((field_l == "text") and ("value" in var)):
            return var["value"]

        # Accessing cell column?
        elif ((field_l == "column") and ("col" in var)):
            return var["col"] + 1

        # Accessing cell row?
        elif ((field_l == "row") and ("row" in var)):
            return var["row"] + 1

        # Can't find field.
        else:
            return "NULL"

    # Punt and just see if we can return the value of a variable
    # with the same name as the field.
    blah = list(globals().keys())
    blah.sort()
    if (field in locals()):
        return locals[field]
    elif (field in globals()):
        return globals[field]
    elif (field in globals_calling_scope):
        return globals_calling_scope[field]
    else:
        return var

# This function is here to ensure that we return the same global
# shellcode variable as what is updated by emulated VBA functions
# defined in this file.
def get_raw_shellcode_data():
    from .. import vba_context
    return vba_context.shellcode

def run_external_function(func_name, context, params, lib_info):
    """
    Fake running an external DLL function with the given parameters.
    """
    call_str = str(func_name) + "(" + str(params) + ")"
    context.report_action('External Call', call_str, lib_info)
    return 1

def run_function(func_name, context, params):
    """
    Run a VBA library function with the given parameters.
    """

    # Rename python WScript.Shell.Run() calls.
    func_name = func_name.lower()
    if (func_name == "run"):
        func_name = "runshell"

    # Create an object for emulating the function.
    if (func_name not in VBA_LIBRARY):
        return None
    func_obj = VBA_LIBRARY[func_name]
    return func_obj.eval(context, params=params)


def _read_cell(sheet, row, col):

    # Read and process the cell.
    try:
        raw_cell = sheet.cell(row, col)
        r = str(raw_cell).replace("text:", "")
        if (r.startswith("'") and r.endswith("'")):
            r = r[1:-1]
        if (r.startswith('u')):
            r = r[1:]
        if (r.startswith("'") and r.endswith("'") and (len(r) >= 2)):
            r = r[1:-1]
        if (r.startswith('"') and r.endswith('"') and (len(r) >= 2)):
            r = r[1:-1]
        if (r == "empty:u''"):
            r = ""
        if (r.startswith("number:")):
            r = r[len("number:"):]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Excel Read: Cell(" + str(col) + ", " + str(row) + ") = '" + str(r) + "'")
        r = { "value" : r,
              "row" : row + 1,
              "col" : col + 1 }
        return r

    except Exception as e:

        # Failed to read cell.
        return None
