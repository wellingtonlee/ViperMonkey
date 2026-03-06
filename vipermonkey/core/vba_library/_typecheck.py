"""ViperMonkey: VBA Library — Type checking functions."""
from ._common import *

class IsObject(VbaLibraryFunc):
    """
    IsObject() function (stubbed).
    """

    def eval(self, context, params=None):
        # Say everything is an object and see what happens.
        return True

    def num_args(self):
        return 1

    def return_type(self):
        return "BOOLEAN"

class IsEmpty(VbaLibraryFunc):
    """
    IsEmpty() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return True
        item = params[0]

        # Handle flat out empty values.
        if ((item is None) or (item == "NULL")):
            return True

        # Handle Excel cells.
        if (isinstance(item, dict) and ("value" in item)):
            return self.eval(context, [item["value"]])
        if (item == "empty:u''"):
            return True

        # Handle list type data structures.
        if ((hasattr(item, '__len__')) and (len(item) == 0)):
            return True
        return False

    def num_args(self):
        return 1

class TypeName(VbaLibraryFunc):
    """
    TypeName() function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Return VB type.
        val = params[0]
        if ((val == "NULL") or (val == "")):
            return "Empty"
        if (isinstance(val, bool)):
            return "Boolean"
        if (isinstance(val, str)):
            if (val.lower() == "adodb.stream"):
                return "ADODB.Stream"
            return "String"
        if (isinstance(val, int)):
            return "Integer"
        if (isinstance(val, float)):
            return "Double"
        return "NULL"

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class VarType(VbaLibraryFunc):
    """
    VarType() function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return 0

        # Return VB type.
        val = params[0]
        if ((val == "NULL") or (val == "")):
            return 0
        if (isinstance(val, bool)):
            return 11
        if (isinstance(val, str)):
            return 8
        if (isinstance(val, int)):
            return 2
        if (isinstance(val, float)):
            return 5
        if (isinstance(val, long)):
            return 3
        return 0

    def num_args(self):
        return 1

class IsNumeric(VbaLibraryFunc):
    """
    IsNumeric() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"

        arg = str(params[0])
        try:
            tmp = float(arg)
            return True
        except (ValueError, TypeError):
            return False

class IsNull(VbaLibraryFunc):
    """
    IsNull() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return False
        arg = params[0]
        return ((arg is None) or (arg == "NULL") or (arg == 0) or (arg == ""))

class IsArray(VbaLibraryFunc):
    """
    IsArray() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        return isinstance(params[0], list)

class IsObject(VbaLibraryFunc):
    """
    IsObject() function. Currently stubbed to always return True.
    """

    def eval(self, context, params=None):

        # Say everything is an object.
        return True
