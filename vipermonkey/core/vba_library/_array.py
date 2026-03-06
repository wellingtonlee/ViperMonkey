"""ViperMonkey: VBA Library — Array and collection functions."""
from ._common import *

class Items(VbaLibraryFunc):
    """
    Modified version of Scripting.Dcitionary.Items(). ViperMonkey modifies
    these calls to take the underlying dict containing the items as the 1st
    parameter.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Were we given the dict to work with?
        the_map = None
        index = None
        if ((len(params) >= 2) and (isinstance(params[0], dict))):

            # Dict is 1st parameter.
            the_map = params[0]
            # Item index is 2nd parameter.
            index = coerce_to_int(params[1])

        # Are we reading from a With Scripting.Dictionary?
        elif ((context.with_prefix_raw is not None) and
              (context.contains(str(context.with_prefix_raw)))):

            # Item index is 1st parameter.
            index = coerce_to_int(params[0])

            # Is the With variable value a dict?
            the_map = context.get(str(context.with_prefix_raw))
            if (not isinstance(the_map, dict)):
                return "NULL"
        else:
            return "NULL"

        # Items() handles the added entries in order, so this thing
        # does not act like a standard hash map.
        if ("__ADDED_ITEMS__" not in the_map):
            return "NULL"
        added_items = the_map["__ADDED_ITEMS__"]

        # Is the index valid?
        if ((index < 0) or (index >= len(added_items))):
            return "NULL"

        # Return the item by index.
        return added_items[index]

class AddItem(VbaLibraryFunc):
    """
    ListBox AddItem() VB object method.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 2)):
            return None
        the_list = params[0]
        value = params[1]
        if (not isinstance(the_list, list)):
            return None
        r = list(the_list)
        r.append(value)
        return r

class Add(VbaLibraryFunc):
    """
    Add() VB object method. Currently only adds to Scripting.Dictionary objects is supported.
    """

    def eval(self, context, params=None):
        """
        params[0] = object
        params[1] = key
        params[2] = value
        """

        # Sanity check.
        if ((params is None) or (len(params) < 3)):
            return

        # Get the object (dict), key, and value.
        obj = params[0]
        key = params[1]
        val = params[2]
        if (not isinstance(obj, dict)):
            return

        # Add to the map.
        obj[key] = val
        if ("__ADDED_ITEMS__" not in obj):
            obj["__ADDED_ITEMS__"] = []
        obj["__ADDED_ITEMS__"].append(val)

        # Done.
        return obj

class Array(VbaLibraryFunc):
    """
    Create an array.
    """

    def eval(self, context, params=None):
        r = []
        if ((len(params) == 1) and (params[0] == "NULL")):
            return []
        r = list(params)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Array: return %r" % r)
        return r

class UBound(VbaLibraryFunc):
    """
    UBound() array function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        arr = params[0]
        # TODO: Handle multidimensional arrays.
        if ((arr is None) or (not hasattr(arr, '__len__'))):
            log.error("UBound(" + str(arr) + ") cannot be computed.")
            return 0
        r = len(arr) - 1
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("UBound: return %r" % r)
        return r

class LBound(VbaLibraryFunc):
    """
    LBound() array function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        arr = params[0]
        # TODO: Handle multidimensional arrays.
        r = 0
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("LBound: return %r" % r)
        return r

class Count(VbaLibraryFunc):
    """
    Document or Scripting.Dictionary Count() method.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or
            (len(params) == 0) or
            (not isinstance(params[0], dict))):
            return "NULL"

        # Return the # of Added items.
        # Subtract 1 due to "__ADDED_ITEMS__" entry in dict.
        return (len(params[0]) - 1)

class Arguments(VbaLibraryFunc):
    """
    WScriptShell Arguments field.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        return "_COMMAND_LINE_ARG_" + str(params[0])
