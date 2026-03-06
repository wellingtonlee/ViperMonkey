"""ViperMonkey: VBA Library — String manipulation functions."""
from ._common import *

class _Chr(VbaLibraryFunc):
    """
    Implementation of Chr() and ChrW() used in Python JIT code.
    This is also used under the covers by lib_functions.Chr.eval().
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Chr() called basically on a Cell object?
        param = params[0]
        if (isinstance(param, dict) and ("value" in param)):
            param = param["value"]

        # Proper float conversion for Chr().
        if (isinstance(param, float)):
            param = int(round(param))

        # NOTE: in the specification, the parameter is expected to be an integer
        # But in reality, VBA accepts a string containing the representation
        # of an integer in decimal, hexadecimal or octal form.
        # It also ignores leading and trailing spaces.
        # Examples: Chr("65"), Chr("&65 "), Chr(" &o65"), Chr("  &H65")
        # => need to parse the string as integer
        # It also looks like floating point numbers are allowed.
        try:
            param = coerce_to_int(param)
        except (ValueError, TypeError):
            log.error("%r is not a valid chr() value. Returning ''." % params[0])
            return ''

        # Figure out whether to create a unicode or ascii character.
        converter = chr
        if (param < 0):
            param = param * -1
        if (param > 255):
            converter = unichr

        # Do the conversion.
        try:
            r = converter(param)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Chr(" + str(param) + ") = " + r)
            return r
        except Exception as e:
            log.error(str(e))
            log.error("%r is not a valid chr() value. Returning ''." % param)
            return ""

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class Chr(_Chr):
    pass

class ChrB(_Chr):
    pass

class ChrW(_Chr):
    pass

class Len(VbaLibraryFunc):
    """
    Len() function.
    """

    def eval(self, context, params=None):
        if (isinstance(params[0], int)):
            return len(str(params[0]))
        val = utils.str_convert(params[0])
        if (hasattr(params[0], '__len__')):

            # Is this a string?
            if (isinstance(val, str)):

                # If this is VBScript strings are sensible and we can just return the length.
                if (context.is_vbscript):
                    return len(val)

                # Convert the string to a VbStr to handle mized ASCII/wide char weirdness.
                vb_val = vb_str.VbStr(val, context.is_vbscript)
                return vb_val.len()

            # Something with a length that is not a string.
            else:
                return len(val)
        else:
            log.error("Len: " + str(type(params[0])) + " object has no len(). Returning 0.")
            return 0

    def num_args(self):
        return 1

class LenB(VbaLibraryFunc):
    """
    LenB() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return 0
        # TODO: Somehow take the default character set into account.
        try:
            return len(params[0])
        except TypeError:
            return 0

    def num_args(self):
        return 1

class Mid(VbaLibraryFunc):
    """
    6.1.2.11.1.25 Mid / MidB function

    IMPORTANT NOTE: Not to be confused with the Mid statement 5.4.3.5!
    """

    def eval(self, context, params=None):
        if (params is None):
            log.error("Invalid arguments " + str(params) + " to Mid().")
            return ""
        if ((len(params) > 0) and (params[0] == "ActiveDocument")):
            params = params[1:]
        if (params is None):
            log.error("Invalid arguments " + str(params) + " to Mid().")
            return ""
        if (len(params) not in (2,3)):
            log.error("Invalid arguments " + str(params) + " to Mid().")
            return ""
        s = params[0]
        # "If String contains the data value Null, Null is returned."
        if ((s is None) or (s == "NULL")): return "\x00"
        # If start is NULL, NULL is also returned.
        if ((params[1] is None) or (params[1] == "NULL")): return "\x00"
        if not isinstance(s, str):
            s = utils.str_convert(s)
        start = 0
        try:
            start = utils.int_convert(params[1])
        except (ValueError, TypeError):
            pass

        # Convert the string to a VbStr to handle mized ASCII/wide char weirdness.
        vb_s = None
        s_len = len(s)
        if (not context.is_vbscript):
            vb_s = vb_str.VbStr(s, context.is_vbscript)
            s_len = vb_s.len()

        # "If Start is greater than the number of characters in String,
        # Mid returns a zero-length string ("")."
        if (start > s_len):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Mid: start>len(s) => return ""')
            return ''

        # What to do when start<=0 is not specified:
        if (start <= 0):
            return "NULL"

        # If length not specified, return up to the end of the string:
        if (len(params) == 2):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Mid: no length specified, return s[%d:]=%r' % (start-1, s[start-1:]))
            return s[start-1:]
        length = 0
        try:
            length = utils.int_convert(params[2])
        except (ValueError, TypeError):
            pass

        # "If omitted or if there are fewer than Length characters in the text
        # (including the character at start), all characters from the start
        # position to the end of the string are returned."
        if start+length-1 > s_len:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Mid: start+length-1>len(s), return s[%d:]' % (start-1))
            if context.is_vbscript:
                return s[start-1:]
            else:
                return vb_s.get_chunk(start - 1, vb_s.len()).to_python_str()

        # What to do when length<=0 is not specified:
        if length <= 0:
            return ''

        # Regular Mid().
        if context.is_vbscript:
            r = s[start - 1:start-1+length]
        else:
            r = vb_s.get_chunk(start - 1, start - 1 + length).to_python_str()

        # Done.
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('Mid: return s[%d:%d]=%r' % (start - 1, start-1+length, r))
        return r

    def num_args(self):
        return 2

    def return_type(self):
        return "STRING"

class MidB(Mid):
    pass

class Left(VbaLibraryFunc):
    """
    Left function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 2)):
            return "NULL"
        if (len(params) > 2):
            params = params[-2:]
        s = params[0]
        if s is None: return None

        # Arg should be a string.
        s = utils.safe_str_convert(s)

        # Don't modify the "**MATCH ANY**" special value.
        if (s.strip() == "**MATCH ANY**"):
            return s

        # "If String contains the data value Null, Null is returned."
        start = 0
        try:
            start = utils.int_convert(params[1])
        except (ValueError, TypeError):
            pass

        # Convert the string to a VbStr to handle mized ASCII/wide char weirdness.
        vb_s = vb_str.VbStr(s, context.is_vbscript)

        # "If Start is greater than the number of characters in String,
        # Left returns the whole string.
        if (start > vb_s.len()):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Left: start>len(s) => return s')
            return s

        # Return empty string if start <= 0.
        if (start <= 0):
            return ""

        # Return characters from start of string.
        #r = s[:start]
        r = vb_s.get_chunk(0, start).to_python_str()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('Left: return s[0:%d]=%r' % (start, r))
        return r

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class Right(VbaLibraryFunc):
    """
    Right function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 2)):
            return "NULL"
        if (len(params) > 2):
            params = params[-2:]
        s = params[0]

        # Don't modify the "**MATCH ANY**" special value.
        if (str(s).strip() == "**MATCH ANY**"):
            return s

        # "If String contains the data value Null, Null is returned."
        if s is None: return None
        if not isinstance(s, str):
            s = str(s)
        start = 0
        try:
            start = utils.int_convert(params[1])
        except (ValueError, TypeError):
            pass

        # Convert the string to a VbStr to handle mized ASCII/wide char weirdness.
        vb_s = vb_str.VbStr(s, context.is_vbscript)

        # "If Start is greater than the number of characters in String,
        # Right returns the whole string.
        if (start > vb_s.len()):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug('Right: start>len(s) => return s')
            return s

        # Return empty string if start <= 0.
        if (start <= 0):
            return ""

        # Return characters from end of string.
        #r = s[(len(s) - start):]
        r = vb_s.get_chunk(vb_s.len() - start, vb_s.len()).to_python_str()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug('Right: return s[%d:]=%r' % (start, r))
        return r

    def num_args(self):
        return 2

    def return_type(self):
        return "STRING"

class Trim(VbaLibraryFunc):
    """
    Trim() string function.
    """

    def eval(self, context, params=None):

        # Sanity check arguments.
        if ((params is None) or (len(params) == 0)):
            log.error("Invalid paramater to Trim().")
            return ""

        # Trim the string.
        r = str(params[0]).strip()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Trim: return %r" % r)
        return r

    def return_type(self):
        return "STRING"

class RTrim(VbaLibraryFunc):
    """
    RTrim() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        r = None
        if (isinstance(params[0], int)):
            r = str(params[0])
        else:
            r = params[0].rstrip()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("RTrim: return %r" % r)
        return r

    def return_type(self):
        return "STRING"

class LTrim(VbaLibraryFunc):
    """
    LTrim() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return ""
        r = None
        if (isinstance(params[0], int)):
            r = str(params[0])
        else:
            r = params[0].lstrip()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("LTrim: return %r" % r)
        return r

    def return_type(self):
        return "STRING"

class AscW(VbaLibraryFunc):
    """
    AscW() character function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        c = params[0]
        if (c == "NULL"):
            return 0
        if (isinstance(c, int)):
            r = c
        else:
            c = str(c)
            if (len(c) > 0):
                r = ord(str(c)[0])
            else:
                r = 0
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("AscW: return %r" % r)
        return r

class AscB(AscW):
    pass

class StrComp(VbaLibraryFunc):
    """
    StrComp() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 2)):
            return "NULL"
        s1 = params[0]
        s2 = params[1]
        method = 0
        if (len(params) >= 3):
            try:
                method = utils.int_convert(params[2])
            except Exception as e:
                log.error("StrComp: Invalid comparison method. " + str(e))
                pass
        if (method == 0):
            s1 = s1.lower()
            s2 = s2.lower()
        if (s1 == s2):
            return 0
        if (s1 < s2):
            return -1
        return 1

class StrPtr(VbaLibraryFunc):
    """
    External StrPtr() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Do we have a variable name?
        arg = str(params[0])
        if (arg.startswith("&")):

            # Just return the name of the variable being pointed to by the string pointer.
            return arg[1:]

        # We don't have a variable, so just turn it into a "pointer".
        return ("&" + str(params[0]))

    def return_type(self):
        return "STRING"

class StrConv(VbaLibraryFunc):
    """
    StrConv() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # TODO: Actually implement this properly.

        # Get the conversion type to perform.
        conv = None
        if (len(params) > 1):
            conv = utils.int_convert(eval_arg(params[1], context=context))

        # Do the conversion.
        r = params[0]
        if (isinstance(r, str)):
            if (conv):
                if (conv == 1):
                    r = r.upper()
                if (conv == 2):
                    r = r.lower()
                if (conv == 64):

                    # We are converting the string to unicode. ViperMonkey assumes
                    # unless otherwise noted that all strings are unicode. Make sure
                    # that the string is represented as a regular str object so that
                    # it is treated as unicode.
                    r = str(r)

                if (conv == 128):

                    # The string is being converted from unicode to ascii. Mark this
                    # by representing the string with the from_unicode_str class.
                    r = from_unicode_str(r)

        elif (isinstance(r, list)):

            # Handle list of ASCII values.
            all_int = True
            for i in r:
                if (not isinstance(i, int)):
                    all_int = False
                    break
            if (all_int):
                tmp = ""
                for i in r:
                    if (i < 0):
                        continue
                    try:
                        tmp += chr(i)
                        #if (conv == 64):
                        #    tmp += "\0"
                    except (ValueError, OverflowError):
                        pass
                r = tmp

            else:
                log.error("StrConv: Unhandled type.")
                r = ''

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("StrConv: return %r" % r)
        return r

    def return_type(self):
        return "STRING"

class Split(VbaLibraryFunc):
    """
    Split() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # TODO: Actually implement this properly.
        string = utils.safe_str_convert(params[0])
        sep = " "
        if ((len(params) > 1) and
            (isinstance(params[1], str)) and
            (len(params[1]) > 0)):
            sep = str(params[1])

        # Let's assume that splitting on char 0x00 means break
        # up into individual characters.
        if (sep == chr(0)):
            r = []
            for c in string:
                r.append(c)
            return r

        r = string.split(sep)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Split: return %r" % r)
        #print "SPLIT!!"
        #print r
        return r

class StrReverse(VbaLibraryFunc):
    """
    StrReverse() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        # TODO: Actually implement this properly.
        string =''
        if ((params[0] is not None) and (len(params) > 0)):
            string = params[0]
            if ((not isinstance(params[0], str)) and
                (not isinstance(params[0], str))):
                string = str(params[0])
        r = string[::-1]
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("StrReverse: return %r" % r)
        return r

    def return_type(self):
        return "STRING"

class Filter(VbaLibraryFunc):
    """
    The VBA Filter function returns a subset of a supplied string array, based on supplied criteria.

    The syntax of the function is: Filter( SourceArray, Match, [Include], [Compare] )

    The function arguments are:

    SourceArray	-	The original array of Strings, that you want to filter.
    Match	-	The string that you want to search for, within each element of the supplied SourceArray.
    [Include]	-
    An option boolean argument that specifies whether the returns array should consist of elements that include or do not include the supplied Match String.

    This can have the value True or False, meaning:

    True	-	Only return elements that include the Match String
    False	-	Only return elements that do not include the Match String
    If the [Include] argument is omitted, it takes on the default value True.

    [Compare]	-
    An optional argument, specifying the type of String comparison to make.

    This can be any of the following values:

    vbBinaryCompare	-	performs a binary comparison (0)
    vbTextCompare	-	performs a text comparison (1)
    vbDatabaseCompare	-	performs a database comparison (2)
    If omitted, the [Compare] argument takes on the default value vbBinaryCompare.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 2)):
            log.warning("Too few arguments to Filter() call. Returning NULL")
            return "NULL"

        # SourceArray and Match are required.
        source_array = params[0]
        match = params[1]
        if (not isinstance(source_array, list)):
            log.warning("SourceArray argument to Filter() call not an array. Returning NULL")
            return "NULL"
        if (not isinstance(match, str)):
            log.warning("Match argument to Filter() call not a string. Returning NULL")
            return "NULL"

        # Include/exclude argument is optional.
        include = True
        if (len(params) >=3):
            include = params[2]
            if (not isinstance(match, str)):
                log.warning("Include/exclude argument to Filter() call not a boolean. Returning NULL")
                return "NULL"

        # Compare argument is optional.
        compare = 0
        if (len(params) >=4):
            compare = params[3]
            if (not isinstance(compare, int)):
                log.warning("Compare argument to Filter() call not an int. Returning NULL")
                return "NULL"

        # Currently only handling vbBinaryCompare.
        if (compare != 0):
            log.warning("Only handling Compare == vbBinaryCompare in Filter() call. Returning NULL")
            return "NULL"

        # Find the items to return.
        r = []
        for s in source_array:
            found_match = (match in s)
            if (not include):
                found_match = not found_match
            if found_match:
                r.append(s)

        # Done.
        return r

class Replace(VbaLibraryFunc):
    """
    Replace() string function.

    The Replace function syntax has these named arguments:

    expression	Required. String expression containing substring to replace.
    find	Required. Substring being searched for.
    replace	Required. Replacement substring.
    start	Optional. Start position for the substring of expression to be searched and returned. If omitted, 1 is assumed.
    count	Optional. Number of substring substitutions to perform. If omitted, the default value is -1, which means, make all possible substitutions.
    compare	Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. See Settings section for values.
    """

    def eval(self, context, params=None):
        if (params is None):
            return ""
        if (len(params) < 3):
            if (len(params) > 0):
                return params[0]
            else:
                return ""
        # TODO: Handle start, count, and compare parameters.
        string = params[0]
        # Handle Excel cells.
        if (isinstance(string, dict) and ("value" in string)):
            string = str(string["value"])
        if (string is None):
            string = ''
        string = str(string)
        pat = str(params[1])
        if ((pat is None) or (pat == '')):
            return string
        rep = str(params[2])
        if ((rep is None) or (rep == 0) or (rep == "NULL")):
            rep = ''

        # Wide string to change and not wide char pattern/replacement?
        if (vb_str.is_wide_str(string) and
            ((not vb_str.is_wide_str(pat)) or (not vb_str.is_wide_str(rep)))):

            # Convert the string to change to ASCII.
            log.warning("Replace() called on wide string w. ASCII pattern and replacement. Converting to ASCII ...")
            string = vb_str.convert_wide_to_ascii(string)

        # regex replacement?
        if (params[-1] == "<-- USE REGEX -->"):

            # Don't do a regex replacement of everything.
            if (pat.strip() != "."):
                try:
                    pat1 = pat.replace("$", "\\$").replace("-", "\\-")
                    fix_dash_pat = r"(\[.\w+)\\\-(\w+\])"
                    pat1 = re.sub(fix_dash_pat, r"\1-\2", pat1)
                    fix_dash_pat1 = r"\((\w+)\\\-(\w+)\)"
                    pat1 = re.sub(fix_dash_pat1, r"[\1-\2]", pat1)
                    rep = re.sub(r"\$(\d)", r"\\\1", rep)
                    r = re.sub(pat1, rep, string)
                except Exception as e:
                    log.error("Regex replace " + str(params) + " failed. " + str(e))
                    r = string

        # Regular string replacement?
        else:
            r = string.replace(pat, rep)

        # Done.
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Replace: return %r" % r)
        return r

    def return_type(self):
        return "STRING"

class Join(VbaLibraryFunc):
    """
    Join() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        strings = params[0]
        sep = " "
        if (len(params) > 1):
            sep = str(params[1])
        if (sep == "NULL"):
            sep = ""
        r = ""
        if (isinstance(strings, list)):
            for s in strings:
                tmp_s = utils.safe_str_convert(s)
                r += tmp_s + sep
        else:
            r = str(strings)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Join: return %r" % r)
        return r

    def return_type(self):
        return "STRING"

class InStr(VbaLibraryFunc):
    """
    InStr() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 2)):
            return "NULL"

        # Were we given a start position?
        start = 0
        s1 = params[0]
        if (s1 is None):
            s1 = ''
        s2 = params[1]
        if (s2 is None):
            s2 = ''
        if (isinstance(params[0], int)):
            if (len(params) < 3):
                return False
            start = params[0] - 1
            if (start < 0):
                start = 0
            s1 = params[1]
            s2 = params[2]

        # Were we given a search type?
        search_type = 1
        if (isinstance(params[-1], int)):
            search_type = params[-1]
            if (search_type not in (0, 1)):
                search_type = 1

        # Only works on lists or strings.
        if ((not isinstance(s1, list)) and (not isinstance(s1, str))):
            return None
        if ((not isinstance(s2, list)) and (not isinstance(s2, str))):
            return None

        # Always say we found the substring if the string in which we
        # are searching is the special wildcard string.
        if (s1 == "**MATCH ANY**"):
            return 1

        # TODO: Figure out how VB binary search works. For now just do text search.
        r = None
        if (len(s1) == 0):
            r = 0
        elif (len(s2) == 0):
            r = start
        elif (start > len(s1)):
            r = 0
        else:
            if (s2 in s1[start:]):
                r = s1[start:].index(s2) + start + 1
            else:
                r = 0
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("InStr: %r returns %r" % (self, r))
        return r

class InStrRev(VbaLibraryFunc):
    """
    InStrRev() string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 2)):
            return "NULL"

        # Were we given a start position?
        start = 0
        s1 = params[0]
        if (s1 is None):
            s1 = ''
        s2 = params[1]
        if (s2 is None):
            s2 = ''
        if (isinstance(params[0], int)):
            start = params[0] - 1
            if (start < 0):
                start = 0
            if (len(params) < 3):
                return 0
            s1 = params[1]
            s2 = params[2]

        # Were we given a search type?
        s1 = str(s1)
        s2 = str(s2)
        search_type = 1
        if (isinstance(params[-1], int)):
            search_type = params[-1]
            if (search_type not in (0, 1)):
                search_type = 1

        # TODO: Figure out how VB binary search works. For now just do text search.
        r = None
        if (len(s1) == 0):
            r = 0
        elif (len(s2) == 0):
            r = start
        elif (start > len(s1)):
            r = 0
        else:
            if (s2 in s1):
                r = s1[start:].rindex(s2) + start + 1
            else:
                r = 0
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("InStr: %r returns %r" % (self, r))
        return r


class String(VbaLibraryFunc):
    """
    String() repeated character string creation function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 2)):
            return "NULL"
        r = ''
        try:
            num = utils.int_convert(params[0])
            char = params[1]
            r = char * num
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("String: %r returns %r" % (self, r))
        return r

class Str(VbaLibraryFunc):
    """
    Str() convert number to string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return ""
        r = str(params[0])
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Str: %r returns %r" % (self, r))
        return r

class Val(VbaLibraryFunc):
    """
    Val() convert string to number function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"

        # Sanity check.
        if ((params[0] is None) or (not isinstance(params[0], str))):
            r = ''
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Str: %r returns %r" % (self, r))
            return r

        # Ignore whitespace.
        tmp = utils.str_convert(params[0]).strip().replace(" ", "")

        # No nulls.
        tmp = tmp.replace("\x00", "")

        # The VB Val() function is ugly. Look for VB hex encoding.
        nums = re.compile(r"&[Hh][0-9A-Fa-f]+")
        matches = nums.search(tmp)
        if (hasattr(matches, "group")):
            tmp = nums.search(tmp).group(0).replace("&H", "0x").replace("&h", "0x")
            r = float(int(tmp, 16))
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Val: %r returns %r" % (self, r))
            return r

        # The VB Val() function is ugly. Try to use a regular expression to pick out
        # the 1st valid number string.
        nums = re.compile(r"[+-]?\d+(?:\.\d+)?")
        matches = nums.search(tmp)
        if (hasattr(matches, "group")):
            tmp = nums.search(tmp).group(0)

            # Convert this to a float or int.
            r = None
            if ("." in tmp):
                r = float(tmp)
            else:
                r = int(tmp)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Val: %r returns %r" % (self, r))
            return r

        # Can't find a valid number to convert. This is probably incorrect behavior.
        r = 0
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Val: Invalid Value: %r returns %r" % (self, r))
        return r

class Base64Decode(VbaLibraryFunc):
    """
    Base64Decode() function used by some malware. Note that this is not part of Visual Basic.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        txt = params[0]
        if (txt is None):
            txt = ''
        r = utils.b64_decode(txt)
        if (r is None):
            r = "NULL"
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Base64Decode: %r returns %r" % (self, r))
        return r

class Base64DecodeString(Base64Decode):
    pass

class CleanString(VbaLibraryFunc):
    """
    CleanString() function removes certain characters from the character stream, or translates them
    https://docs.microsoft.com/en-us/office/vba/api/word.application.cleanstring
    """

    def eval(self,context,params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        txt=params[0]
        if (txt is None):
            txt = ''
        if isinstance(txt,str):
            a = [c for c in txt]
            for i in range(len(a)):
                c = a[i]
                if ord(c) == 7:
                    if i>0 and ord(a[i-1]) == 13:
                        a[i] = chr(9)
                    else:
                        a[i] = ''
                if ord(c) == 10:
                    if i>0 and ord(a[i-1]) == 13:
                        a[i] = ''
                    else:
                        a[i] = chr(13)
                if ord(c) == 31 or ord(c) == 172 or ord(c) == 182:
                    a[i] = ''
                if ord(c) == 160 or ord(c) == 176 or ord(c) == 183:
                    a[i] = chr(32)
            r = "".join(a)
        else:
            # punt for things like CleanString(99), which shows up as an integer
            r = txt
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("CleanString: %r returns %r" % (self,r))
        return r

    def return_type(self):
        return "STRING"

class Space(VbaLibraryFunc):
    """
    Space() string function.
    """

    def eval(self, context, params=None):
        n = utils.int_convert(params[0])
        r = " " * n
        return r

    def return_type(self):
        return "STRING"

class UCase(VbaLibraryFunc):
    """
    UCase() string function.
    """

    def eval(self, context, params=None):
        r = str(params[0]).upper()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("UCase: %r returns %r" % (self, r))
        return r

    def return_type(self):
        return "STRING"

class LCase(VbaLibraryFunc):
    """
    LCase() string function.
    """

    def eval(self, context, params=None):
        r = str(params[0]).lower()
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("LCase: %r returns %r" % (self, r))
        return r

    def return_type(self):
        return "STRING"

class Unescape(VbaLibraryFunc):
    """
    Unescape() strin unescaping method (stubbed).
    """

    def eval(self, context, params=None):

        # Get the string to unescape.
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        s = str(params[0])

        # It reverses the transformation performed by the Escape
        # method by removing the escape character ("\") from each
        # character escaped by the method. These include the \, *, +,
        # ?, |, {, [, (,), ^, $, ., #, and white space characters.
        s = s.replace("\\\\", "\\")
        s = s.replace("\\*", "*")
        s = s.replace("\\+", "+")
        s = s.replace("\\?", "?")
        s = s.replace("\\|", "|")
        s = s.replace("\\{", "{")
        s = s.replace("\\[", "[")
        s = s.replace("\\(", "(")
        s = s.replace("\\)", ")")
        s = s.replace("\\^", "^")
        s = s.replace("\\$", "$")
        s = s.replace("\\.", ".")
        s = s.replace("\\#", "#")
        s = s.replace("\\ ", " ")
        # TODO: Figure out more whitespace characters.

        # In addition, the Unescape method unescapes the closing
        # bracket (]) and closing brace (}) characters.
        if ("\\]" in s):
            start = s.rindex("\\]")
            end = start + len("\\]")
            s = s[:start] + "]" + s[end:]
        if ("\\}" in s):
            start = s.rindex("\\}")
            end = start + len("\\}")
            s = s[:start] + "}" + s[end:]

        # It replaces the hexadecimal values in verbatim string
        # literals with the actual printable characters. For example,
        # it replaces @"\x07" with "\a", or @"\x0A" with "\n". It
        # converts to supported escape characters such as \a, \b, \e,
        # \n, \r, \f, \t, \v, and alphanumeric characters.
        #
        # TODO: Do the hex unescaping.

        # Not documented, but it looks like %xx% is also handled as hex
        # unescaping.
        pat = r"%([0-9a-fA-F][0-9a-fA-F])"
        hex_strs = re.findall(pat, s)
        for h in hex_strs:
            s = s.replace("%" + h, chr(int("0x" + h, 16)))

        # Return the unsescaped string.
        return s

    def return_type(self):
        return "STRING"
