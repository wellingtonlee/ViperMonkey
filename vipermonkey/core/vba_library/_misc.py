"""ViperMonkey: VBA Library — Miscellaneous functions (dialog, object, registry, etc.)."""
from ._common import *

class Choose(VbaLibraryFunc):
    """
    Choose() choice function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 2)):
            return "NULL"

        # The index of the value to pick should be the 1st
        # argument.
        index = None
        try:
            index = int(params[0])
        except Exception as e:
            log.warning("Invalid index passed to Choice(). Returning NULL. " + str(e))
            return "NULL"

        # Is the desired value in the choice list?
        if (index > (len(params) + 1)):
            log.warning("Invalid index passed to Choice(). Too large.")
            return "NULL"

        # Return the choice value.
        return params[index]

class GetSaveAsFilename(VbaLibraryFunc):
    """
    GetSaveAsFilename() function (stubbed).
    """

    def eval(self, context, params=None):
        return 'C:\\Users\\admin\\AppData\\Local\\Faked_SaveAs_File_Name.dat'

    def num_args(self):
        return 0

    def return_type(self):
        return "STRING"

class MsgBox(VbaLibraryFunc):
    """
    6.1.2.8.1.13 MsgBox
    """

    def eval(self, context, params=None):
        context.report_action('Display Message', params[0], 'MsgBox', strip_null_bytes=True)
        return 1  # vbOK

    def num_args(self):
        return 1

class Quit(VbaLibraryFunc):
    """
    Wscript.Quit(). Just keeps going.
    """

    def eval(self, context, params=None):
        log.warning("Ignoring Wscript.Quit() call. Execution is continuing...")
        return 1

    def num_args(self):
        return 0

class Switch(VbaLibraryFunc):
    """
    Switch() logic flow function.
    """

    def eval(self, context, params=None):

        # We need an even number of parameters.
        if ((len(params) == 0) or
            (len(params) % 2 != 0)):
            return 'NULL'

        # Return the 1st true case.
        pos = 0
        while (pos < (len(params) - 1)):
            if params[pos] is True:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Switch(%r): return %r" % (self, params[pos + 1]))
                return params[pos + 1]
            pos += 2

        # If we get here nothing is true.
        return 'NULL'

    def num_args(self):
        return 2

class Error(VbaLibraryFunc):
    """
    Stubbed Error() method.
    """

    def eval(self, context, params=None):
        return "Some error message..."

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class Exists(VbaLibraryFunc):
    """
    Document or Scripting.Dictionary Exists() method.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return False

        # Scripting.Dictionary Exists()?
        if ((len(params) == 2) and (isinstance(params[0], dict))):
            r = (params[1] in params[0])
            return r

        # Document Exists(). Default to False.
        return False

class RegWrite(VbaLibraryFunc):
    """
    RegWrite() function.
    """

    def eval(self, context, params=None):
        context.report_action("Registry Write", str(params), "Registry Write", strip_null_bytes=True)
        return "NULL"

class SetStringValue(VbaLibraryFunc):
    """
    SetStringValue() function.
    """

    def eval(self, context, params=None):
        context.report_action("Registry Write", str(params), "Set String Value", strip_null_bytes=True)
        return "NULL"

class GetRef(VbaLibraryFunc):
    """
    GetRef() function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Get the function.
        obj_name = str(params[0])
        if (not context.contains(obj_name)):
            return "NULL"
        return context.get(obj_name)

class RegRead(VbaLibraryFunc):
    """
    RegRead() function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 1)):
            return ""

        # Fake some registry reads.
        key = str(params[0])
        context.report_action('Read Registry', key, 'RegRead', strip_null_bytes=True)
        if (key == 'HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment\\PROCESSOR_ARCHITECTURE'):
            return "x86"

        # Not faked.
        return ""

class IIf(VbaLibraryFunc):
    """
    IIf() if-like function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 3)):
            return "NULL"
        guard = params[0]
        true_part = params[1]
        false_part = params[2]
        if (guard):
            return true_part
        else:
            return false_part

class CallByName(VbaLibraryFunc):
    """
    CallByName() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 3)):
            return "NULL"

        # Report interesting external commands run.
        cmd = str(params[1])
        obj = str(params[0])
        args = ''
        if (len(params) >= 4):
            args = params[3]
        if (("run" in cmd.lower()) or ("create" in cmd.lower()) or ("wscript.shell" in obj.lower())):
            context.report_action("Run", args, 'Interesting Function Call', strip_null_bytes=True)
        for pos in range(0, len(params)):
            if ((str(params[pos]).lower() == "wscript") and ((pos + 1) < len(params))):
                context.report_action("Run", params[pos + 1], 'Interesting Function Call', strip_null_bytes=True)
        # CallByName("['WinHttp.WinHttpRequest.5.1', 'Open', 1, 'GET', 'http://deciodc.org/bin/office1...")
        if ((("Open" in cmd) and ("WinHttpRequest" in obj)) or
            ((len(params) > 5) and (str(params[3]).lower() == "get"))):
            url = str(params[4])
            if (url.startswith("tp://")):
                url = "ht" + url
            context.report_action("GET", url, 'Interesting Function Call', strip_null_bytes=True)
        # CallByName(([DoBas, 'Arguments', VbLet, aas], {}))
        if ((cmd == "Arguments") or (cmd == "Path")):
            context.report_action("CallByName", args, 'Possible Scheduled Task Setup', strip_null_bytes=True)
        # CallByName(['shell.application', 'shellexecute', 1, ...
        if (cmd.lower() == "shellexecute"):
            if (len(params) > 4):
                run_cmd = str(params[3]) + " " + str(params[4])
                context.report_action('Execute Command', run_cmd, 'Shell function', strip_null_bytes=True)

        # Are we using this to read text from a GUI element?
        if ((cmd == "Tag") or (cmd == "Text")):

            # Looks like it. Lets return the text. This is read from a for variable.
            try:
                return context.get(str(params[0]) + "." + cmd)
            except KeyError:
                pass

        # Opening a file?
        if (cmd.lower() == "createtextfile"):

            # Open the file.
            opener = CreateTextFile()
            opener.eval(context, [args])

        # Writing to a file?
        if (cmd.lower() == "writeline"):

            # Write to the file.
            writer = WriteLine()
            writer.eval(context, [args])

        # Closing a file?
        if (cmd.lower() == "close"):

            # Close the file.
            closer = Close()
            closer.eval(context, [args])

        # Do nothing.
        return None

class Raise(VbaLibraryFunc):
    """
    Raise() exception/error function.
    """

    def eval(self, context, params=None):
        msg = "Raise exception " + str(params)
        context.set_error(msg)

class KeyString(VbaLibraryFunc):
    """
    KeyString() function.
    """

    def eval(self, context, params=None):

        # Key string value map.
        key_vals = {
            1 : "Left Button",
            2 : "Right Button",
            3 : "Cancel",
            4 : "Middle Button",
            8 : "Backspace",
            9 : "Tab",
            12 : "Clear (Num 5)",
            13 : "Return",
            16 : "Shift",
            17 : "Control",
            18 : "Alt",
            19 : "Pause",
            20 : "Caps Lock",
            27 : "Esc",
            32 : "Space",
            33 : "Page Up",
            34 : "Page Down",
            35 : "End",
            36 : "Home",
            37 : "Left",
            38 : "Up",
            39 : "Right",
            40 : "Down",
            41 : "Not Avail",
            42 : "Not Avail",
            43 : "Not Avail",
            44 : "Print Screen",
            45 : "Insert",
            46 : "Del",
            47 : "Not Avail",
            48 : "0",
            49 : "1",
            50 : "2",
            51 : "3",
            52 : "4",
            53 : "5",
            54 : "6",
            55 : "7",
            56 : "8",
            57 : "9",
            65 : "A",
            66 : "B",
            67 : "C",
            68 : "D",
            69 : "E",
            70 : "F",
            71 : "G",
            72 : "H",
            73 : "I",
            74 : "J",
            75 : "K",
            76 : "L",
            77 : "M",
            78 : "N",
            79 : "O",
            80 : "P",
            81 : "Q",
            82 : "R",
            83 : "S",
            84 : "T",
            85 : "U",
            86 : "V",
            87 : "W",
            88 : "X",
            89 : "Y",
            90 : "Z",
            96 : "Num 0",
            97 : "Num 1",
            98 : "Num 2",
            99 : "Num 3",
            100 : "Num 4",
            101 : "Num 5",
            102 : "Num 6",
            103 : "Num 7",
            104 : "Num 8",
            105 : "Num 9",
            106 : "Num *",
            107 : "Num +",
            108 : "Not Avail",
            109 : "Num -",
            110 : "Num .",
            111 : "Num /",
            112 : "F1",
            113 : "F2",
            114 : "F3",
            115 : "F4",
            116 : "F5",
            117 : "F6",
            118 : "F7",
            119 : "F8",
            120 : "F9",
            121 : "F10",
            122 : "F11",
            123 : "F12",
            124 : "F13",
            125 : "F14",
            126 : "F15",
            127 : "F16",
            128 : "F17",
            129 : "F18",
            130 : "F19",
            131 : "F20",
            132 : "F21",
            133 : "F22",
            134 : "F23",
            135 : "F24",
            144 : "Num Lock",
            145 : "Scroll Lock",
            160 : "Shift",
            161 : "Shift",
            162 : "Ctrl",
            163 : "Ctrl",
            164 : "Alt",
            165 : "Alt",
            172 : "M",
            173 : "D",
            174 : "C",
            175 : "B",
            176 : "P",
            177 : "Q",
            178 : "J",
            179 : "G",
            183 : "F",
            186 : ";",
            187 : "=",
            188 : ",",
            189 : "-",
            190 : ".",
            191 : "/",
            192 : "`",
            194 : "F15",
            219 : "[",
            220 : "\\",
            221 : "]",
            222 : "'",
            226 : "\\"
        }

        v1 = None
        v2 = None
        try:
            v1 = int(params[0])
            if (len(params) >= 2):
                v2 = int(params[1])
        except Exception as e:
            log.error("KeyString: Invalid args " + str(params) + ". " + str(e))
            return ""

        r = ""
        if (v1 in key_vals):
            r += key_vals[v1]
        if (v2 is not None):
            r += ","
            if (v2 in key_vals):
                r += key_vals[v2]

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("KeyString: args = " + str(params) + ", return " + r)
        return r

    def return_type(self):
        return "STRING"

class CreateShortcut(VbaLibraryFunc):
    """
    CreateShortcut() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return
        path = params[0]
        context.report_action("Shortcut Creation", path, 'Shortcut Created', strip_null_bytes=True)

class CreateObject(VbaLibraryFunc):
    """
    CreateObject() function (stubbed).
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return ""

        # Track contents of data written to 'ADODB.Stream'.
        obj_type = utils.safe_str_convert(params[0]).lower()
        if (obj_type == 'ADODB.Stream'.lower()):
            context.open_file('ADODB.Stream')

        # Handle certain object types.
        if (obj_type == "Scripting.Dictionary".lower()):
            r = {}
            # Track the added items in order as well as by key.
            r["__ADDED_ITEMS__"] = []
            return r

        # Just return a string representation of the name of the object
        # being created.
        return str(obj_type)

class LanguageID(VbaLibraryFunc):
    """
    Stubbed LanguageID() reference.
    """

    def eval(self, context, params=None):

        # This is usually used for gating, so have it match anything.
        return "**MATCH ANY**"

    def num_args(self):
        return 1

class Assert(VbaLibraryFunc):
    """
    Assert() debug function. Stubbed.
    """

    def eval(self, context, params=None):
        pass

class Popup(VbaLibraryFunc):
    """
    Popup() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return
        msg = params[0]
        context.report_action("Popup", str(msg), '')

class Print(VbaLibraryFunc):
    """
    Debug.Print function.
    """

    def _handle_file_print(self, context, params):

        # Sanity check.
        if (len(params) != 2):
            log.warning("Wrong # of arguments for Print " + str(params))
            return

        # 1st arg should be file ID.
        fileid = "#" + str(params[0])

        # 2nd arg should be data to write.
        data = utils.safe_str_convert(params[1])

        # Try writing the file.
        context.write_file(fileid, data)

    def eval(self, context, params=None):

        # Sanity check.
        if (params is None):
            return

        # Print #NN to a file ID?
        if (len(params) == 2):
            return self._handle_file_print(context, params)

        # Regular Debug.Print() ?
        if (len(params) != 1):
            log.warning("Wrong # of arguments for Print " + str(params))
            return

        # Save writes that look like they are writing URLs.
        data_str = utils.safe_str_convert(params[0])
        if (("http:" in data_str.lower()) or ("https:" in data_str.lower())):
            context.report_action('Write URL', data_str, 'Debug Print')

        if (params[0] is not None):
            if (not context.throttle_logging):
                context.report_action("Debug Print", data_str, '')

class Debug(Print):
    """
    Debug() debugging function.
    """
    pass

class Echo(Print):
    """
    WScript.Echo() debugging function.
    """
    pass
