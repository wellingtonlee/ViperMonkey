"""ViperMonkey: VBA Library — System, process execution, and environment functions."""
from ._common import *

class Sleep(VbaLibraryFunc):
    """
    Stubbed Sleep() function.
    """

    def eval(self, context, params=None):
        pass

    def num_args(self):
        return 1

class PrivateProfileString(VbaLibraryFunc):
    """
    PrivateProfileString method.
    """

    def eval(self, context, params=None):
        return "**MATCH ANY**"

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class Shell(VbaLibraryFunc):
    """
    6.1.2.8.1.15 Shell
    Function Shell(PathName As Variant, Optional WindowStyle As VbAppWinStyle = vbMinimizedFocus)
    As Double

    Runs an executable program and returns a Double representing the implementation-defined
    program's task ID if successful, otherwise it returns the data value 0.
    """

    def eval(self, context, params=None):

        # This might be the string "shell".
        if (params is None):
            return "shell"
        try:
            params.remove('ThisDocument')
            params.remove('BuiltInDocumentProperties')
        except ValueError:
            pass

        # Get the command to run.
        command = params[0]

        # Is the command invalid?
        if (not isinstance(command, str)):

            # No, Shell() will throw an error.
            msg = "Shell(" + str(command) + ") throws an error."
            context.set_error(msg)
            return 0

        # We have a valid shell command. Track it.
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Shell command type: " + str(type(command)))
        log.info('Shell(%r)' % command)
        context.report_action('Execute Command', command, 'Shell function', strip_null_bytes=True)
        return 0

    def num_args(self):
        return 1

class ExecuteStatement(Shell):
    pass

class ShellExecute(Shell):
    """
    shell.application.ShellExecute() function.
    """

    def eval(self, context, params=None):

        if ((params is None) or (len(params) < 2)):
            return 0
        command = str(params[0])
        args = str(params[1])
        log.info('ShellExecute(%r %r)' % (command, args))
        context.report_action('Execute Command', command + " " + args, 'Shell function', strip_null_bytes=True)
        return 0

class Eval(VbaLibraryFunc):
    """
    VBScript expression Eval() function.
    """

    def eval(self, context, params=None):

        # Pull out the expression to eval.
        if ((params is None) or (len(params) < 1)):
            return 0
        expr = utils.strip_nonvb_chars(str(params[0]))

        # Save original expression.
        orig_expr = expr

        # We are executing a string, so any "" in the string are really '"' when
        # we execute the string. Maybe?
        expr = expr.replace('""', '"')

        # Parse it. Assume this is an expression.
        r = None
        try:
            obj = expressions.expression.parseString(expr, parseAll=True)[0]

            # Evaluate the expression in the current context.
            # TODO: Does this actually get evalled in the current context?
            r = obj

        except ParseException:

            # Maybe replacing the '""' with '"' was a bad idea. Try the original
            # command.
            try:
                log.warning("Parsing failed on modified expression. Trying original expression ...")
                obj = expressions.expression.parseString(orig_expr, parseAll=True)[0]
                r = obj
            except ParseException:
                log.error("Parse error. Cannot evaluate '" + orig_expr + "'")
                return "NULL"

        # Do any final evaulation needed.
        if (isinstance(r, VBA_Object)):
            r = r.eval(context)
        return r

    def return_type(self):
        return "UNKNOWN"

class Execute(VbaLibraryFunc):
    """
    WScript Execute() function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or
            (len(params) == 0) or
            (isinstance(params[0], VBA_Object)) or
            (isinstance(params[0], VbaLibraryFunc))):
            return "NULL"

        # Save the command.
        command = utils.strip_nonvb_chars(str(params[0]))
        context.report_action('Execute Command', command, 'Execute() String', strip_null_bytes=True)
        command += "\n"

        # Fix invalid string assignments.
        command = strip_lines.fix_vba_code(command)

        # Save original command string.
        orig_command = command

        # We are executing a string, so any "" in the string are really '"' when
        # we execute the string.
        command = command.replace('""', '"')

        # Have we already parsed this?
        obj = None
        if (orig_command in parse_cache):
            obj = parse_cache[orig_command]

        # We have not parsed this previously.
        else:

            # Parse it.
            try:
                obj = modules.module.parseString(command, parseAll=True)[0]
            except ParseException:
                pass

            # Was is parsed?
            if obj is None:

                # Maybe replacing the '""' with '"' was a bad idea. Try the original
                # command.
                try:
                    log.warning("Parsing failed on modified command. Trying original command ...")
                    obj = modules.module.parseString(orig_command, parseAll=True)[0]
                except ParseException:
                    pass

            # Was is parsed?
            if obj is None:

                # Next attempt. Try cutting off the final line and executing.
                if ("\n" in orig_command.strip()):
                    short_command = orig_command.strip()[:orig_command.strip().rindex("\n")]
                    try:
                        log.warning("Parsing failed on original command. Trying shortened command minus last line ...")
                        obj = modules.module.parseString(short_command, parseAll=True)[0]
                    except ParseException:
                        pass

            # Was is parsed?
            if obj is None:

                # Try deleteing first non-alphabetic characters and reparsing.
                pos = 0
                ascii_pat = r"[A-Za-z]"
                while (pos < len(orig_command)):
                    if (re.match(ascii_pat, orig_command[pos])):
                        break
                    pos += 1
                short_command = orig_command[pos:].replace("\x1c", "\r").replace("\x1d", "\n")
                try:
                    log.warning("Parsing failed on original command. Trying shortened command up to first alphabetic character ...")
                    obj = modules.module.parseString(short_command, parseAll=True)[0]
                except ParseException:
                    pass

            # Cannot ever parse this. Punt.
            if obj is None:
                if (len(orig_command) > 50):
                    orig_command = orig_command[:50] + " ..."
                log.error("Parse error. Cannot evaluate '" + orig_command + "'")
                return "NULL"

        # Cache the parsed VB.
        parse_cache[orig_command] = obj

        # Are we execing this code inside JIT generated Python code?
        # Note that the dict of local variable values to update when we exec the
        # generated Python code is passed as the 2nd to last argument to Execute().
        if ((params[-1] == "__JIT_EXEC__") and
            (_eval_python(obj, context, add_boilerplate=True, namespace=params[-2]))):
            return "NULL"

        # No JIT. Do regular emulation.

        # Evaluate the expression in the current context.
        # TODO: Does this actually get evalled in the current context?
        r = obj
        if (isinstance(obj, VBA_Object)):

            # Load any new function definitions into the current context.
            obj.load_context(context)

            # Emulate the parsed code.
            r = obj.eval(context)

        # Add any functions declared in the execution to the global
        # context.
        return r

class ExecuteGlobal(Execute):
    """
    WScript ExecuteGlobal() function.
    """
    pass

class AddCode(Execute):
    """
    Visual Basic script control AddCode() method..
    """
    pass

class AddFromString(Execute):
    """
    Office programmatic macro editing method..
    """
    pass

class GetCursorPos(VbaLibraryFunc):
    """
    Faked GetCursorPos() function. Returns random location.
    """

    def eval(self, context, params=None):
        if ((var_names is None) or (len(var_names) == 0)):
            return 1

        # Set the given parameter to a random position.
        var_name = str(var_names[0])
        context.set(var_name + ".*", random.randint(100, 10000), force_global=True)

        return 0

class RtlMoveMemory(VbaLibraryFunc):
    """
    External RtlMoveMemory() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return

        # Report the memory move.
        context.report_action("External Call", "RtlMoveMemory(" + str(params) + ")", "RtlMoveMemory", strip_null_bytes=True)

        # Track the shellcode bytes.
        if (len(params) < 3):
            return
        from .. import vba_context
        vba_context.add_shellcode_data(params[0], params[1], params[2])

class GetByteCount_2(VbaLibraryFunc):
    """
    String encoder object method.
    """

    def eval(self, context, params=None):
        if ((len(params) == 0) or (not isinstance(params[0], str))):
            return 0
        return len(params[0])

class GetBytes_4(VbaLibraryFunc):
    """
    String encoder object method.
    """

    def eval(self, context, params=None):
        if ((len(params) == 0) or (not isinstance(params[0], str))):
            return []
        r = []
        for c in params[0]:
            r.append(ord(c))
        return r

class TransformFinalBlock(VbaLibraryFunc):
    """
    Base64 encoder object method.
    """

    def eval(self, context, params=None):
        if ((len(params) != 3) or (not isinstance(params[0], list))):
            return "NULL"

        # Pull out the byte values and start/end of the bytes to decode.
        vals = params[0]
        start = 0
        try:
            start = int(params[1])
        except (ValueError, TypeError):
            pass
        end = len(vals) - 1
        try:
            end = int(params[2])
        except (ValueError, TypeError):
            pass
        if (end > len(vals) - 1):
            end = len(vals) - 1
        if (start > end):
            start = end - 1

        # Reconstruct the base64 encoded string.
        base64_str = ""
        end += 1
        for b in vals[start : end]:
            base64_str += chr(b)

        # Decode the base64 encoded string.
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("TransformFinalBlock(): Try base64 decode of '" + base64_str + "'...")
        r = utils.b64_decode(base64_str)
        if (r is None):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("TransformFinalBlock(): Base64 decode fail.")
            r = "NULL"

        # Return the decoded string.
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Decoded string: " + r)
        return r

    def return_type(self):
        return "STRING"

class RunShell(VbaLibraryFunc):
    """
    Stubbed WScript.Shell Run() method.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return
        context.report_action('Execute Command', str(params[0]), 'WScript.Shell.Run()', strip_null_bytes=True)

class OnTime(VbaLibraryFunc):
    """
    Stubbed emulation of Application.OnTime(). Just immediately calls
    the callback function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 2)):
            return "NULL"

        # The name of the callback function should be the 2nd argument.
        callback_name = str(params[1])

        # Is this function defined?
        callback = None
        try:
            callback = context.get(callback_name)
        except KeyError:
            log.warning("OnTime() callback function '" + callback_name + "' not found.")
            return "NULL"
        from .. import procedures
        if (not isinstance(callback, procedures.Function) and
            not isinstance(callback, procedures.Sub)):
            log.warning("OnTime() callback function '" + callback_name + "' found, but not a function.")
            return "NULL"

        # Emulate the callback function.
        log.info("Running OnTime() callback function '" + callback_name + "'.")
        return eval_arg(callback, context=context)

class Environ(VbaLibraryFunc):
    """
    Environ() function for getting environment variable values.
    """

    def eval(self, context, params=None):

        # Common environment variables.
        env_vars = {}
        env_vars["ALLUSERSPROFILE".lower()] = 'C:\\ProgramData'
        env_vars["APPDATA".lower()] = 'C:\\Users\\admin\\AppData\\Roaming'
        env_vars["CommonProgramFiles".lower()] = 'C:\\Program Files\\Common Files'
        env_vars["CommonProgramFiles(x86)".lower()] = 'C:\\Program Files (x86)\\Common Files'
        env_vars["CommonProgramW6432".lower()] = 'C:\\Program Files\\Common Files'
        env_vars["COMPUTERNAME".lower()] = 'ADJH676F'
        env_vars["ComSpec".lower()] = 'C:\\WINDOWS\\system32\\cmd.exe'
        env_vars["DriverData".lower()] = 'C:\\Windows\\System32\\Drivers\\DriverData'
        env_vars["HOMEDRIVE".lower()] = 'C:'
        env_vars["HOMEPATH".lower()] = '\\Users\\admin'
        env_vars["LOCALAPPDATA".lower()] = 'C:\\Users\\admin\\AppData\\Local'
        env_vars["LOGONSERVER".lower()] = '\\\\HEROG76'
        env_vars["NUMBER_OF_PROCESSORS".lower()] = '4'
        env_vars["OneDrive".lower()] = 'C:\\Users\\admin\\OneDrive'
        env_vars["OS".lower()] = 'Windows_NT'
        env_vars["Path".lower()] = 'C:\\ProgramData\\Oracle\\Java\\javapath'
        env_vars["PATHEXT".lower()] = '.COM;.EXE;.BAT;.CMD;.VBS;.VBE;.JS;.JSE;.WSF;.WSH;.MSC'
        env_vars["PROCESSOR_ARCHITECTURE".lower()] = 'AMD64'
        env_vars["PROCESSOR_IDENTIFIER".lower()] = 'Intel64 Family 6 Model 158 Stepping 9, GenuineIntel'
        env_vars["PROCESSOR_LEVEL".lower()] = '6'
        env_vars["PROCESSOR_REVISION".lower()] = '9e09'
        env_vars["ProgramData".lower()] = 'C:\\ProgramData'
        env_vars["ProgramFiles".lower()] = 'C:\\Program Files'
        env_vars["ProgramFiles(x86)".lower()] = 'C:\\Program Files (x86)'
        env_vars["ProgramW6432".lower()] = 'C:\\Program Files'
        env_vars["PROMPT".lower()] = '$P$G'
        env_vars["PSModulePath".lower()] = 'C:\\Program Files\\WindowsPowerShell\\Modules;C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\Modules;C:\\Program Files\\Microsoft Message Analyzer\\PowerShell\\'
        env_vars["PUBLIC".lower()] = 'C:\\Users\\Public'
        env_vars["SESSIONNAME".lower()] = 'Console'
        env_vars["SystemDrive".lower()] = 'C:'
        env_vars["SystemRoot".lower()] = 'C:\\WINDOWS'
        env_vars["TEMP".lower()] = 'C:\\Users\\admin\\AppData\\Local\\Temp'
        env_vars["TMP".lower()] = 'C:\\Users\\admin\\AppData\\Local\\Temp'
        env_vars["USERDNSDOMAIN".lower()] = 'REMOTE.FOURTHWALL.COM'
        env_vars["USERDOMAIN".lower()] = 'FOURTHWALL'
        env_vars["USERDOMAIN_ROAMINGPROFILE".lower()] = 'FOURTHWALL'
        env_vars["USERNAME".lower()] = 'admin'
        env_vars["USERPROFILE".lower()] = 'C:\\Users\\admin'
        env_vars["VS110COMNTOOLS".lower()] = 'C:\\Program Files (x86)\\Microsoft Visual Studio 11.0\\Common7\\Tools\\'
        env_vars["VS120COMNTOOLS".lower()] = 'C:\\Program Files (x86)\\Microsoft Visual Studio 12.0\\Common7\\Tools\\'
        env_vars["VS140COMNTOOLS".lower()] = 'C:\\Program Files (x86)\\Microsoft Visual Studio 14.0\\Common7\\Tools\\'
        env_vars["VSSDK140Install".lower()] = 'C:\\Program Files (x86)\\Microsoft Visual Studio 14.0\\VSSDK\\'
        env_vars["windir".lower()] = 'C:\\WINDOWS'

        # Get the environment variable name.
        var_name = utils.safe_str_convert(params[0]).strip('%')

        # Is this an environment variable we know?
        if context.expand_env_vars and var_name.lower() in env_vars:
            r = env_vars[var_name.lower()]
        else:
            r = "%{}%".format(var_name.upper())

        # Done.
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Environ: %r returns %r" % (self, r))
        return r

    def return_type(self):
        return "STRING"

class ExpandEnvironmentStrings(Environ):
    pass

class Run(VbaLibraryFunc):
    """
    Application.Run() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return 0

        # Get the name of the function to call.
        func_name = str(params[0])

        # Strip the name of the function down if needed.
        if ("." in func_name):
            func_name = func_name[func_name.rindex(".") + 1:]

        # Get any parameters to pass to the function to call.
        call_params = None
        if (len(params) > 1):
            call_params = params[1:]

        # Can we find the function to call?
        try:
            context.report_action("Run", func_name, 'Interesting Function Call', strip_null_bytes=True)
            s = context.get(func_name)
            return s.eval(context=context, params=call_params)
        except KeyError:
            log.warning("Application.Run() failed. Cannot find function " + str(func_name) + ".")
            return 0

class Exec(VbaLibraryFunc):
    """
    Application.Exec() function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return 1

        # Get the command to run.
        cmd = str(params[0])
        context.report_action("Execute Command", cmd, 'Shell function', strip_null_bytes=True)

        # Say it was successful.
        return 0

class ExecQuery(VbaLibraryFunc):
    """
    Application.ExecQuery() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"

        # Get the query to run.
        cmd = str(params[0])
        context.report_action("Execute Query", cmd, 'Query', strip_null_bytes=True)

        # Return some data for some queries.
        if (cmd.lower() == "select * from win32_process"):
            return [{"name" : "wscript.exe"},
                    {"name" : "cscript.exe"},
                    {"name" : "word.exe"},
                    {"name" : "excel.exe"},]

        # Say it was successful.
        return ["", ""]

class WinExec(VbaLibraryFunc):
    """
    WinExec() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"

        cmd = params[0]
        context.report_action("Run", cmd, 'Interesting Command Execution', strip_null_bytes=True)
        return ''

class GetTickCount(VbaLibraryFunc):
    """
    GetTickCount() function. Randomly increments the tick count.
    """

    def eval(self, context, params=None):
        global ticks
        ticks += random.randint(100, 10000)
        return ticks

class Send(VbaLibraryFunc):
    """
    Faked emulation of HTTP send(). Always returns 200.
    """

    def eval(self, context, params=None):
        return 200

class WriteProcessMemory(VbaLibraryFunc):
    """
    WriteProcessMemory() external method.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        context.report_action('Write Process Memory', str(params), 'External Function: kernel32.dll / WriteProcessMemory', strip_null_bytes=True)

        # Track the shellcode bytes.
        if (len(params) < 4):
            return
        from .. import vba_context
        vba_context.add_shellcode_data(params[1], params[2], params[3])
