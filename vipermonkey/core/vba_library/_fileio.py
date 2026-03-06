"""ViperMonkey: VBA Library — File I/O and filesystem functions."""
from ._common import *

class Kill(VbaLibraryFunc):
    """
    Kill statement.
    """

    def eval(self, context, params=None):
        if ((params is not None) and (len(params) > 0)):
            context.report_action('Delete File', params[0], 'Kill', strip_null_bytes=True)
        return ""

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class RmDir(VbaLibraryFunc):
    """
    RmDir statement.
    """

    def eval(self, context, params=None):
        if ((params is not None) and (len(params) > 0)):
            context.report_action('Delete Directory', params[0], 'RmDir', strip_null_bytes=True)
        return ""  # vbOK

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class CreateFolder(VbaLibraryFunc):
    """
    CreatFolder() method.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        folder = str(params[0])
        context.report_action('Create Folder', folder, 'CreateFolder()', strip_null_bytes=True)
        return 0

    def num_args(self):
        return 1

class BuildPath(VbaLibraryFunc):
    """
    BuildPath() method
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 2)):
            return "NULL"
        return str(params[0]) + str(params[1])

class GetSpecialFolder(VbaLibraryFunc):
    """
    GetSpecialFolder() function
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "UNKNOWN_FOLDER\\"
        try:
            typ = int(params[0])
            if (typ == 0):
                return "C:\\Windows\\"
            elif (typ == 1):
                return "C:\\Windows\\system32\\"
            elif (typ == 2):
                return "C:\\Documents and Settings\\admin\\Local Settings\\Temp\\"
            else:
                return "UNKNOWN_FOLDER\\"
        except (ValueError, TypeError):
            return "UNKNOWN_FOLDER\\"

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class GetFolder(VbaLibraryFunc):
    """
    GetFolder() function
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "UNKNOWN_FOLDER\\"
        context.report_action('Get Folder', "GetFolder(" + str(params) + ")", '---', strip_null_bytes=True)
        return params[0]

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class MakeSureDirectoryPathExists(VbaLibraryFunc):
    """
    MakeSureDirectoryPathExists() VB function (stubbed).
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return 1
        context.report_action("Create Folder", params[0], 'Interesting Function Call', strip_null_bytes=True)
        return 1

    def num_args(self):
        return 1

class FolderExists(VbaLibraryFunc):
    """
    FolderExists() VB function (stubbed).
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return False

        # Is this a directory that is expected to exist?
        expected_dirs = set(["c:\\users", "c:\\programdata"])
        curr_dir = str(params[0]).lower()
        return ((curr_dir in expected_dirs) or (curr_dir[:-1] in expected_dirs))

    def num_args(self):
        return 1

class GetFile(VbaLibraryFunc):
    """
    GetFile() VB method (stubbed).
    """

    def eval(self, context, params=None):
        if (params is None):
            return
        context.report_action('Get File', "GetFile(" + str(params) + ")", '---', strip_null_bytes=True)

    def num_args(self):
        return 1

class FileLen(VbaLibraryFunc):
    """
    FileLen() VB function (stubbed). Always returns -1.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return -1
        context.report_action('Check File Length', "FileLen(" + str(params) + ")", '---', strip_null_bytes=True)
        return -1

    def num_args(self):
        return 1

class FileCopy(VbaLibraryFunc):
    """
    FileCopy() VB function (stubbed).
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 2)):
            return
        context.report_action('Copy File', "FileCopy(" + str(params) + ")", '---', strip_null_bytes=True)

    def num_args(self):
        return 2

class CopyFile(FileCopy):
    pass

class CopyHere(VbaLibraryFunc):
    """
    CopyHere() VB function (stubbed).
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return
        context.report_action('Copy File', "CopyHere(" + str(params) + ")", '---', strip_null_bytes=True)

    def num_args(self):
        return 1

class FileExists(VbaLibraryFunc):
    """
    FileExists() VB function (stubbed).
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return False
        fname = str(params[0])
        if ("powershell" in fname.lower()):
            return True
        if ("cmd.exe" in fname.lower()):
            return True
        if ("explorer.exe" in fname.lower()):
            return True
        if ("c:\\programdata" in fname.lower()):
            return True
        return False

    def num_args(self):
        return 1

class ChDir(VbaLibraryFunc):
    """
    ChDir() function.
    """

    def eval(self, context, params=None):
        if ((params is not None) and (len(params) > 0)):
            context.report_action('Change Directory', params[0], 'ChDir', strip_null_bytes=True)
        return ""  # vbOK

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class Dir(VbaLibraryFunc):
    """
    Dir() file/directory finding function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return ""
        pat = str(params[0])
        attrib = None
        # TODO: Handle multiple attributes.
        if (len(params) > 1):
            attrib = params[1]

        # Handle a special case for a maldoc that looks for things
        # not existing in a certain directory.
        if (("\\Microsoft\\Corporation\\" in pat) or
            ("\\AppData\\Roaming\\Microsoft" in pat) or
            ("\\AppData\\Local\\Temp" in pat)):
            return ""

        # Just act like we found something always.
        r = pat.replace("*", "foo")

        # TODO: Figure out how to simulate actual file searches.
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Dir: %r returns %r" % (self, r))
        return r

    def return_type(self):
        return "STRING"

class CurDir(VbaLibraryFunc):
    """
    CurDir() function.
    """

    def eval(self, context, params=None):
        return "~"

    def return_type(self):
        return "STRING"

class Close(VbaLibraryFunc):
    """
    File Close statement.
    """

    def eval(self, context, params=None):

        # Are we closing a file pointer?
        file_id = None
        if ((params is not None) and
            (len(params) == 1) and
            (params[0] is not None) and
            (isinstance(params[0], str)) and
            (params[0].startswith('#'))):

            # Get the ID of the file being closed.
            try:
                file_id = context.get(params[0])
            except KeyError:
                file_id = str(params[0])

        # Close() object method call?
        else:

            # TODO: Currently the object on which Close() is being called is not
            # being tracked. We will only handle the Close() if there is only 1
            # current open file.
            if not context.open_files:
                log.warning("Cannot process Close(). No open files.")
                return

            if len(context.open_files) > 1:
                log.warning("More than 1 file is open. Closing an arbitrary file.")
                file_id = context.get_interesting_fileid()
            else:
                # Get the ID of the file.
                file_id = context.open_files.keys()[0]

        # We are actually closing a file.
        context.close_file(file_id)


class Put(VbaLibraryFunc):
    """
    File Put statement.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 2)):
            return "NULL"

        # Get the ID of the file.
        file_id = params[0]

        # TODO: Handle writing at a given file position.

        # Get the data.
        data = params[1]
        if (len(params) == 3):
            data = params[2]

        # Has the file been opened?
        if (not context.file_is_open(file_id)):
            context.open_file(file_id)

        context.write_file(file_id, data)

class WriteByte(VbaLibraryFunc):
    """
    MemoryStream WriteByte() method.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return
        context.report_action('Write Process Memory', str(params), 'MemoryStream.WriteByte', strip_null_bytes=True)

class WriteLine(VbaLibraryFunc):
    """
    File WriteLine() method.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"

        # Get the data.
        data = params[0]
        if (len(params) == 3):
            data = params[2]

        # Save writes that look like they are writing URLs.
        data_str = utils.safe_str_convert(data)
        if (("http:" in data_str) or ("https:" in data_str)):
            context.report_action('Write URL', data_str, 'File Write')

        # TODO: Currently the object on which WriteLine() is being called is not
        # being tracked. We will only handle the WriteLine() if there is only 1
        # current open file.
        if ((context.open_files is None) or (len(context.open_files) == 0)):
            log.error("Cannot process WriteLine(). No open files.")
            return
        file_id = None
        if (len(context.open_files) > 1):
            log.warning("More than 1 file is open. Writing to an arbitrary file.")
            file_id = context.get_interesting_fileid()
            log.warning("Writing to '" + str(file_id) + "' .")
        else:

            # Get the ID of the file.
            file_id = context.open_files.keys()[0]

        # TODO: Handle writing at a given file position.

        context.write_file(file_id, data)
        context.write_file(file_id, b'\n')

class WriteText(VbaLibraryFunc):
    """
    File WriteText() method.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"

        # Get the data.
        txt = params[0]
        if (len(params) == 3):
            txt = params[2]

        # Set the text value of the string as a faux variable. Make this
        # global as a hacky solution to handle fields in user defined objects.
        #
        # We are appending the written data to whatever is already there.

        # Assume we are writing to ADODB.Stream.ReadText
        var_name = "ADODB.Stream.ReadText"
        if (not context.contains(var_name)):
            context.set(var_name, "", force_global=True)
        final_txt = context.get(var_name) + txt
        context.set(var_name, final_txt, force_global=True)

class ReadText(VbaLibraryFunc):
    """
    ReadText() stream method (stubbed).
    """

    def eval(self, context, params=None):

        # Doing base64 conversion with a VBA object?
        with_str = str(context.with_prefix).strip()
        if (with_str.endswith("GetDecodedContentStream")):
            var_name = with_str.replace("GetDecodedContentStream", "GetEncodedContentStream") + ".ReadText"
            if (context.contains(var_name)):
                return context.get(var_name)

        # TODO: Currently the stream object on which ReadText() is
        # being called is not being tracked. We will only handle the
        # ReadText() if there is only 1 current open file.
        if not context.open_files:
            log.error("Cannot process ReadText(). No open streams.")
            return
        if len(context.open_files) > 1:
            log.error("Cannot process ReadText(). Too many open streams.")
            return

        # Simulate the read.

        # Get the ID of the file.
        file_id = context.open_files.keys()[0]

        # TODO: This function takes a parameter that specifies the number of bytes to read!!

        # Get the data to read.
        raw_data = context.open_files[file_id]

        # Return the data.
        return raw_data

    def return_type(self):
        return "STRING"

class EOF(VbaLibraryFunc):
    """
    Stubbed EOF file method.
    """

    def eval(self, context, params=None):
        return True

    def num_args(self):
        return 1

class CreateTextFile(VbaLibraryFunc):
    """
    CreateTextFile() method.
    """

    def eval(self, context, params=None):
        if not params:
            return "NULL"

        # Get the name of the file being opened.
        try:
            fname = context.get(params[0])
        except KeyError:
            fname = str(params[0])

        # Do we have a numeric file ID?
        file_id = ""
        if (len(params) > 1):
            file_id = params[1]

        # Save that the file is opened.
        context.open_file(fname, file_id)
        context.report_action('File Access', fname, "")

        # This could be an external WebDAV access.
        if (fname.startswith("\\\\")):

            # Pull out the mapped drive ID.
            if ("\\" in fname[2:]):
                end = fname[2:].index("\\") + 2
                drive_id = fname[2:end].strip()
                if (re.search(r"[\w_]{1,100}\.\w{2,10}", drive_id) is not None):
                    context.save_intermediate_iocs("http://" + drive_id)

        # How about returning the name of the opened file.
        return fname

class Open(CreateTextFile):
    """
    Open() file function. Also Open() HTTP function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Is this a HTTP GET?
        if ((len(params) >= 2) and
            ((str(params[0]).strip() == "GET") or
             (str(params[1]).startswith("ftp://")) or
             (str(params[1]).startswith("http://")) or
             (str(params[1]).startswith("https://")))):
            url = str(params[1])
            if (url.startswith("tp://")):
                url = "ht" + url
            context.report_action("GET", url, 'Interesting Function Call', strip_null_bytes=True)

        # It is a regular file open.
        else:
            super(Open, self).eval(context, params)

class OpenTextFile(CreateTextFile):
    """
    OpenTextFile() file function.
    """
    pass

class DeleteFile(VbaLibraryFunc):
    """
    File delete DeleteFile() call.
    """

    def eval(self, context, params=None):
        if (params is None):
            return
        if (len(params) > 1):
            context.report_action('Delete File', str(params[1]), 'DeleteFile() Call', strip_null_bytes=True)
        if (len(params) == 1):
            context.report_action('Delete File', str(params[0]), 'DeleteFile() Call', strip_null_bytes=True)

class MoveFile(VbaLibraryFunc):
    """
    File move MoveFile() call.
    """

    def eval(self, context, params=None):
        if (params is None):
            return
        if (len(params) > 1):
            context.report_action('Move File', "MoveFile(" + str(params[0]) + ", " + str(params[1]) + ")",
                                  'MoveFile() Call', strip_null_bytes=True)

class GetExtensionName(VbaLibraryFunc):

    def eval(self, context, params=None):
        if (params is None):
            return
        r = ""
        if (len(params) >= 1):
            fname = str(params[0])
            if ("." in fname):
                r = fname[fname.rindex("."):]
        return r

class GetParentFolderName(VbaLibraryFunc):
    """
    GetParentFolderName() method.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Pull the parent directory.
        curr_dir = str(params[0])
        if ("\\" in curr_dir):
            r = curr_dir[:curr_dir.rindex("\\")+1]
        else:
            r = "C:\\"
        return r

    def num_args(self):
        return 1

class DriveExists(VbaLibraryFunc):
    """
    DriveExists() function for checking to see if a drive exists.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        drive = str(params[0]).lower()
        r = False
        # Assume the C: drive is always there.
        if ((drive == 'c') or (drive == 'c:')):
            r = True
        return r

class SaveToFile(VbaLibraryFunc):
    """
    SaveToFile() ADODB.Stream method.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return ""

        # Just return the file name. This is used in
        # expressions.MemberAccessExpression._handle_savetofile().
        r = str(params[0])
        context.last_saved_file = r
        return r

    def return_type(self):
        return "STRING"

class SaveAs(VbaLibraryFunc):
    """
    ActiveDocument.SaveAs() method.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 2)):
            return 0

        # Pull out the name of the file to save to and the format
        # for saving.
        new_fname = str(params[0])
        fmt = params[1]
        for param in params:
            if (isinstance(param, expressions.NamedArgument)):
                if (param.name == "FileName"):
                    new_fname = param.value
                if (param.name == "FileFormat"):
                    fmt = param.value

        # Save the current doc to a file.

        # Handle saving as text.
        # wdFormatText = 2
        if (fmt != 2):
            return 0

        # Get the doc text.
        doc_text = None
        try:
            paragraphs = context.get("ActiveDocument.Paragraphs".lower())
            doc_text = ""
            for p in paragraphs:
                doc_text += p + "\n"
        except KeyError:
            return 0

        # Open the saveas file.
        opener = CreateTextFile()
        opener.eval(context, [new_fname])

        # Write the data.
        writer = WriteLine()
        writer.eval(context, [doc_text])

        # Close the file.
        closer = Close()
        closer.eval(context, [])

        # Done.
        return 1

class SaveAs2(SaveAs):
    pass

class FreeFile(VbaLibraryFunc):
    """
    FreeFile() function.
    """

    def eval(self, context, params=None):

        # Return index of next open file.
        v = len(context.open_files) + 1
        return v

class Unprotect(VbaLibraryFunc):
    """
    Stubbed Unprotect() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return
        passwd = str(params[0])
        context.report_action('Unprotect()', passwd, 'Try Sheet Unprotect Password', strip_null_bytes=True)

class Write(VbaLibraryFunc):
    """
    Write() method.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"

        # Get the data.
        data = str(params[0])

        # Save writes that look like they are writing URLs.
        if (("http:" in data) or ("https:" in data)):
            context.report_action('Write URL', data, 'File Write', strip_null_bytes=True)

        # TODO: Currently the object on which Write() is being called is not
        # being tracked. We will only handle the Write() if there is only 1
        # obvious open file.
        if not context.open_files:
            log.error("Cannot process Write(). No open files.")
            return
        files = context.open_files.keys()
        if len(files) > 1:
            # Skip ADODB.Stream when guessing what file to write to.
            tmp_files = []
            for f in files:
                if (f.strip() == "ADODB.Stream"):
                    continue
                tmp_files.append(f)
            files = tmp_files
            if len(files) > 1:
                log.error("Cannot process Write(). Too many open files.")
                return

        # Simulate the write.

        # Get the ID of the file.
        file_id = files[0]
        log.info("Writing data to " + str(file_id) + " .")

        context.write_file(file_id, data)
