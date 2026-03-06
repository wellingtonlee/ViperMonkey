"""ViperMonkey: VBA Library — Communication and network functions."""
from ._common import *

class MultiByteToWideChar(VbaLibraryFunc):
    """
    MultiByteToWideChar() kernel32.dll function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 5)):
            return "NULL"

        # We have (hopefully) preprocessed this call so that the entire byte array
        # is passed as the 3rd parameter.
        data = params[2]
        if (not isinstance(data, list)):
            return "NULL"

        # Is the given string represented in wide chars?
        is_wide_char = True
        skip = False
        for b in data:
            skip = (not skip)
            if (skip):
                continue
            if (b != 0):
                is_wide_char = False
                break

        # Convert this to a string. If this is a ASCII string represented in wide
        # chars skip every 2nd byte (assume these are 0).
        r = ""
        skip = True
        for b in data:
            skip = (not skip)
            if (((not isinstance(b, int)) or (b > 255) or skip) and
                (is_wide_char)):
                continue
            r += chr(b)

        # Get the name of the variable where the result is stored.
        name = params[4]

        # Update the result variable with the converted string.
        context.set(name, r)
        return len(r)

    def num_args(self):
        return 5

class URLDownloadToFile(VbaLibraryFunc):
    """
    URLDownloadToFile() external function
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 3)):
            return
        context.report_action('Download URL', str(params[1]), 'External Function: urlmon.dll / URLDownloadToFile', strip_null_bytes=True)
        context.report_action('Write File', str(params[2]), 'External Function: urlmon.dll / URLDownloadToFile', strip_null_bytes=True)
        return 1

    def num_args(self):
        return 3

class URLDownloadToFileA(URLDownloadToFile):
    pass

class URLDownloadToFileW(URLDownloadToFile):
    pass

class URLDownloadToFile(VbaLibraryFunc):
    """
    URLDownloadToFile() external function.
    """

    def eval(self, context, params=None):
        if (params is None):
            return
        if (len(params) >= 3):
            context.report_action('Download URL', str(params[1]), 'External Function: urlmon.dll / URLDownloadToFile', strip_null_bytes=True)
            context.report_action('Write File', str(params[2]), 'External Function: urlmon.dll / URLDownloadToFile', strip_null_bytes=True)

class FollowHyperlink(VbaLibraryFunc):
    """
    FollowHyperlink() function.
    """

    def eval(self, context, params=None):
        if (params is None):
            return
        if (len(params) >= 1):
            context.report_action('Download URL', str(params[0]), 'FollowHyperLink', strip_null_bytes=True)

class Navigate(VbaLibraryFunc):
    """
    Navigate() function for loading a URL in a web browser.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        url = str(params[0])
        if (url.startswith("tp://")):
            url = "ht" + url
        context.report_action("GET", url, 'Load in browser', strip_null_bytes=True)

class LoadXML(VbaLibraryFunc):
    """
    LoadXML() MSXML2.DOMDocument.3.0 method.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return ""

        # Get the XML.
        xml = str(params[0]).strip()

        # Is this some base64?
        if (xml.startswith("<B64DECODE")):

            # Yes it is. Pull it out.
            start = xml.index(">") + 1
            end = xml.rindex("<")
            xml = xml[start:end].strip()

            # It looks like maybe this magically does base64 decode? Try that.
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("LoadXML(): Try base64 decode of '" + xml + "'...")
            decoded = utils.b64_decode(xml)
            if (decoded is not None):
                xml = decoded.replace(chr(0), "")
            else:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("LoadXML(): Base64 decode fail.")

        # Return the XML or base64 string.
        return xml

class InternetGetConnectedState(VbaLibraryFunc):
    """
    InternetGetConnectedState() function from wininet.dll.
    """

    def eval(self, context, params=None):

        # Always connected.
        return True

class InternetOpenA(VbaLibraryFunc):
    """
    InternetOpenA() function from wininet.dll.
    """

    def eval(self, context, params=None):

        # Always succeeds.
        return True

class CreateElement(VbaLibraryFunc):
    """
    Faked emulation of things like 'CreateObject("Microsoft.XMLDOM").createElement("tmp")'.
    """

    def eval(self, context, params=None):

        # Assume that this is something like 'CreateObject("Microsoft.XMLDOM").createElement("tmp")'.
        return "Microsoft.XMLDOM"

class SetTimeouts(VbaLibraryFunc):
    """
    ServerXMLHTTP SetTimeouts() method (stubbed).
    """

    def eval(self, context, params=None):
        pass

class SetOption(VbaLibraryFunc):
    """
    ServerXMLHTTP SetOption() method (stubbed).
    """

    def eval(self, context, params=None):
        pass

class SetRequestHeader(VbaLibraryFunc):
    """
    ServerXMLHTTP SetRequestHeader() method.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 2)):
            return

        # Field is 1st arg.
        field = utils.safe_str_convert(params[0])

        # Value is 2nd arg.
        val = utils.safe_str_convert(params[1])

        # Save the header value.
        info = "'" + field + "' ==> '" + val + "'"
        context.report_action('Set HTTP Header', info, "ServerXMLHTTP::SetRequestHeader()")
