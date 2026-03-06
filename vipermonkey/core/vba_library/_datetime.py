"""ViperMonkey: VBA Library — Date and time functions."""
from ._common import *

class MonthName(VbaLibraryFunc):
    """
    MonthName() function. Currently only returns results in Italian.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        num = params[0]
        if ((not isinstance(num, int)) or (num > 12) or (num < 1)):
            return "NULL"
        # TODO: Somehow specify the language for the months.
        months = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
        return months[num-1]

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class WeekDay(VbaLibraryFunc):
    """
    VBA WeekDay function
    """

    def eval(self, context, params=None):

        # Get date string.
        if ((params is None) or (len(params) == 0)):
            return 1
        date_str = str(params[0]).replace("#", "")
        date_obj = None

        # TODO: Handle more and more date formats.

        # 4/20/1889
        if (date_str.count("/") == 2):
            try:
                date_obj = datetime.strptime(date_str, '%m/%d/%Y')
            except (ValueError, TypeError):
                pass

        if (date_obj is not None):
            r = date_obj.weekday()
            # Looks like VBA week day is off by 2 from Python week day.
            r += 2
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("WeekDay(%r): return %r" % (date_str, r))
            return r
        return 1

    def num_args(self):
        return 1

class Format(VbaLibraryFunc):
    """
    VBA Format function
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Are we faking a value for this particular format call?
        r = params[0]
        if (len(params) > 1):
            typ = str(params[1])

            # Fake up a date if needed.
            # TODO: Currently this fake date is specific to a campaign targeting Italy.
            if (typ.lower() == "long date"):
                r = "gioved\xc3\xac 27 giugno 2019"

            # Let's match any currency checks.
            if (typ.lower() == "currency"):
                r = "**MATCH ANY**"

        # Done.
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Format(%r): return %r" % (self, r))
        return r

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class Hour(VbaLibraryFunc):
    """
    Hour() time function (stubbed).
    """

    def eval(self, context, params=None):
        return 13

class Day(VbaLibraryFunc):
    """
    Day() function. This is currently partially implemented.
    """

    def eval(self, context, params=None):
        # This is usually used for gating, so have it match anything.
        return "**MATCH ANY**"
    """
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        txt = params[0]
        if ((txt is None) or (txt == "NULL")):
            txt = ''
        r = str(txt)

        # It looks like this should pull the day out of a date string. See if we can
        # handle a simple date string.
        f = r.split("/")
        if (len(f) == 3):
            try:
                r = int(f[1])
            except (ValueError, TypeError):
                pass

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Day: %r returns %r" % (self, r))
        return r
    """

class Month(VbaLibraryFunc):
    """
    Excel Month() function. Currently stubbed.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        try:
            arg = int(params[0])
            if (arg == 1):
                return 12
            if (arg < 33):
                return 1
            if (arg < 61):
                return 2
            if (arg < 92):
                return 3
            if (arg < 101):
                return 4

            # TODO: Handle other values.
            return 1

        except (ValueError, TypeError):
            pass

        return 1

class DatePart(VbaLibraryFunc):
    """
    DatePart() function. Currently (very) stubbed to just return 3.
    """

    def eval(self, context, params=None):
        return 3

class Date(VbaLibraryFunc):
    """
    Date() function. Currently stubbed to just return the current date as
    a Python datetime object.
    """

    def eval(self, context, params=None):
        return date.today()

class DateAdd(VbaLibraryFunc):
    """
    DateAdd() function. Currently stubbed to just return the current date as
    a Python datetime object.
    """

    def eval(self, context, params=None):
        return date.today()

class Year(VbaLibraryFunc):
    """
    Year() function. Currently stubbed.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        t = params[0]
        r = "**MATCH ANY**"
        if ((isinstance(t, datetime)) or (isinstance(t, date))):
            r = int(t.year)
        return r

class Minute(VbaLibraryFunc):
    """
    Minute() function. Currently stubbed.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        t = params[0]
        r = 0
        if (isinstance(t, datetime)):
            r = int(t.minute)
        return r

class Second(VbaLibraryFunc):
    """
    Second() function. Currently stubbed.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        t = params[0]
        r = 0
        if (isinstance(t, datetime)):
            r = int(t.second)
        try:
            d = datetime.strptime(t, '%H:%M:%S')
            r = int(d.second)
        except (ValueError, TypeError):
            pass
        return r

class Timer(VbaLibraryFunc):
    """
    Timer() method (stubbed).
    """

    def eval(self, context, params=None):
        return int(time.mktime(datetime.now().timetuple()))

class DateDiff(VbaLibraryFunc):
    """
    Stubbed DateDiff() function.
    """

    def eval(self, context, params=None):
        return 15904387438 + (5000 - random.randint(100, 10000))
