"""ViperMonkey: VBA Library — Math, numeric, and type conversion functions."""
from ._common import *

class Int(VbaLibraryFunc):
    """
    Int() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        # TODO: Actually implement this properly.
        val = params[0]
        try:
            if (isinstance(val, str) and (val.lower().startswith("&h"))):
                val = "0x" + val[2:]
                r = int(val, 16)
            elif (isinstance(val, str) and
                (val.lower().startswith("i")) and
                (len(val) == 3)):
                val = "0x" + val[1:]
                r = int(val, 16)
            elif (isinstance(val, str) and (("e" in val) or ("E" in val))):
                r = int(decimal.Decimal(val))
            else:
                r = utils.int_convert(val)
            # -32,768 to 32,767
            if ((r > 32767) or (r < -32768)):
                # Overflow. Assume On Error Resume Next.
                r = "NULL"
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Int: return %r" % r)
            return r
        except Exception as e:
            log.error("Int(): Invalid call int(%r) [%s]. Returning ''." % (val, str(e)))
            return ''

class CInt(Int):
    """
    Same as Int() for our purposes.
    """
    pass

class Oct(VbaLibraryFunc):
    """
    Oct() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return "NULL"
        val = params[0]
        try:
            r = oct(val)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Oct: return %r" % r)
            return r
        except (TypeError, ValueError):
            log.error("Oct(): Invalid call oct(%r). Returning ''." % val)
            return ''

class Sgn(VbaLibraryFunc):
    """
    Sgn() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        num = params[0]
        r = ''
        try:
            n = utils.int_convert(num)
            if n == 0:
                r = 0
            else:
                r = int(math.copysign(1, n))
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Sgn: %r returns %r" % (self, r))
        return r

class Sqr(VbaLibraryFunc):
    """
    Sqr() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            num = utils.int_convert(params[0]) + 0.0
            r = math.sqrt(num)
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Sqr: %r returns %r" % (self, r))
        return r

class Abs(VbaLibraryFunc):
    """
    Abs() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            num = utils.int_convert(params[0])
            r = abs(num)
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Abs: %r returns %r" % (self, r))
        return r

class Fix(VbaLibraryFunc):
    """
    Fix() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            num = float(params[0])
            r = math.floor(num)
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Fix: %r returns %r" % (self, r))
        return r

class Round(VbaLibraryFunc):
    """
    Round() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            num = float(params[0])
            sig = 0
            if (len(params) == 2):
                sig = utils.int_convert(params(1))
            r = round(num, sig)
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Round: %r returns %r" % (self, r))
        return r

class Hex(VbaLibraryFunc):
    """
    Hex() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            num = utils.int_convert(params[0])
            # Number treated as an unsigned int by VBA.
            if (num < 0):
                num += (1 << 32)
            r = hex(num).replace("0x","").upper()
            # VBA chops FFs from the start of the string down to 1 FF.
            if (r.startswith("FF")):
                r = "FF" + r[r.rindex("FF") + len("FF"):]
                if ((len(r) % 2) != 0):
                    r = "F" + r
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Hex: %r returns %r" % (self, r))
        return r

class CByte(VbaLibraryFunc):
    """
    CByte() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            tmp = str(params[0]).upper()
            if (tmp.lower().startswith("&h")):
                tmp = tmp.lower().replace("&h", "0x")
                tmp = int(tmp, 16)
            num = int(round(float(tmp)))
            r = num
            if (r > 255):
                r = 255
        except Exception as e:
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("CByte: %r returns %r" % (self, r))
        return r

class CLng(VbaLibraryFunc):
    """
    CLng() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"

        # Handle abstracted pointers to memory.
        val = params[0]
        if (isinstance(val, str) and
            (not val.lower().startswith("&h")) and
            (val.startswith("&"))):
            return val

        # Actually try to convert to a number.
        r = ''
        try:
            tmp = val
            if (isinstance(tmp, str)):
                tmp = val.upper()
                if (tmp.lower().startswith("&h")):
                    tmp = tmp.lower().replace("&h", "0x")
                    tmp = int(tmp, 16)
                elif (len(tmp) == 1):
                    tmp = ord(tmp)
            r = round(tmp)
            if ((r > 2147483647) or (r < -2147483647)):
                # Overflow. Assume On Error Resume Next.
                r = "NULL"
        except (ValueError, TypeError, AttributeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("CLng: %r returns %r" % (self, r))
        return r

class CBool(VbaLibraryFunc):
    """
    CBool() type conversion function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        val = params[0]
        r = 0
        if val is True or val == 1:
            r = 1
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("CBool: %r returns %r" % (self, r))
        return r

class CDate(VbaLibraryFunc):
    """
    CDate() type conversion function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        # TODO: For now this is stubbed out. Handling dates correctly is hard.
        r = 12345
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("CDate: %r returns %r" % (self, r))
        return r

class CStr(VbaLibraryFunc):
    """
    CStr() type conversion function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        val = params[0]
        r = str(val)
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("CStr: %r returns %r" % (self, r))
        return r

    def return_type(self):
        return "STRING"

class CSng(VbaLibraryFunc):
    """
    CSng() type conversion function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            tmp = params[0].upper()
            if (tmp.lower().startswith("&h")):
                tmp = tmp.lower().replace("&h", "0x")
                tmp = int(tmp, 16)
            r = float(tmp)
        except (ValueError, TypeError, AttributeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("CSng: CSng(%r) returns %r" % (params[0], r))
        return r

class CVar(VbaLibraryFunc):
    """
    CVar() type conversion function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"

        # We are not tracking variant types, so work as a pass-through.
        return params[0]

class Atn(VbaLibraryFunc):
    """
    Atn() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            num = float(params[0])
            r = math.atan(num)
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Atn: %r returns %r" % (self, r))
        return r

class Tan(VbaLibraryFunc):
    """
    Tan() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            num = float(params[0])
            r = math.tan(num)
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Tan: %r returns %r" % (self, r))
        return r

class Cos(VbaLibraryFunc):
    """
    Cos() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            num = float(params[0])
            r = math.cos(num)
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Cos: %r returns %r" % (self, r))
        return r

class Log(VbaLibraryFunc):
    """
    Log() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = 0.0
        try:
            num = float(params[0])
            r = math.log(num)
        except ValueError as e:
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.error("Log(" + str(params[0]) + ") failed. " + str(e))
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Log: %r returns %r" % (self, r))
        return r

class Exp(VbaLibraryFunc):
    """
    Exp() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = params[0]
        try:
            num = float(params[0])
            r = math.exp(num)
        except Exception as e:
            log.error("Exp(" + str(params[0]) + ") failed. " + str(e))
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Exp: %r returns %r" % (self, r))
        return r

class Sin(VbaLibraryFunc):
    """
    Sin() math function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        r = ''
        try:
            num = float(params[0])
            r = math.sin(num)
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Sin: %r returns %r" % (self, r))
        return r

class RGB(VbaLibraryFunc):
    """
    RGB() color function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 3)):
            return "NULL"
        r = ''
        try:
            red = utils.int_convert(params[0])
            green = utils.int_convert(params[1])
            blue = utils.int_convert(params[2])
            r = red + (green * 256) + (blue * 65536)
        except (ValueError, TypeError):
            pass
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("RGB: %r returns %r" % (self, r))
        return r

class QBColor(VbaLibraryFunc):
    """
    QBColor() color lookup function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return 0
        val = int(params[0])
        if ((val < 0) or (val > 15)):
            return 0
        lookup = {
            0 : 0,
            1 : 8388608,
            2 : 32768,
            3 : 8421376,
            4 : 128,
            5 : 8388736,
            6 : 32896,
            7 : 12632256,
            8 : 8421504,
            9 : 16711680,
            10 : 65280,
            11 : 16776960,
            12 : 255,
            13 : 16711935,
            14 : 65535,
            15 : 16777215
        }
        return lookup[val]

    def num_args(self):
        return 1

class Pmt(VbaLibraryFunc):
    """
    Pmt() payment computation function.

    Returns a Double specifying the payment for an annuity based on
    periodic, fixed payments and a fixed interest rate.

    Pmt(rate, nper, pv [, fv [, type ]] )

    The Pmt function has these named arguments:

    rate Required. Double specifying interest rate per period. For
    example, if you get a car loan at an annual percentage rate (APR)
    of 10 percent and make monthly payments, the rate per period is
    0.1/12, or 0.0083.

    nper Required. Integer specifying total number of payment periods
    in the annuity. For example, if you make monthly payments on a
    four-year car loan, your loan has a total of 4 * 12 (or 48)
    payment periods.

    pv Required. Double specifying present value (or lump sum) that a
    series of payments to be paid in the future is worth now. For
    example, when you borrow money to buy a car, the loan amount is
    the present value to the lender of the monthly car payments you
    will make.

    fv Optional. Variant specifying future value or cash balance you
    want after you've made the final payment. For example, the future
    value of a loan is $0 because that's its value after the final
    payment. However, if you want to save $50,000 over 18 years for
    your child's education, then $50,000 is the future value. If
    omitted, 0 is assumed.

    type Optional. Variant specifying when payments are due. Use 0 if
    payments are due at the end of the payment period, or use 1 if
    payments are due at the beginning of the period. If omitted, 0 is
    assumed.

    '               This function, together with the four following
    '               it (Pv, Fv, NPer and Rate), can calculate
    '               a certain value associated with a regular series of
    '               equal-sized payments.  This series can be fully described
    '               by these values:
    '                     Pv   - present value
    '                     Fv   - future value (at end of series)
    '                     PMT  - the regular payment
    '                     nPer - the number of 'periods' over which the
    '                            money is paid
    '                     Rate - the interest rate per period.
    '                            (type - payments at beginning (1) or end (0) of
    '                            the period).
    '               Each function can determine one of the values, given the others.
    '
    '               General Function for the above values:
    '
    '                                                      (1+rate)^nper - 1
    '               pv * (1+rate)^nper + PMT*(1+rate*type)*----------------- + fv  = 0
    '                                                            rate
    '               rate == 0  ->  pv + PMT*nper + fv = 0
    '
    '               Thus:
    '                     (-fv - pv*(1+rate)^nper) * rate
    '               PMT = -------------------------------------
    '                     (1+rate*type) * ( (1+rate)^nper - 1 )
    '
    '               PMT = (-fv - pv) / nper    : if rate == 0
    """
    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 3)):
            return "NULL"

        r = ''
        try:
            # Pull out the arguments.
            rate = float(params[0])
            nper = utils.int_convert(params[1]) + 0.0
            pv = float(params[2])
            fv = 0
            if (len(params) >= 4):
                fv = float(params[3])
            typ = 0
            if (len(params) >= 5):
                typ = float(params[4])

            # Compute the payments.
            if (((1 + rate * typ) * (pow(1 + rate, nper) - 1)) != 0):
                r = ((-fv - pv * pow(1 + rate, nper)) * rate)/((1 + rate * typ) * (pow(1 + rate, nper) - 1))
            else:
                r = 0
        except (ValueError, TypeError, ZeroDivisionError, OverflowError):
            pass

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Pmt: %r returns %r" % (self, r))
        return r

class Randomize(VbaLibraryFunc):
    """
    Randomize RNG function.
    """

    def eval(self, context, params=None):
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Randomize(): Stubbed out as NOP")
        return ''

class Rnd(VbaLibraryFunc):
    """
    Rnd() RNG function.
    """

    def eval(self, context, params=None):
        return random.random()

    def num_args(self):
        return 0

class CDbl(VbaLibraryFunc):
    """
    CDbl() type conversion function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        try:
            # Handle hex.
            tmp = str(params[0]).upper()
            if (tmp.lower().startswith("&h")):
                tmp = tmp.replace("&h", "0x")
                tmp = int(tmp, 16)

            # VBA rounds the significant digits.
            #return round(float(params[0]), 11)
            return float(tmp)

        except Exception as e:
            log.error("CDbl(" + str(params[0]) + ") failed. " + str(e))
            return 0

class CVErr(VbaLibraryFunc):
    """
    CVErr() Excel error string function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        err = None
        try:
            err = int(params[0])
        except (ValueError, TypeError):
            pass
        vals = {2007 : "#DIV/0!",
                2042 : "#N/A",
                2029 : "#NAME?",
                2000 : "#NULL!",
                2036 : "#NUM!",
                2023 : "#REF!",
                2015 : "#VALUE!"}
        if (err in vals):
            return vals[err]
        return ""

class Not(VbaLibraryFunc):
    """
    Boolean Not() called as a function.
    """

    def eval(self, context, params=None):

        if ((len(params) == 0) or (not isinstance(params[0], bool))):
            log.warning("Cannot compute Not(" + str(params) + ").")
            return "NULL"
        return (not params[0])

class VarPtr(VbaLibraryFunc):
    """
    Faked VarPtr() function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return

        # Report on the full byte array given to VarPtr().
        val = params[0]
        context.report_action("External Call", "VarPtr(" + str(val) + ")", "VarPtr", strip_null_bytes=True)

class NumPut(VbaLibraryFunc):
    """
    DynamicWrapperX.NumPut() method. This simulates the NumPut() byte writing actions
    by writing the byte values to a DOM_NumPut.dat file.
    """

    def eval(self, context, params=None):

        if (params is None):
            return

        # Do we need to open the simulated file?
        if ("DOM_NumPut.dat" not in context.open_files):
            context.open_file("DOM_NumPut.dat")

        # Get the byte to write.
        if (len(params) < 3):
            return
        val = params[0]
        pos = params[2]

        # Write the byte.
        # TODO: Use the position parameter to write the byte to the proper position.
        context.write_file("DOM_NumPut.dat", chr(val))

class GetLocale(VbaLibraryFunc):
    """
    GetLocale() Function.
    """

    def eval(self, context, params=None):

        # Match anything compared to this result.
        return "**MATCH ANY**"

class RandBetween(VbaLibraryFunc):
    """
    Excel RANDBETWEEN() function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) < 2)):
            return "NULL"
        lower = coerce_to_int(params[0])
        upper = coerce_to_int(params[1])
        return random.randint(lower, upper)

    def num_args(self):
        return 2
