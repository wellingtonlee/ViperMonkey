"""ViperMonkey: VBA Library — Excel, spreadsheet, and document functions."""
from ._common import *

class ExecuteExcel4Macro(VbaLibraryFunc):
    """
    ExecuteExcel4Macro() dynamic XLM evaluation function.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) == 0)):
            return 0
        xlm = str(params[0])
        context.report_action('XLM Macro Execution', xlm, 'Dynamic XLM Macro Execution', strip_null_bytes=True)
        return 0

    def num_args(self):
        return 1

class BuiltInDocumentProperties(VbaLibraryFunc):
    """
    Simulate calling ActiveDocument.BuiltInDocumentProperties('PROPERTYNAME')
    """

    def eval(self, context, params=None):

        if ((params is None) or (len(params) < 1)):
            return "NULL"

        # Get the property we are looking for.
        prop = str(params[0])
        r = context.read_metadata_item(prop)
        if (r == ""):
            r = "NULL"
        return r

    def num_args(self):
        return 1

    def return_type(self):
        return "STRING"

class Item(BuiltInDocumentProperties):
    """
    Assumes that Item() is only called on BuiltInDocumentProperties.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Were we given the dict to work with?
        with_dict = None
        index = None
        if ((len(params) >= 2) and (isinstance(params[0], dict))):

            # Dict is 1st parameter.
            with_dict = params[0]
            # Item index is 2nd parameter.
            index = coerce_to_int(params[1])

        # Are we reading from a With Scripting.Dictionary?
        elif ((context.with_prefix_raw is not None) and
              (context.contains(str(context.with_prefix_raw)))):

            # Get the item index.
            index = None
            try:
                index = coerce_to_int(params[0])
            except (ValueError, TypeError):
                return "NULL"

            # Is the With variable value a dict?
            with_dict = context.get(str(context.with_prefix_raw))
            if (not isinstance(with_dict, dict)):
                with_dict = None

        # Are we reading from a Scripting.Dictionary?
        if (with_dict is not None):

            # Valid key?
            if (index in with_dict):
                return with_dict[index]
            return "NULL"

        # Not a workable Scripting.Dictionary.Item() call. Treat as
        # BuiltInDocumentProperties.Item()
        return super(Item, self).eval(context, params)

class International(VbaLibraryFunc):
    """
    application.international() Function.
    """

    def eval(self, context, params=None):

        # Match anything compared to this result.
        return "**MATCH ANY**"

class Shapes(VbaLibraryFunc):
    """
    Shapes() object reference. Stubbed.
    """

    def eval(self, context, params=None):

        # Just return the string representation of the access. This is used in
        # vba_object._read_from_object_text()
        if ((params is None) or (len(params) == 0)):
            return ""
        return "Shapes('" + str(params[0]) + "')"

class InlineShapes(VbaLibraryFunc):
    """
    InlineShapes() object reference. Stubbed.
    """

    def eval(self, context, params=None):

        # Just return the string representation of the access. This is used in
        # vba_object._read_from_object_text()
        if ((params is None) or (len(params) == 0)):
            return ""
        return "InlineShapes('" + str(params[0]) + "')"

class Paragraphs(VbaLibraryFunc):
    """
    Get a specific paragraph.
    """

    def eval(self, context, params=None):

        # Get the paragraphs.
        paragraphs = None
        try:
            paragraphs = context.get("ActiveDocument.Paragraphs".lower())
        except KeyError:
            return "NULL"

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            log.error("Paragraphs() called with no arguments. Returning all paragraphs.")
            return paragraphs

        # Get the paragraph index.
        index = None
        try:
            index = coerce_to_int(params[0])
        except (ValueError, TypeError):
            log.error("%r is not a valid index value. Returning NULL." % params[0])
            return "NULL"

        # Do we have a paragraph at this index?
        if (index >= len(paragraphs)):
            log.error("Paragraphs(" + str(index) + ") out of range. Returning NULL.")
            return "NULL"

        # Return the paragraph.
        r = paragraphs[index]
        return r

class CheckSpelling(VbaLibraryFunc):
    """
    Application.CheckSpelling() function. Currently stubbed.
    """

    def eval(self, context, params=None):

        # TODO: Find and use a Python spell checker to check the spelling
        # of the argument.

        # For now just say everything is correctly spelled.
        return True

class Specialfolders(VbaLibraryFunc):
    """
    Excel Specialfolders() function. Currently stubbed.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        return "%" + str(params[0]) + "%"

class Rows(VbaLibraryFunc):
    """
    This emulates geting the Rows field of an Excel sheet.
    Currently stubbed out to just return a list of dicts with row numbers.
    """

    def eval(self, context, params=None):

        # Do we have a loaded Excel file?
        if (context.loaded_excel is None):
            context.increase_general_errors()
            log.warning("Cannot emulate Rows field access. No Excel file loaded.")
            return []

        # Get the sheet name if given.
        sheet_name = "__NO SHEET NAME__"
        if ((params is not None) and (len(params) > 0)):
            sheet_name = str(params[0]).strip()

        # Get the sheet from which to pull rows.
        sheet = None
        if (sheet_name in context.loaded_excel.sheet_names()):
            sheet = context.loaded_excel.sheet_by_name(sheet_name)

        # No sheet name. Just find the sheet with the most rows.
        else:

            # Look through all sheets.
            max_rows = -1
            for sheet_index in range(0, len(context.loaded_excel.sheet_names())):

                # Load the current sheet.
                curr_sheet = None
                try:
                    curr_sheet = context.loaded_excel.sheet_by_index(sheet_index)
                except Exception:
                    context.increase_general_errors()
                    log.warning("Cannot process Cells() call. No sheets in file.")
                    return "NULL"

                # Does this have the most rows?
                curr_rows = get_num_rows(curr_sheet)
                if (curr_rows > max_rows):
                    max_rows = curr_rows
                    sheet = curr_sheet

        # Return a list of dicts with row info.
        r = []
        num_rows = get_num_rows(sheet)
        for i in range(0, num_rows + 1):
            r.append({ "Row" : i })

        # Return the row info.
        return r

    def num_args(self):
        return 0

class Cells(VbaLibraryFunc):
    """
    Excel Cells() function.
    Currently only handles Cells(x, y) calls.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            log.error("Parameters of Cells() call are None or empty.")
            return "NULL"

        # Do we have a loaded Excel file?
        if (context.loaded_excel is None):
            context.increase_general_errors()
            log.warning("Cannot process Cells() call. No Excel file loaded.")
            return "NULL"

        # Is this actually the Cells field of a Sheet?
        if ("Sheet" in str(type(params[0]))):

            # Pull out all the cells  in the sheet and return that.
            return excel.pull_cells_sheet(params[0])

        # The row and column must always be the 1st 2 arguments.
        if (len(params) < 2):
            context.increase_general_errors()
            log.warning("Cells() called with < 2 arguments. Returning NULL.")
            return "NULL"

        # The sheet will be the 3rd argument, if given.
        sheet = None
        if (len(params) >= 3):
            sheet = params[2]
            if (not hasattr(sheet, "cell")):
                log.warning("3rd Cells() argument is not sheet object. Returning NULL.")
                return "NULL"

        # Get the indices of the cell.

        # Column.
        tmp = params[1]
        if (isinstance(tmp, dict) and ("value" in tmp)):
            tmp = tmp["value"]
        col = None
        try:
            col = int(tmp) - 1
        except (ValueError, TypeError):
            try:
                col = excel_col_letter_to_index(tmp)
            except (ValueError, TypeError):
                context.increase_general_errors()
                log.warning("Cannot process Cells() call. Column " + str(params[1]) + " invalid.")
                return "NULL"

        # Row.
        tmp = params[0]
        if (isinstance(tmp, dict) and ("value" in tmp)):
            tmp = tmp["value"]
        row = None
        try:
            row = int(tmp) - 1
        except (ValueError, TypeError):
            context.increase_general_errors()
            log.warning("Cannot process Cells() call. Row " + str(params[0]) + " invalid.")
            return "NULL"

        # If we were not given a sheet use the current active sheet (if we know it).
        if ((sheet is None) and
            hasattr(context.loaded_excel, "active_sheet_name")):
            try:
                sheet_name = context.loaded_excel.active_sheet_name
                sheet = context.loaded_excel.sheet_by_name(sheet_name)
            except ValueError as e:
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Can't find active sheet. " + str(e))

        # Now try the sheet with the most cells if we still need to guess the sheet.
        if (sheet is None):
            sheet = get_largest_sheet(context.loaded_excel)

        # Punt if we still have no idea what sheet to work with.
        if (sheet is None):
            return "NULL"

        # Return the cell contents.
        cell_val = _read_cell(sheet, row, col)
        if (cell_val is not None):
            return cell_val

        # The largest sheet or given sheet did not work. Try each sheet until we read a cell.
        for sheet_index in range(0, len(context.loaded_excel.sheet_names())):

            # Load the current sheet.
            sheet = None
            try:
                sheet = context.loaded_excel.sheet_by_index(sheet_index)
            except Exception:
                context.increase_general_errors()
                log.warning("Cannot process Cells() call. No sheets in file.")
                return "NULL"

            # Return the cell contents.
            cell_val = _read_cell(sheet, row, col)
            if (cell_val is not None):
                return cell_val
            continue

        # Can't read the cell.
        context.increase_general_errors()
        log.warning("Failed to read Cell(" + str(col) + ", " + str(row) + "). (1)")
        return "NULL"

class Sheets(VbaLibraryFunc):
    """
    Excel Sheets() function.
    """

    def eval(self, context, params=None):

        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("Get sheet Sheets(" + str(params) + ") ...")

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Sheets() params are bad. Returning None")
            return None

        # Do we have a loaded Excel file?
        if (context.loaded_excel is None):
            context.increase_general_errors()
            log.warning("Cannot process Sheets() call. No Excel file loaded.")
            return "NULL"

        # Get the sheet with the given identifier.
        sheet_id = str(params[0])

        # First try treating this as a sheet name.
        try:
            curr_sheet = context.loaded_excel.sheet_by_name(str(sheet_id))
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Returning sheet with name '" + str(sheet_id) + "'")
            return curr_sheet
        except Exception as e:
            pass

        # Next see if the sheet ID is an index.
        try:
            sheet_id = int(sheet_id) - 1
        except Exception as e:
            return None
        try:
            curr_sheet = context.loaded_excel.sheet_by_index(sheet_id)
            if (log.getEffectiveLevel() == logging.DEBUG):
                log.debug("Returning sheet with index " + str(sheet_id))
            return curr_sheet
        except Exception as e:
            log.warning("Did not find sheet with index " + str(sheet_id))
            return None

class Worksheets(Sheets):
    """
    Excel Worksheets() function.
    """
    pass

class Value(VbaLibraryFunc):
    """
    Excel cell  Value() function (actually field).
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            return "NULL"

        # Get the cell value.
        cell = params[0]
        r = cell
        if (isinstance(cell, dict) and ("value" in cell)):
            r = cell["value"]
        #print "CELL VALUE!!"
        #print r
        return r

class UsedRange(VbaLibraryFunc):
    """
    Excel UsedRange() function.
    """

    def eval(self, context, params=None):

        # Try each sheet and return the cells from the sheet with the most cells
        # if no sheet is given.
        sheet = None
        if ((params is not None) and
            (len(params) >= 1) and
            ("Sheet" in str(type(params[0])))):
            sheet = params[0]
        else:
            sheet = get_largest_sheet(context.loaded_excel)
        if (sheet is None):
            return []

        # Return all of the defined cells. Each cell is represented as a dict
        # with keys 'value', 'row', and 'col'.
        r = pull_cells_sheet(sheet)
        return r

    def num_args(self):
        return 0

class Range(VbaLibraryFunc):
    """
    Excel Range() function.
    """

    def _get_row_and_column(self, cell_str):
        """
        Get a numeric row and column from a "i93" style Excel cell reference.
        """

        # Pull out the cell index.
        cell_index = str(cell_str).replace('"', "").replace("'", "")

        # Pull out the cell column and row.
        col = ""
        row = ""
        for c in cell_index:
            if (c.isalpha()):
                col += c
            else:
                row += c

        # Convert the row and column to numeric indices for xlrd.
        row = int(row) - 1
        col = excel_col_letter_to_index(col)

        # Done.
        return (row, col)

    def _read_cell_list(self, sheet, cell_str, return_dict):
        """
        Read multiple cells specified by a "i93:i424" cell string.
        """

        # Get the start and end cell.
        fields = cell_str.split(":")
        if (len(fields) != 2):
            log.warning("Improper cell range " + cell_str + " specified. Range() is returning NULL.")
            return "NULL"
        start = fields[0]
        end = fields[1]

        # Get start and end rows and columns.
        start_row, start_col = self._get_row_and_column(start)
        end_row, end_col = self._get_row_and_column(end)

        # Read all the cells, in row by row order.
        r = []
        row_incr = 0
        if ((end_row - start_row) != 0):
            row_incr = (end_row - start_row)/abs(end_row - start_row)
        col_incr = 0
        if ((end_col - start_col) != 0):
            col_incr = (end_col - start_col)/abs(end_col - start_col)
        curr_row = start_row
	#print "=========== READ CELLS!!! ================"
        while (curr_row != end_row):
            curr_col = start_col
            while True:
                val = None
                try:
                    if return_dict:
                        # Return actual dict, not str.
                        val = sheet.cell_dict(curr_row, curr_col)
                    else:
                        val = str(sheet.cell_value(curr_row, curr_col))
                except Exception:
                    pass
                if (val is not None):
                    #print "(" + str(curr_row) + ", " + str(curr_col) + ")"
                    #print "'" + str(val) + "'"
                    r.append(val)
                if (curr_col == end_col):
                    break
                curr_col += col_incr
            if (curr_row == end_row):
                break
            curr_row += row_incr

        # Return the cell values.
        #print "=========== CELLS!!! ================"
        #print cell_str
        #print r
        #sys.exit(0)
        return r

    def eval(self, context, params=None):

        # TODO: Need to track the index of each cell for full
        # emulation of a range. Probably need a Range object
        # implementation.

        # Sanity check.
        if (params is None):
            log.warning("Range() called with no parameters.")
            return "NULL"

        # Do we have a loaded Excel file?
        if (context.loaded_excel is None):

            # It can be the case that we have Range object in Word macro
            if len(params) == 2 and isinstance(params[0], int) and isinstance(params[1], int):
                return context.globals["activedocument.content.text"][params[0]:params[1]]

            else:
                context.increase_general_errors()
                log.warning("Cannot process Range() call. No Excel file loaded.")
                return "NULL"

        # Were we given an Excel sheet object?
        sheet = None
        for p in params:
            if ("Sheet" in str(type(p))):
                sheet = p
                break

        # Return a cell dict rather than the cell value?
        return_dict = False
        if len(params) >= 2 and params[1] is True:
            return_dict = True

        # Currently only handles Range(x) calls.
        if ((len(params) != 1) and (not return_dict) and (sheet is None)):
            context.increase_general_errors()
            log.warning("Only 1 argument Range() calls supported. Returning NULL.")
            return "NULL"

        # Was Range() called on a single, already read cell?
        if (isinstance(params[0], dict)):

            # This is an indirect cell read. The cell address should be in the
            # value of the current cell.
            the_cell = params[0]
            if ("value" in the_cell):
                next_index = the_cell["value"]
                new_params = [next_index]
                for p in params[1:]:
                    new_params.append(p)
                return self.eval(context, new_params)

            # Unexpected. This is not a proper read cell dict.
            log.warning("Unexpected cell dict " + str(the_cell) + ". Range() returning NULL.")
            return "NULL"

        # If we were given a sheet, only look there for the cell. Otherwise
        # look at all sheets.
        sheets = None
        if (sheet is not None):
            sheets = [sheet]
        else:
            sheets = []
            for sheet_index in range(0, len(context.loaded_excel.sheet_names())):
                sheet = None
                try:
                    sheet = context.loaded_excel.sheet_by_index(sheet_index)
                    sheets.append(sheet)
                except Exception:
                    context.increase_general_errors()
                    log.warning("Cannot process Range() call. No sheets in file.")
                    return "NULL"

        # Try the given sheets until we read a cell.
        r = None
        col = None
        for sheet in sheets:

            # Multiple cells?
            range_index = str(params[0])
            if (":" in range_index):
                try:
                    return self._read_cell_list(sheet, range_index, return_dict)
                except Exception as e:
                    # Try the next sheet.
                    continue

            # Get the cell contents.
            try:

                # Pull out the cell value.
                row, col = self._get_row_and_column(params[0])
                if return_dict:
                    # Return actual dict, not str.
                    val = sheet.cell_dict(row, col)
                else:
                    val = str(sheet.cell_value(row, col))

                # Return the cell value.
                sheet_name = ""
                if (hasattr(sheet, "name")):
                    sheet_name = sheet.name
                log.info("Read cell (" + range_index + ") from sheet " + str(sheet_name) + " = '" + str(val) +"'")
                return val

            except Exception as e:
                # Try the next sheet.
                if (log.getEffectiveLevel() == logging.DEBUG):
                    log.debug("Cell read failed. " + str(e))
                continue

        # We did not get the cell.
        row = "??"
        col = "??"
        try:
            row, col = self._get_row_and_column(params[0])
        except (ValueError, TypeError):
            pass
        #print sheet
        log.warning("Failed to read cell (" + str(row) + ", " + str(col) + ") [" + str(params[0]) + "] (2)")
        context.increase_general_errors()
        return "NULL"

class CountA(VbaLibraryFunc):
    """
    Excel CountA() function.
    """

    def eval(self, context, params=None):

        # Sanity check.
        if ((params is None) or (len(params) == 0)):
            log.warning("No arguments passed to CountA(). Returning NULL")
            return "NULL"
        if (not isinstance(params[0], list)):
            log.warning("CountA() needs list argument, not " + str(type(params[0])) + ". Returning NULL")
            return "NULL"

        # Return a count of all the non-empty cells.
        r = len(params[0])
        return r

class SpecialCells(VbaLibraryFunc):
    """
    Excel SpecialCells() method. Not directly used.
    """

    def eval(self, context, params=None):

        # 1st arg should be a list of cell values, 2nd arg the type of cell to include.
        if ((params is None) or (len(params) < 2)):
            log.warning("Not enough arguments passed to SpecialCells(). Returning NULL")
            return "NULL"

        # Sometimes the args are swapped. Handle that.
        cells = None
        cell_type = None
        if (isinstance(params[0], list) and isinstance(params[1], int)):
            cells = params[0]
            cell_type = params[1]
        if (isinstance(params[1], list) and isinstance(params[0], int)):
            cells = params[1]
            cell_type = params[0]
        if (cells is None):
            log.warning("Incorrect argument types passed to SpecialCells(). Returning NULL")
            return "NULL"
        #if (cell_type != 2):
        #    log.warning("Only handling SpecialCells(xlCellTypeConstants). Returning NULL")
        #    return "NULL"

        # Currently only handling cell type xlCellTypeConstants.
        r = []
        for cell in cells:
            cell_value = str(cell)
            if (isinstance(cell, dict)):
                cell_value = str(cell["value"])
            if (len(cell_value) == 0):
                continue
            if (not cell_value.startswith("=")):
                r.append(cell)

        # Done.
        return r

class Variable(VbaLibraryFunc):
    """
    Get document variable.
    """

    def eval(self, context, params=None):
        if ((params is None) or (len(params) < 1)):
            return "NULL"
        var = str(params[0]).strip()
        var = var.replace("activedocument.customdocumentproperties(", "").\
              replace(")", "").\
              replace("'","").\
              replace('"',"").\
              replace('.value',"").\
              strip()
        r = context.get_doc_var(var)
        if (r is None):
            r = ""
        if (log.getEffectiveLevel() == logging.DEBUG):
            log.debug("ActiveDocument.Variable(" + var + ") = " + str(r))
        return r

class Variables(Variable):
    pass
