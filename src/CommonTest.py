import openpyxl
import pandas as pd
from enum import IntEnum, Enum
from typing import List, Optional, Dict


class _Column(IntEnum):
    """
    Enum representing the column indices used in the Results worksheet.
    Using IntEnum allows the values to be used directly as integers.
    """
    ROW     = 1
    NAME    = 2
    EXP_VAL = 3
    OPER    = 4
    ACT_VAL = 5
    CHECK   = 6


class _Result(Enum):
    """
    Enum representing possible validation results.
    """
    PASS = "PASS"
    FAIL = "FAIL"


class _Operation(Enum):
    """
    Enum defining supported comparison operations.

    Each entry contains:
    (Excel operator symbol, human-readable description)
    """
    EQ   = ("="  , "Equals")
    SEQ  = ("="  , "Equals")
    NE   = ("<>" , "Not Equals")
    NEQ  = ("<>" , "Not Equals")
    GE   = (">=" , "Greater Than or Equals")
    GT   = (">"  , "Greater Than")
    LE   = ("<=" , "Less Than or Equals")
    LT   = ("<"  , "Less Than")
    TL   = ("+/-", "Within")
    NTL  = ("+/-", "Not Within")


class _CellFormat(Enum):
    """
    Enum defining reusable cell formatting styles used in the workbook.
    """
    REDFONT = openpyxl.styles.Font(
        color="0000FF",
        bold=True)

    ORANGEFILL = openpyxl.styles.PatternFill(
        start_color="FFA500",
        end_color="FFA500",
        fill_type="solid")

    HYPERLINK = openpyxl.styles.Font(
        color="0000FF",
        underline="single")


class CommonTest:
    """
    Main class responsible for:
    - Loading CSV files
    - Converting them to Excel worksheets
    - Performing validation checks
    - Writing results to a dedicated Results worksheet
    """

    def __init__(self):
        """
        Initializes the CommonTest object and default configuration values.
        """

        # Workbook / worksheet state
        self.workbook   = None
        self.results_ws = None
        self.active_ws  = None
        self.currentResultsRow = None

        # Track total rows and columns for all imported CSV files
        self.total_rows = []
        self.total_cols = []

        # Default column mapping for Results sheet
        self._rowCol    = _Column.ROW
        self._nameCol   = _Column.NAME
        self._expValCol = _Column.EXP_VAL
        self._operCol   = _Column.OPER
        self._actValCol = _Column.ACT_VAL
        self._checkCol  = _Column.CHECK

        # Default result values
        self._passVal = _Result.PASS
        self._failVal = _Result.FAIL

        # Dictionary mapping operation names to their symbols/descriptions
        self._basicOperations = {op.name: op.value for op in _Operation}

        # Cell formatting shortcuts
        self._redFont = _CellFormat.REDFONT
        self._orangeFill = _CellFormat.ORANGEFILL
        self._hyperLinkFont = _CellFormat.HYPERLINK


    def initializeTest(self, csv_files: List, output_file: str) -> None:
        """
        Initializes the test environment.

        Reads multiple CSV files, converts them into Excel worksheets,
        and prepares the Results worksheet.

        Parameters
        ----------
        csv_files : List
            List of CSV file paths.
        output_file : str
            Output Excel file path.
        """

        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            self.workbook = writer.book
            header_format = self.workbook.add_format({'bold': True})

            # Process each CSV file and create a worksheet
            for index, csv_file in enumerate(csv_files):
                self._csvReadAndAddSheet(
                    index=index,
                    csv_file=csv_file,
                    writer=writer,
                    header_format=header_format
                )

            # Create the Results worksheet
            self._createResultsFile()

        # Re-open workbook with openpyxl for editing
        self._openResultsWorkbook(output_file)


    def _csvReadAndAddSheet(self, index, csv_file, writer, header_format):
        """
        Reads a CSV file and adds it as a worksheet to the Excel workbook.

        Also auto-adjusts column widths and freezes the header row.
        """

        sheet_name = f"AnalyzedData-{index+1}"

        # Read CSV into DataFrame
        df = pd.read_csv(csv_file, dtype=str)

        # Write DataFrame to Excel sheet
        df.to_excel(writer, sheet_name=sheet_name, index=False)

        self.active_ws = writer.sheets[sheet_name]

        # Write formatted header row
        for col_num, value in enumerate(df.columns.values):
            self.active_ws.write(0, col_num, value, header_format)

        # Auto-adjust column widths
        for i, col_num  in enumerate(df.columns):
            max_length = max(df[col_num].astype(str).str.len().max(), len(col_num))
            self.active_ws.set_column(i, i, max_length + 2)

        # Freeze header row
        self.active_ws.freeze_panes(1,0)

        rows, cols = df.shape

        # Store dataset dimensions
        self.total_rows.append(rows)
        self.total_cols.append(cols)


    def endTest(self, output_file: str) -> None:
        """
        Saves the final Excel workbook.

        Parameters
        ----------
        output_file : str
            Output file path.
        """
        self.workbook.save(output_file)


    def _createResultsFile(self) -> None:
        """
        Creates the Results worksheet where validation output will be stored.
        """
        self.results_ws = self.workbook.add_worksheet("Results")
        self._formatResultsWorksheet()


    def _formatResultsWorksheet(self) -> None:
        """
        Applies formatting and header row to the Results worksheet.
        """

        header_format = self.workbook.add_format({
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "border": 1
        })

        headers = [
            "Row Index",
            "Name",
            "Actual Value",
            "Operation",
            "Expected Value"
        ]

        # Write headers
        for col_num, header in enumerate(headers):
            self.results_ws.write(0, col_num, header, header_format)

        self.results_ws.freeze_panes(1,0)


    def _openResultsWorkbook(self, output_file: str) -> None:
        """
        Opens the generated Excel workbook using openpyxl.

        This allows further cell modifications and formatting.
        """

        self.workbook   = openpyxl.load_workbook(output_file)
        self.active_ws  = self.workbook.active
        self.results_ws = self.workbook['Results']


    def _activateWorksheet(self, csvSheet: int) -> None:
        """
        Activates a specific worksheet by index.

        Parameters
        ----------
        csvSheet : int
            Sheet number to activate.
        """
        self.active_ws = self.workbook[f"AnalyzedData-{csvSheet}"]


    def getRowNumber(self, searchString: str, colNum: int,
                     csvSheet: Optional[int] = 1,
                     startRow: Optional[int] = None) -> int:
        """
        Finds the first row containing a given value in a specified column.

        Returns the row number if found.
        """

        self._activateWorksheet(csvSheet)
        startRow = startRow or self.active_ws.min_row

        iterator = self.active_ws.iter_rows(
            min_col=colNum,
            max_col=colNum,
            min_row=startRow,
            values_only=True
        )

        for rowNum, rowVal in enumerate(iterator, start=startRow):
            if rowVal[0] == searchString:
                return rowNum


    def getColumnNumber(self, searchString: str,
                        csvSheet: Optional[int] = 1) -> int:
        """
        Finds a column index by searching the header row.
        """

        self._activateWorksheet(csvSheet)

        iterator = self.active_ws.iter_cols(
            min_row=1,
            max_row=1,
            values_only=True
        )

        for colNum, colVal in enumerate(iterator,
                                        start=self.active_ws.min_column):
            if colVal[0] == searchString:
                return colNum


    def findAllRows(self, searchString: str, colNum: int,
                    csvSheet: Optional[int] = 1) -> List:
        """
        Returns all row indices containing a specified value in a column.
        """

        allRows = []

        self._activateWorksheet(csvSheet)

        iterator = self.active_ws.iter_rows(
            min_col=colNum,
            max_col=colNum,
            values_only=True
        )

        for rowNum, rowVal in enumerate(iterator,
                                        start=self.active_ws.min_row):
            if rowVal[0] == searchString:
                allRows.append(rowNum)

        return allRows


    def findRowsIntersect(self, searchStringDict: Dict,
                          csvSheet: Optional[int] = 1) -> List:
        """
        Returns rows that match ALL column/value pairs.

        Performs a set intersection of matching rows.
        """

        intersection = None

        for key, value in searchStringDict.items():
            allRows = set(self.findAllRows(
                searchString=value,
                colNum=key,
                csvSheet=csvSheet))

            intersection = allRows if intersection is None else (intersection & allRows)

        return list(intersection)


    def findRowsUnion(self, searchStringDict: Dict,
                      csvSheet: Optional[int] = 1) -> List:
        """
        Returns rows that match ANY column/value pair.

        Performs a set union of matching rows.
        """

        union = set()

        for key, value in searchStringDict.items():
            union |= set(self.findAllRows(
                searchString=value,
                colNum=key,
                csvSheet=csvSheet))

        return list(union)

    def _returnStrCoordRC(self) -> str:
        """
        Generates relative R1C1-style Excel references for the
        Actual Value and Expected Value columns relative to the
        Check column.

        Returns
        -------
        tuple[str, str]
            Relative references to actual and expected value cells.
        """

        # Relative reference to Actual Value column
        actVal = f'RC[{self._actValCol - self._checkCol}]'

        # Relative reference to Expected Value column
        expVal = f'RC[{self._expValCol - self._checkCol}]'

        return actVal, expVal


    def _basicFormula(self, operator: str) -> str:
        """
        Builds a simple Excel IF formula that compares
        the Actual Value to the Expected Value.

        Parameters
        ----------
        operator : str
            Excel comparison operator (e.g. '=', '<>', '>=', '<').

        Returns
        -------
        str
            Excel IF formula returning PASS or FAIL.
        """

        # Get relative cell references
        actVal, expVal = self._returnStrCoordRC()

        # Construct IF formula
        return f"=IF({actVal} {operator} {expVal}, {self._passVal}, {self._failVal})"


    def _commandTolerance(self, tol: float, cmmd: str) -> str:
        """
        Builds an Excel formula that checks if the Actual Value
        falls within a tolerance range of the Expected Value.

        Parameters
        ----------
        tol : float
            Allowed tolerance value.
        cmmd : str
            Command specifying tolerance behavior (TL or NTL).

        Returns
        -------
        str
            Excel formula evaluating tolerance condition.
        """

        # Get relative cell references
        actVal, expVal = self._returnStrCoordRC()

        # Upper bound check
        check1 = f"{actVal} <= {expVal} + {tol}"

        # Lower bound check
        check2 = f"{actVal} => {expVal} - {tol}"

        # Combine checks into AND condition
        formula = f"AND({check1}, {check2})"

        # If command is "Not Within Tolerance", invert result
        if cmmd in ("NTL"):
            formula = f"NOT({formula})"

        # Wrap final condition in IF statement returning PASS/FAIL
        return f"=IF({formula}, {self._passVal}, {self._failVal})"


    def expectedValuesCheck(self, expectedValue: str, actualValue: str) -> None:
        """
        Parses the expected value command and writes validation
        formulas and information into the Results worksheet.

        Parameters
        ----------
        expectedValue : str
            Expected value string containing command and value.
        actualValue : str
            Actual value retrieved from dataset.
        """

        expVal = None

        # Split command and value components
        stringSplit = expectedValue.trim().split(",", 1)

        # Extract operation command (EQ, NE, TL, etc.)
        commandStr = stringSplit[0]

        # Lookup operator symbol and description
        symbol, description = self._basicOperations.get(commandStr,
                                                        (None, None))

        # Handle tolerance commands
        if commandStr in ("TL","NTL"):

            # Extract expected value and tolerance
            parts = stringSplit[1].rsplit(",", 1)
            expVal, tolerance = parts[0], parts[1]

            # Generate tolerance formula
            val = self._commandWithtinTolerance(tol=tolerance, cmmd=commandStr)

            # Update operation description
            description = f"{description} {symbol} {tolerance}"

        # Handle standard comparison operations
        elif commandStr in self._basicOperations:

            val = self._basicFormula(operator=symbol)
            expVal = stringSplit[1]

        # Invalid command
        else:
            val = None

        # --- Write values to Results worksheet ---

        # Expected value
        self._setResultCellValue(value=expVal,
                                 resultRow=self.currentResultsRow,
                                 resultCol=self._expValCol)

        # Actual value
        self._setResultCellValue(value=actualValue,
                                 resultRow=self.currentResultsRow,
                                 resultCol=self._actValCol)

        # Operation description
        self._setResultCellValue(value=description,
                                 resultRow=self.currentResultsRow,
                                 resultCol=self._operCol)

        # Excel formula check
        self._setResultCellValue(value=val,
                                 resultRow=self.currentResultsRow,
                                 resultCol=self._checkCol)


    def _increaseResultsRow(self) -> None:
        """
        Advances the results row pointer so the next test result
        is written to the following row.
        """
        self.currentResultsRow += 1


    def setCellValue(self, row: int, col: int, value: str,
                     csvSheet: Optional[int] = 1) -> None:
        """
        Sets a cell value in the selected worksheet and highlights
        it using a red font.

        Parameters
        ----------
        row : int
            Row index.
        col : int
            Column index.
        value : str
            Value to write into the cell.
        csvSheet : int, optional
            Sheet index to modify.
        """

        # Activate correct worksheet
        self._activateWorksheet(csvSheet=csvSheet)

        # Write value and apply red font
        self.active_ws(row=row, column=col, value=value).font = self._redFont


    def getCellValue(self, row: int, col: int,
                     csvSheet: Optional[int] = 1) -> str:
        """
        Retrieves a value from a worksheet cell.

        The accessed cell is highlighted with an orange fill.

        Parameters
        ----------
        row : int
            Row index.
        col : int
            Column index.
        csvSheet : int, optional
            Sheet index to read from.

        Returns
        -------
        str
            Cell value.
        """

        # Activate worksheet
        self._activateWorksheet(csvSheet=csvSheet)

        # Highlight cell
        self.active_ws.cell(row=row, column=col).fill = self._orangeFill

        # Return value
        return self.active_ws.cell(row=row, column=col).value


    def _setCellHyperlink(self, dataRow: int, dataCol: int,
                          resultRow: int, resultCol: int,
                          csvSheet: int) -> None:
        """
        Creates a hyperlink in the Results worksheet pointing
        to the corresponding cell in a data worksheet.
        """

        cell = self.results_ws(row=resultRow, column=resultCol)

        # Hyperlink reference to dataset cell
        cell.hyperlink = f"#AnalyzedData-{csvSheet}!R{dataRow}C{dataCol}"

        # Apply hyperlink formatting
        cell.font = self._hyperLinkFont


    def _setResultCellValue(self, value: str,
                            resultRow: int,
                            resultCol: int) -> None:
        """
        Writes a value into a cell in the Results worksheet.
        """

        self.results_ws(row=resultRow, column=resultCol, value=value)


    def _setCellComment(self, value: str,
                        resultRow: int,
                        resultCol: int) -> None:
        """
        Adds a comment to a cell in the Results worksheet.
        """

        self.results_ws(row=resultRow, column=resultCol).comment = value


    def addDataNameResults(self, titleStr: str, dataRow: int,
                           dataCol: Optional[int] = 0,
                           cmmt: Optional[str] = None,
                           csvSheet: Optional[int] = 1) -> None:
        """
        Writes the test name and data row reference to the Results sheet
        and creates hyperlinks back to the original dataset.

        Parameters
        ----------
        titleStr : str
            Test name or description.
        dataRow : int
            Row in the dataset where the value was found.
        dataCol : int
            Column in dataset (optional).
        cmmt : str
            Optional comment.
        csvSheet : int
            Dataset sheet index.
        """

        # Write test name
        self._setResultCellValue(value=titleStr,
                                 resultRow=self.currentResultsRow,
                                 resultCol=self._nameCol)

        # Write dataset row index
        self._setResultCellValue(value=dataRow,
                                 resultRow=self.currentResultsRow,
                                 resultCol=self._rowCol)

        # Add hyperlink to dataset row
        self._setCellHyperlink(dataRow=dataRow,
                               dataCol=1,
                               resultRow=self.currentResultsRow,
                               resultCol=self._rowCol,
                               csvSheet=csvSheet)

        # Optional hyperlink to specific column
        if dataCol > 0:
            self._setCellHyperlink(dataRow=dataRow,
                                   dataCol=dataCol,
                                   resultRow=self.currentResultsRow,
                                   resultCol=self._actValCol,
                                   csvSheet=csvSheet)

        # Add optional comment
        if not cmmt is None:
            self._setCellComment(value=openpyxl.comments.Comment(cmmt))


    def writeResults(self, titleStr: str, dataRow: int,
                     expectedValue: str, actualValue: str,
                     dataCol: Optional[int] = 0,
                     cmmt: Optional[str] = None,
                     csvSheet: Optional[int] = 1) -> None:
        """
        High-level helper function that writes a complete test
        result entry to the Results worksheet.

        This includes:
        - Test name
        - Data row reference
        - Expected value
        - Actual value
        - Validation formula

        Parameters
        ----------
        titleStr : str
            Name of the test.
        dataRow : int
            Dataset row number.
        expectedValue : str
            Expected value command.
        actualValue : str
            Actual value from dataset.
        dataCol : int
            Dataset column index.
        cmmt : str
            Optional comment.
        csvSheet : int
            Dataset sheet index.
        """

        # Write metadata for this test
        self.addDataNameResults(titleStr=titleStr,
                                dataRow=dataRow,
                                dataCol=dataCol,
                                cmmt=cmmt,
                                csvSheet=csvSheet)

        # Perform expected vs actual validation
        self.expectedValuesCheck(expectedValue=expectedValue,
                                 actualValue=actualValue)

        # Move to next result row
        self._increaseResultsRow()