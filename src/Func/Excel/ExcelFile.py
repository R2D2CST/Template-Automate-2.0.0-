# Native Python Libraries.

# Third Party Python Libraries.
"We use xlwings as our excel object file constructor."
import xlwings as xl

# Self Built Python Libraries.


class ExcelFile:
    """
    Class with methods for Excel Handling
        Att:
        excelBookPath (str): Excel book absolute path including its
        file extension.
        Methods:
            . builtDatabaseTemplate (None)->None:
                builds a database excel template model.
    """

    def __init__(self, excelBookPath) -> None:
        self.excelBookPath = excelBookPath
        self.WORD_DB_SHEET = "WordDB"
        self.EXCEL_DB_SHEET = "ExcelDB"
        pass

    def buildDatabaseTemplate(self) -> None:
        """
        Function takes the absolute file path and generates corresponding excel
        file (.xlsx) for a template of how to build a proper database.
        """
        # Creates an excel book
        wb = xl.Book()

        # Creates sheet for rendering word documents as 'wordDB'
        wordDB = wb.sheets.add(self.WORD_DB_SHEET)

        # Creates sheet for rendering excel documents as 'excelDB'
        excelDB = wb.sheets.add(self.EXCEL_DB_SHEET)

        # Data for 'wordDB' sheet
        wordHeaders = [
            "keyWordRunName",
            "Keyword Header 1",
            "Keyword Header 2",
            "Keyword Header 3",
        ]
        wordData = [
            ["keyWord1", "keyWord2", "keyWord3", "keyWord4"],
            ["keyWordValue1", "keyWordValue2", "keyWordValue3", "keyWordValue4"],
            ["keyWordValue5", "keyWordValue6", "keyWordValue7", "keyWordValue8"],
            ["keyWordValue9", "keyWordValue10", "keyWordValue11", "keyWordValue12"],
        ]

        # Fill headers and data for 'wordDB' sheet in one go
        wordDB.range("A1:A4").value = [[header] for header in wordHeaders]
        wordDB.range("B1:E4").value = wordData

        # We delete objects that are not in use.
        del wordDB
        del wordHeaders
        del wordData

        # Data for 'excelDB' sheet
        excelHeaders = [
            "Type your headers",
            "Header 1",
            "Header 2",
            "Header 3",
        ]
        excelData = [
            ["Sheet Name Pointer", "Sheet1", "Sheet1", "Sheet1"],
            ["Cell Position Pointer", "A1", "A2", "A3"],
            ["Run 1", "keyWordRunName1", "Value1", "Value2"],
            ["Run 2", "keyWordRunName2", "Value3", "Value4"],
            ["Run 3", "keyWordRunName3", "Value5", "Value6"],
        ]

        # Fill headers and data for 'excelDB' sheet in one go
        excelDB.range("A1:D1").value = [excelHeaders]
        excelDB.range("A2:D6").value = excelData

        # We delete objects that are not in use.
        del excelDB
        del excelHeaders
        del excelData

        # We delete the default sheet (Sheet) if exists.
        try:
            defaultSheet = wb["Sheet"]
            wb.remove(defaultSheet)
        except Exception as e:
            print(f"Error presentado al crear el archivo de Excel: {e}")

        # Save the workbook
        wb.save(self.excelBookPath)
        wb.close()

        return None

    def buildTemplateStructure(self) -> None:
        """
        Function takes the absolute file path and generates corresponding excel
        file (.xlsx) for a template of how to build a rendering template.
        """

        # Creates an excel book
        wb = xl.Book()

        try:
            # Creates sheet for rendering word documents as 'wordDB'
            sheet1 = wb.sheets.add("Sheet1")
        # We delete the default sheet (Sheet) if exists.
        except Exception as e:
            sheet1 = wb.sheets["Sheet1"]
            print("{e}")

        # Data for 'Sheet 1' sheet
        sheet1Headers = ("Watch Above the change",)
        # Fill headers and data for 'wordDB' sheet in one go
        sheet1.range("B1:B3").value = sheet1Headers

        # We delete objects that are not in use.
        del sheet1
        del sheet1Headers

        # Save the workbook
        wb.save(self.excelBookPath)
        wb.close()

        return None

    def readDatabaseWordSheet(self) -> dict:
        """
        Opens a excel database file, reads the content in the wordSheet then
        builds a dictionary of runKeyWord : {keyword: value}.

        Args: None
        Return: renderContext (dict[str:dict[str:str]): dictionary built as
                runKeyWord : {keyword: value}
        Raises: None
        """

        # We prepare the workbook and sheet to read.
        renderContext: dict = dict()
        wb = xl.Book(self.excelBookPath)
        sheet = wb.sheets[self.WORD_DB_SHEET]
        data = sheet.range("A1").expand().value
        wb.close()

        # We delete objects that are not in use.
        del wb
        del sheet

        # headers = data[0]  # Gets all the headers (row 1).

        keys = data[0]  # Gets all dictionary keys (row 2).

        # Gets all keyRunWords values (from row 2 onward).
        runKeys = [row[0] for row in data[1:]]

        # Gets all the values in the database.
        values = [
            row[1:] for row in data[2:]
        ]  # Valores a emparejar (omitimos la primera columna).
        values = [valuesInRow for valuesInRow in data[1:]]

        # We ensamble the dictionary.
        renderContext = {}
        i: int = 0  # Remember values is not a list is a matrix we need this counter.
        for run_key in runKeys:
            runDictionary: dict = {}
            for j in range(len(keys)):
                runDictionary[keys[j]] = values[i][j]
            renderContext[run_key] = runDictionary
            i += 1
            del runDictionary

        return renderContext

    def readDatabaseExcelSheet(self) -> dict:
        """
        Reads the excelDB sheet and constructs a dictionary where each "corrida"
        (run) contains the headers with corresponding sheet, cell, and value.
        Args: None
        Returns:
            runDict (dict): A dictionary with the structure:
            {
                "Run Name": {
                    "Header 1": (sheet, cell, value),
                    ...
                },
                ...
            }
        Raise: None
        """
        # Open the workbook and the specified sheet
        wb = xl.Book(self.excelBookPath)
        sheet = wb.sheets[self.EXCEL_DB_SHEET]

        # Read all data starting from A1 and expanding to the used range
        data = sheet.range("A1").expand().value
        wb.close()

        # We delete objects that are not in use.
        del wb
        del sheet

        # Extract headers and the sheet/cell pointers
        headers = data[0][1:]  # Skip the first element as it's the 'Type your headers'
        sheets = data[1][1:]  # Sheet Name Pointers
        cells = data[2][1:]  # Cell Position Pointers

        # Initialize the dictionary to store the results
        runDict = {}

        # Iterate through each run (starting from the 4th row)
        for row in data[3:]:
            runName = row[0]  # First column in each row represents the run name
            values = row[1:]  # Values corresponding to the headers

            # Create a dictionary for the current run
            context = {}
            for header, sheetName, cell, value in zip(headers, sheets, cells, values):
                context[header] = (sheetName, cell, value)

            # Add the run dictionary to the result dictionary under the run name
            runDict[runName] = context

        return runDict

    @staticmethod
    def renderTemplate(
        inputTemplatePath: str,
        outputTemplatePath: str,
        sheetList: list,
        cellList: list,
        valueList: list,
    ) -> None:
        """
        Renderiza una plantilla de Excel en función de los valores proporcionados y guarda el resultado
        en la ruta de salida especificada.

        Args:
            inputTemplatePath (str): Ruta absoluta del archivo de plantilla de Excel a utilizar.
            outputTemplatePath (str): Ruta absoluta donde se guardará el archivo renderizado.
            sheetList (list[str]): Lista de nombres de hojas donde se escribirán los valores.
            cellList (list[str]): Lista de celdas en las cuales se colocarán los valores.
            valueList (list[str]): Lista de valores que se insertarán en las celdas correspondientes.

        Returns:
            None
        """
        # Abre el libro de Excel a partir de la plantilla de entrada.
        wb = xl.Book(inputTemplatePath)

        # Itera sobre la lista de hojas, celdas y valores para insertar los valores en las celdas correspondientes.
        for i in range(len(sheetList)):
            sheet = wb.sheets[sheetList[i]]
            sheet.range(cellList[i]).value = valueList[i]

        # Guarda el libro de Excel en la ruta de salida especificada.
        wb.save(outputTemplatePath)
        wb.close()

        del wb  # We delete objects that are not in use.

        return None

    pass
