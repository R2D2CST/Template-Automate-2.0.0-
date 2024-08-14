# Native Python Libraries.
import os
import sys
import pathlib

# Third Party Python Libraries.
from tqdm import tqdm  # trange "Same as range but in progress bar"

# Self Built Python Libraries.
from .OSFunc import OSFunc
from .OSFunc.OSFunc import openDirectory
from .Excel import ExcelFile
from .Word import WordFile
from .Perfromance.Performance import performanceMonitor

# We build our constant directory names in our application.
PROJECT_DIR_NAME: str = "projects"
RENDERS_DIR_NAME: str = "renders"
TEMPLATES_DIR_NAME: str = "templates"
DATABASE_FILE_NAME: str = "database.xlsx"
WORD_FILE_NAME: str = "template.docx"
EX_FILE_NAME_FILE_NAME: str = "template.xlsx"


def createProject(projectName: str) -> None:
    """
    Function builds a directory named projects in case it do not exists already.
    Then builds a new project named as the project name passed into the
    function. Finally builds a template directory inside with some excel
    database and templates examples for the user to practice.

    Args:
        projectName: project name to assign.

    Returns: None

    Raises: None
    """

    # We build a progress bar to tack performance, unit='unit/s' .
    progressBar = tqdm(total=100, desc="Progress: ", unit="%", colour="green")

    appPath: str = OSFunc.getAppPath()  # We get the absolute path for the app.
    progressBar.update(4)

    # Validates if a projects path exists
    projectDirPath: str = os.path.join(appPath, PROJECT_DIR_NAME)
    if not os.path.exists(projectDirPath):
        os.mkdir(projectDirPath)  # Build a folder path named projects if not
    newProjectDirPath: str = os.path.join(projectDirPath, projectName)
    progressBar.update(100 // 6)

    # We delete objects that are not in use.
    del appPath
    del projectDirPath

    # Builds a folder path called as @projectName
    if not os.path.exists(newProjectDirPath):
        os.mkdir(newProjectDirPath)
    progressBar.update(100 // 6)

    # Creates a Excel file database
    databasePath: str = os.path.join(newProjectDirPath, DATABASE_FILE_NAME)
    if not os.path.exists(databasePath):
        # Call function to build an Excel database File.
        databaseFile = ExcelFile.ExcelFile(excelBookPath=databasePath)
        databaseFile.buildDatabaseTemplate()
        pass
    progressBar.update(100 // 6)

    templateDirPath: str = os.path.join(newProjectDirPath, TEMPLATES_DIR_NAME)
    if not os.path.exists(templateDirPath):
        os.mkdir(templateDirPath)
    progressBar.update(100 // 6)

    # We delete objects that are not in use.
    del newProjectDirPath
    del databasePath
    del databaseFile

    wordTemplatePath: str = os.path.join(templateDirPath, WORD_FILE_NAME)
    if not os.path.exists(wordTemplatePath):
        # Call function to build a Word template
        wordTemplateFile = WordFile.wordDocument(wordDocumentPath=wordTemplatePath)
        wordTemplateFile.buildTemplate()
        pass
    progressBar.update(100 // 6)

    excelTemplatePath: str = os.path.join(templateDirPath, EX_FILE_NAME_FILE_NAME)
    if not os.path.exists(excelTemplatePath):
        # Call function to build a Excel template
        excelTemplateFile = ExcelFile.ExcelFile(excelBookPath=excelTemplatePath)
        excelTemplateFile.buildTemplateStructure()
        pass
    progressBar.update((100 // 6))
    progressBar.close()

    # We delete objects that are not in use.
    del progressBar
    del wordTemplatePath
    del wordTemplateFile
    del templateDirPath
    del excelTemplatePath
    del excelTemplateFile

    return None


def enumerateProjectsInside() -> list[str]:
    """
    Returns a list of all projects inside the app.

    Args: None.

    Returns:
    listProjectsNames (list[str]) : returns a list of projects names.
    listProjectsPaths (list[str]) : returns a list of projects paths.
    """
    try:
        appPath: str = OSFunc.getAppPath()
        projectDirPath: str = os.path.join(appPath, PROJECT_DIR_NAME)
        del appPath
        listProjectPaths: list[str] = OSFunc.listFolderPaths(projectDirPath)
        listProjectNames: list[str] = OSFunc.listFolderNames(projectDirPath)
        return listProjectNames, listProjectPaths
    except FileNotFoundError:
        print("¡¡¡File not found, please make sure you have at least one project!!!")
        return list(), list()


def setHighPerformance() -> None:
    """
    Function elevates the process into high priority for better performance.
    """
    osType: str = OSFunc.getOS()
    OSFunc.setHighPriorityApp(osType=osType)
    return None


def renderExcelTemplates(projectPath: str) -> None:
    """
    Function meant to read the Excel Database and render all excel template
    files into a rendered path inside the project directory.

    Args:
        projectPath (str):

    Returns: None
    """

    # Construir rutas de directorios y archivo
    templatesDirPath = os.path.join(projectPath, TEMPLATES_DIR_NAME)
    renderDirPath = os.path.join(projectPath, RENDERS_DIR_NAME)
    databasePath = os.path.join(projectPath, DATABASE_FILE_NAME)

    # Obtener listado de nombres de plantillas
    listTemplatesFilenames: list[str] = OSFunc.listFileNamesWithExt(
        absolutePath=templatesDirPath
    )

    # Filtrar solo los archivos de Excel
    inputPaths: list[str] = [
        os.path.join(templatesDirPath, fileName)
        for fileName in listTemplatesFilenames
        if fileName.endswith((".xls", ".xlsm", ".xlsx"))
    ]

    # We delete objects that are not in use.
    del templatesDirPath

    # Leer el contenido del archivo de base de datos en forma de diccionario
    databaseFile = ExcelFile.ExcelFile(excelBookPath=databasePath)
    runDict: dict[str, dict[str, tuple]] = databaseFile.readDatabaseExcelSheet()

    # We delete objects that are not in use.
    del databaseFile

    # We build the progress bar indicator.
    runs = len(runDict.keys())
    progressBar = tqdm(
        total=runs, desc="Template Runs: ", unit="run per ", colour="green"
    )

    # We delete objects that are not in use.
    del runs

    # Iterar sobre cada corrida en el diccionario
    for runKey, context in runDict.items():
        # Crear un directorio para cada corrida en "renders"
        runDirPath = os.path.join(renderDirPath, runKey)
        if not os.path.exists(runDirPath):
            os.makedirs(runDirPath)

        # Preparar listas para el renderizado
        sheetList, cellList, valueList = [], [], []

        for header in context.keys():
            sheetValue, cellValue, valueValue = context[header]
            sheetList.append(sheetValue)
            cellList.append(cellValue)
            valueList.append(valueValue)

        # Iterar sobre las plantillas y renderizar cada una
        for inputTemplatePath in inputPaths:
            templateName = os.path.basename(inputTemplatePath)
            outputTemplatePath = os.path.join(runDirPath, templateName)

            # Renderizar la plantilla usando la función correspondiente
            ExcelFile.ExcelFile.renderTemplate(
                inputTemplatePath=inputTemplatePath,
                outputTemplatePath=outputTemplatePath,
                sheetList=sheetList,
                cellList=cellList,
                valueList=valueList,
            )

        # We update 1 run
        progressBar.update(1)

    # We end pur progress bar.
    progressBar.close()
    print("Renderizado completado.")

    # We delete objects that are not in use.
    del progressBar

    return None


def renderWordTemplates(projectPath: str) -> None:
    """
    Function meant to read the Excel Database and render all word template
    files into a rendered path inside the project directory.

    Args:
        selectedProjectPath (str):

    Returns: None
    """

    # We open the excel file database and read is content as follows:
    # We read its content and ensamble a dictionary that wil contain:
    # runKeyWord : {keyWord: Values dictionary}
    databaseFilePath = os.path.join(projectPath, DATABASE_FILE_NAME)
    databaseFile = ExcelFile.ExcelFile(excelBookPath=databaseFilePath)
    renderContext = databaseFile.readDatabaseWordSheet()

    # We delete objects that are not in use.
    del databaseFilePath
    del databaseFile

    # We build the progress bar.
    runs = len(renderContext.keys())
    progressBar = tqdm(
        total=runs, desc="Template Runs: ", unit="run per ", colour="green"
    )

    # We delete objects that are not in use.
    del runs

    # We get the absolute path for all files inside the templates directory.
    listTemplatesFilenames: list[str] = []
    templatesDirPath = os.path.join(projectPath, TEMPLATES_DIR_NAME)
    listTemplatesFilenames = OSFunc.listFileNamesWithExt(absolutePath=templatesDirPath)

    # We build a list of the input file paths.
    inputPaths: list[str] = list()
    for fileName in listTemplatesFilenames:
        if fileName.endswith((".doc", ".docx", ".docm")):
            inputPaths.append(os.path.join(templatesDirPath, fileName))

    # We build a directory output path for each run.
    renderRunPaths: list[str] = []
    rendersDirPath: str = os.path.join(projectPath, RENDERS_DIR_NAME)
    for runName in renderContext.keys():
        renderRunPaths.append(os.path.join(rendersDirPath, runName))
    for path in renderRunPaths:
        if not os.path.exists(path):
            os.makedirs(path)

    # We build a key list to retrieve the dictionary context.
    keyList: list[str] = []
    for key in renderContext.keys():
        keyList.append(key)

    # We start a loop for each "run".
    i: int = 0
    j: int = 0
    for runPath in renderRunPaths:

        # We build a list of the output file paths.
        outputPaths: list[str] = list()
        for fileName in listTemplatesFilenames:
            if fileName.endswith((".doc", ".docx", ".docm")):
                outputPaths.append(os.path.join(runPath, fileName))

        for input in inputPaths:
            context = renderContext[keyList[j]]
            output = outputPaths[i]
            WordFile.wordDocument.renderTemplate(
                templatePath=input, outputPath=output, context=context
            )
            i += 1
        i = 0
        j += 1

        # We Update 1 run in our progress bar.
        progressBar.update(1)
        del outputPaths

    # We close the progress bar.
    progressBar.close()

    # We delete objects that are not in use.
    del progressBar

    return None


def renderAllTemplates(projectPath: str) -> None:
    """
    Function meant to read the Excel Database and render all word and excel
    template files into a rendered path inside the project directory.

    Args:
        projectPath (str):

    Returns: None
    """

    renderExcelTemplates(projectPath=projectPath)
    renderWordTemplates(projectPath=projectPath)
    return None


def startMonitoringPerformance() -> None:
    """
    Starts a thread to monitor the performance of the application and log it to a CSV file.
    """
    appPath: str = OSFunc.getAppPath()  # Obtains the directory where the app is running
    performanceMonitor(directory=appPath, sleepTime=5)
    return None
