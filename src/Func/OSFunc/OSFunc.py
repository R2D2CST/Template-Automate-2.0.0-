# Native Python Libraries.
# We import os and sys in order to handle the operative system.
import os
import sys
import platform  # Retrieves as much platform-identifying data as possible.
import psutil
import subprocess

# Third Party Python Libraries.

# Self Built Python Libraries.


def getAppPath() -> str:
    """
    This function gets the path where the current code is been executed, this
    path will be an absolute path to the file in order to get files and saving
    elements.
    Returns:
        > script_directory (str): string containing the directory path absolute
        path.
    """
    # We get the absolute path where the program will be executed.
    if getattr(sys, "frozen", False):
        # If the program has been packed for distribution follows this path.
        script_directory = os.path.dirname(sys.executable)
    else:
        # If program is been executed as a python file.
        script_directory = os.path.dirname(os.path.abspath(sys.argv[0]))
        """
        When this is not main:
        script_directory = os.path.dirname(os.path.abspath(sys.argv[0])) 
        When this is main or we desire the module path:
        script_directory = os.path.dirname(os.path.abspath(__file__) 
        """
    # Returns the script directory.
    return script_directory


def getLibPath() -> str:
    """
    This function gets the library where the current code is been executed, this
    path will be an absolute path to the lib file in order to get files and
    saving elements.
    Returns:
        > script_directory (str): string containing the directory path absolute
        path.
    """
    # We get the absolute path where the program will be executed.
    if getattr(sys, "frozen", False):
        # If the program has been packed for distribution follows this path.
        script_directory = os.path.dirname(sys.executable)
    else:
        # If program is been executed as a python file.
        script_directory = os.path.dirname(os.path.abspath(__file__))
        """
        When this is not main:
        script_directory = os.path.dirname(os.path.abspath(sys.argv[0])) 
        When this is main or we desire the module path:
        script_directory = os.path.dirname(os.path.abspath(__file__) 
        """
    # Returns the script directory.
    return script_directory


def getOS() -> str:
    """
    Function gets the operative system where the application is running in order
    to run some system
    dependant functions.
    Returns:
    > os_type (str): Returns OS type as Linux, Windows, Darwin or Java.
    """
    # Gets the operative system platform where the program is running.
    os_type = platform.system()
    if os_type == "Linux":
        return os_type
    elif os_type == "Windows":
        return os_type
    elif os_type == "Darwin":
        return "MacOS"
    elif os_type == "Java":
        return os_type
    else:
        return "Unknown OS"


def setHighPriorityApp(osType: str) -> None:
    """
    Sets this program as high priority for the OS.

        Args:
            osType (str): operative system as 'Linux', 'Windows', 'MacOs' and
            'Java'
    """
    processID = os.getpid()  # Obtains process ID.
    process = psutil.Process(processID)  # Obtains process.

    # We set the high priority for best performance.
    if osType == "Windows":
        process.nice(psutil.HIGH_PRIORITY_CLASS)
    elif osType == "MacOS":
        process.nice(-20)
    elif osType == "Linux":
        process.nice(-20)
    else:
        print("OS not supported for high priority performance.")
    return None


def listFilePathsExt(absolutePath: str) -> list[str]:
    """
    Lists all file paths with extensions in the given absolute directory path.
    Args:
        absPath (str): Absolute path to the directory.
    Returns:
        list[str]: List of file paths with extensions.
    """
    return [
        # Joins the Absolute path with the file extension.
        os.path.join(absolutePath, file)
        for file in os.listdir(absolutePath)
        if os.path.isfile(os.path.join(absolutePath, file))
    ]


def listFileNames(absolutePath: str) -> list[str]:
    """
    Lists all file names without extensions in the given absolute directory path.
    Args:
        absolutePath (str): Absolute path to the directory.
    Returns:
        list[str]: List of file names without extensions.
    """
    return [
        os.path.splitext(file)[
            0
        ]  # Splits the file extension and returns the file name only
        for file in os.listdir(absolutePath)
        if os.path.isfile(os.path.join(absolutePath, file))
    ]


def listFileNamesWithExt(absolutePath: str) -> list[str]:
    """
    Lists all file names with extensions in the given absolute directory path.
    Args:
        absolutePath (str): Absolute path to the directory.
    Returns:
        list[str]: List of file names with extensions.
    """
    return [
        file  # Returns the file name with extension
        for file in os.listdir(absolutePath)
        if os.path.isfile(os.path.join(absolutePath, file))
    ]


def listFolderPaths(absolutePath: str) -> list[str]:
    """
    Lists all folder paths in the given absolute directory path.
    Args:
        absPath (str): Absolute path to the directory.
    Returns:
        list[str]: List of folder paths.
    """
    return [
        os.path.join(absolutePath, folder)
        for folder in os.listdir(absolutePath)
        if os.path.isdir(os.path.join(absolutePath, folder))
    ]


def listFolderNames(absolutePath: str) -> list[str]:
    """
    Lists all folder names in the given absolute directory path.
    Args:
        absolutePath (str): Absolute path to the directory.
    Returns:
        list[str]: List of folder names.
    """
    return [
        folder  # Returns the folder name
        for folder in os.listdir(absolutePath)
        if os.path.isdir(os.path.join(absolutePath, folder))
    ]


def openDirectory(path: str) -> None:
    """
    Opens the specified directory in the file explorer.

    Args:
        path (str): The absolute path to the directory.
    """
    system = getOS()
    if system == "Windows":
        os.startfile(path)
    elif system == "MacOs":  # macOS
        subprocess.run(["open", path])
    elif system == "Linux":
        subprocess.run(["xdg-open", path])
    return None
