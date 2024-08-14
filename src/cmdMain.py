# Native Python Libraries.
import threading

# Third Party Python Libraries.

# Self Built Python Libraries.
from Func import Func


# We break our terminal based user interface in
def header() -> None:
    """
    Function prints the header for the application
    Args: None
    Returns: None
    Raises: None
    """
    print("________________________________________________")
    print("|  TEMPLATE AUTOMATE PROJECT MANGER            |")
    print("|  VERSION: 2.0.0                              |")
    print("|  AUTHOR: José Arturo Castella Lasaga         |")
    print("|  SUPPORT: qfbarturocastella@gmail.com        |")
    print("________________________________________________")
    print()
    pass


def askAndCheck() -> int:
    """
    Function asks the to choose a selection and checks if selection is an
    integer number. Otherwise it returns a default number 0 as an error escape
    situation.
    Args: None
    Returns: selection (int): User selection as an integer.
    Raises: selection (int): Returns '0' as a value as escape sequence.
    """
    try:
        selection = int(input("Please enter your selection and press enter: "))
    except Exception as e:
        print(f"Error: {e}")
        print("Please enter Round Numbers")
        selection = 0
    return selection


def pause() -> None:
    """
    Pauses the program with a input 'Press enter to continue...'
    Args: None
    Returns: selection (int): User selection as an integer.
    Raises: selection (int): Returns '0' as a value as escape sequence.
    """
    input("Press enter to continue...")
    pass


def mainMenuOption() -> None:
    """
    Prints main menu selection options
    Args: None
    Returns: None
    Raises: None
    """
    print("1.- Exit program.")
    print("2.- About the program.")
    print("3.- Create a new project.")
    print("4.- Work on a existing project.")
    print("5.- Set high performance rendering.")
    print("6.- Activate monitoring performance.")
    pass


def goodbye() -> None:
    """
    Says goodbye to the user
    Args: None
    Returns: None
    Raises: None
    """
    print("See you until next time.")
    pause()
    pass


def listProjects() -> None:
    projectsNameList, projectsPathList = Func.enumerateProjectsInside()
    del projectsPathList
    try:
        for i in range(len(projectsNameList)):
            # Best user experience if counting starts in 1.
            print(f"{i+1}.- {projectsNameList[i]}.")
    except IndexError:
        pause()
    pass


def projectSelection():
    """
    Gets the user selection for the project and its project path.

    Args: None

    Returns:
        projectName (str): name of the selected project.
        projectPath (str): path of the selected project.

    Raises: None
    """
    selection = askAndCheck()
    if selection == 0:
        pass
    else:
        selection -= 1
    projectsNameList, projectsPathList = Func.enumerateProjectsInside()
    try:
        return projectsNameList[selection], projectsPathList[selection]
    except IndexError:
        return "", ""


def listProjectActions():
    """
    Prints main menu selection options
    Args: None
    Returns: None
    Raises: None
    """
    print("1.- Return to Main Menu")
    print("2.- Open project directory path.")
    print("3.- Render word templates in project.")
    print("4.- Render excel templates in project.")
    print("5.- Render both excel and word templates in project.")
    pass


def projectOptionMenu() -> None:
    """
    Displays the actions user can take in a project.
    Args: None
    Returns: None
    Raises: None
    """
    while True:
        listProjects()
        projectName, projectPath = projectSelection()
        if projectName == "":
            continue
        else:
            break
    while True:
        print("Selected Project: " + projectName)
        listProjectActions()
        selection = askAndCheck()
        if selection == 0:
            continue
        elif selection == 1:
            break
        elif selection == 2:
            Func.openDirectory(path=projectPath)
            print("Project Path Open.")
            pause()
            continue
        elif selection == 3:
            Func.renderWordTemplates(projectPath=projectPath)
            continue
        elif selection == 4:
            Func.renderExcelTemplates(projectPath=projectPath)
            continue
        elif selection == 5:
            Func.renderAllTemplates(projectPath=projectPath)
            continue
    pass


def mainLoop() -> None:
    """
    Main loop for handling user experience
    Args: None
    Returns: None
    Raises: None
    """
    header()
    while True:
        mainMenuOption()
        selection: int = askAndCheck()
        if selection == 0:
            pause()
            continue
        elif selection == 1:
            goodbye()
            break
        elif selection == 2:
            header()
            continue
        elif selection == 3:
            projectName = input("Please enter the new project name: ")
            projectName = projectName.lower()
            projectName = projectName.capitalize()
            Func.createProject(projectName=projectName)
            continue
        elif selection == 4:
            projectOptionMenu()
            continue
        elif selection == 5:
            Func.setHighPerformance()
            print("High performance activated.")
            pause()
            continue
        elif selection == 6:
            monitorThread = threading.Thread(target=Func.startMonitoringPerformance)
            monitorThread.start()
            print("Monitoring app performance.")
            pause()
            continue
    pass


if __name__ == "__main__":
    mainLoop()
    exit()
