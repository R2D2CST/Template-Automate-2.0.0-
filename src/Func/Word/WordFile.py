# Native Python Libraries.

# Third Party Python Libraries.
from docx import Document
from docxtpl import DocxTemplate

# Self Built Python Libraries.


class wordDocument:
    """
    Class with methods for Excel Handling
    Att:
        wordDocumentPath (str): Word document absolute path including its
        file extension.
    Meth:
        . builtTemplate (None)->None:
            builds a template example with instructions for the user ro
            render model.
    """

    def __init__(self, wordDocumentPath: str) -> None:
        """ """
        self.wordDocumentPath = wordDocumentPath
        pass

    def buildTemplate(self) -> None:
        """
        Procedure creates a Word document with instructions for creating templates to automate
        with an Excel database.

        Args:
            None
        Returns:
            None
        """

        # Create a new Word document
        wordFile = Document()

        # Add content to the document

        content = [
            "Welcome, this is a sample template to automate your reports or documents using keywords as shown in the following example:{{keyWord1}}",
            "",
            "You will be able to replace these keywords by populating the content in the database.",
            "",
            "In the Excel document, you should fill in the database (database.xlsx) in row 1 with the column keywords the program will read all cells as keywords in row 1 except by A1.",
            "These should be recorded without the double curly brackets (Important: do not include the double curly brackets).",
            "The Column 1 database is exclusive for the run directory name paths.",
            "",
            "Feel free to run the program with the template and the database to see how the system works.",
            "Examples for the keyword -> keyword",
            "",
            "Below these keywords, the program will read each row as the value associated with that keyword.",
            "",
            "keyword1\t:\t{{keyWord1}}",
            "keyword2\t:\t{{keyWord2}}",
            "keyword3\t:\t{{keyWord3}}",
            "keyword4\t:\t{{keyWord4}}",
        ]

        for paragraph in content:
            wordFile.add_paragraph(paragraph)

        # Save the document in the Templates folder
        wordFile.save(self.wordDocumentPath)
        return None

    @staticmethod
    def renderTemplate(templatePath: str, outputPath: str, context: dict) -> None:
        """
        Renders a Word document template using the provided context and saves it to the output directory.

        Args:
            template_path (str): The path to the template file.
            output_path (str): The path where the rendered document will be saved.
            context (dict): The context dictionary containing data to render the template.

        Returns:
            None
        """
        # Load the template
        document = DocxTemplate(templatePath)

        # Render the document using the provided context
        document.render(context)

        # Save the rendered document to the specified path
        document.save(outputPath)
        return None

    pass
