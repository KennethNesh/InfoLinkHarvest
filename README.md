InfoLinkHarvest:
InfoLinkHarvest is a Python application built with tkinter and Selenium, designed to automate the process of searching for specific terms on Google and extracting the search result links into a Word document.

Features:
    Google Search Automation: Utilizes Selenium to perform automated Google searches based on user-provided name and specific search terms.

    Link Extraction: Extracts search result links from Google search pages.

    Document Generation: Saves the extracted links into a Word document (.docx) for easy access and reference.

    User-Friendly GUI: Built using tkinter, providing a simple and intuitive graphical interface for users to input search parameters and initiate the search process.

    Error Handling: Includes basic error handling to notify users of any encountered issues during the search and save process.
Requirements:
    Python 3.x
    tkinter
    Selenium
    Chrome WebDriver
How to Use:
    Clone or download the repository to your local machine.
    Open a terminal or command prompt and navigate to the project directory.
    Install the required dependencies by running:
    Copy code
    pip install -r requirements.txt
    Run the application by executing:
    Copy code
python infolinkharvest.py
Enter the name to search and specific search term in the respective input fields.
Click on the "Search and Save" button to initiate the search process.
Choose a folder to save the generated Word document containing the search result links.
Converting to Executable (EXE) File
You can use PyInstaller to convert the Python script into a standalone executable file for easier distribution. Follow these steps to do so:

Install PyInstaller:
    pip install pyinstaller
Navigate to the project directory in your terminal.
Run PyInstaller with the following command:
    pyinstaller --onefile infolinkharvest.py
This command will create a standalone executable file named infolinkharvest.exe in the dist directory.
You can now distribute the infolinkharvest.exe file to users who do not have Python installed.
Note
Ensure that Chrome WebDriver is installed and configured properly.
This application relies on internet connectivity for performing Google searches.