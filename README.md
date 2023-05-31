This is electron Desktop Application for extracting results from VTU Results Website and generating a report in Excel format. 

To Run this Project you need to have a few things installed on your machine.

1. Node.js
2. NPM
3. Python 3.6
4. Pip

Python requires a few packages to be installed. You can install them by running the following command.

1. xlwings ( Excel Library )
2. chompjs ( JSON Parser ) (Convert JSON to Python Dictionary)
3. pyinstaller ( To build the python script into executable file )

Once you have all of these installed, you can run the following commands to get the project up and running.

1. npm install
2. npm run start

Note : For CSS styling I have used pico.css library. Also, we have used Custom Web Components for the UI.

Web Components 
The web components are written inside the view/components folder
(we can simply import these components in the html file and use them, Eg : app-header & app-footer)
More Web components can be added in the same way for the UI for future development.


Views Folder:
The views folder contains the html files for the UI.
The main html file is index.html which is the main file for the UI.
The other html files are the pages for the UI.

Each html has a script tag which contains the javascript code for the UI.
The script is embedded inside the html file for the sake of simplicity & easy maintenance.

Styles Folder:
The styles folder contains the css files for the UI.
We have used pico.css library for the styling of the UI.
Custom css files can be added & linked to the html for the UI for future development.


The main.js file is the entry point for the electron application.
The main js will call python child process to run the python script.
The python script has to be build using pyinstaller.
pyinstaller will executable file for the python script.
(


To Build Python Script using pyinstaller
Run the following command in the terminal
npm run buildPythonExecutable
This will create an executable file & overwrite existing canaraFetch.exe file in the python folder.


# How to build the project for production

Run the following command in the terminal
npm run package

This will create a folder named "out" in the project directory.
The out folder will contain the executable file for the application.

Imp Note : You have to copy python folder to the vtuResults2ExcelTool-win32-x64 folder manually.
The python folder contains the python script/exe & the Excel template file.

Imp Note : Dont add devDependencies while building project ( move all dev dependencies that are application critical to dependencies)







