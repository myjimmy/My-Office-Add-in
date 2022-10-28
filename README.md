# My-Office-Add-in
[Build an Excel task pane add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator) using Office Add-in.

---
## Prerequisites
For more information, please refer to the following link: <br/>
https://learn.microsoft.com/en-us/office/dev/add-ins/overview/set-up-your-dev-environment

### Node.js
Go to the [Node.js website](https://nodejs.org/en/) and download the latest recommended version.

**npm** usually installed automatically when you install Node.js.

Please run the following commands to check the version of **Node.js** and **npm**:
```powershell
$ node -v
v18.12.0

$ npm -v
8.19.2
```

### Install a code editor
You can use any code editor or IDE that supports client-side development to build your web part, such as:

- [Visual Studio Code](https://code.visualstudio.com/) (recommended)
- [Atom](https://atom.io/)
- [Webstorm](https://www.jetbrains.com/webstorm)

### Install the Yeoman generator — Yo Office
The project creation and scaffolding tool is [Yeoman generator for Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview), affectionately known as Yo Office. You need to install the latest version of [Yeoman](https://github.com/yeoman/yo) and Yo Office. To install these tools globally, run the following command via the command prompt.
```powershell
$ npm install -g yo generator-office
```

### Install and use the Office JavaScript linter
Microsoft provides a JavaScript linter to help you catch common errors when using the Office JavaScript library. To install the linter, run the following two commands:
```powershell
$ npm install office-addin-lint --save-dev
$ npm install eslint-plugin-office-addins --save-dev
```

Run the linter with the following command either in the terminal of an editor, such as Visual Studio Code, or in a command prompt.
```powershell
$ npm run lint
```

## Create the add-in project
For more information, please refer to the following link: <br/>
https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator#create-the-add-in-project

Run the following command to create an add-in project using the **Yeoman generator**.
```powershell
$ yo office
```

When prompted, provide the following information to create your add-in project.

* **Choose a project type:** `Office Add-in Task Pane project`
* **Choose a script type:** `Javascript`
* **What do you want to name your add-in?** `My Office Add-in`
* **Which Office client application would you like to support?** `Excel`

After you complete the wizard, the generator creates the project and installs supporting Node components.

## Explore the project
The add-in project that you've created with the Yeoman generator contains sample code for a basic task pane add-in. If you'd like to explore the components of your add-in project, open the project in your code editor and review the files listed below. When you're ready to try out your add-in, proceed to the next section.
* The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.<br/>To learn more about the manifest.xml file, see [Office Add-ins XML manifest](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests).
* The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.
* The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.
* The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.
