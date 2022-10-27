# My-Office-Add-in
[Build an Excel task pane add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator) using Office Add-in.

---
## Prerequisites
For more information, please refer to the following link: <br/>
https://learn.microsoft.com/en-us/office/dev/add-ins/overview/set-up-your-dev-environment

### Node.js
This guide assumes that you are using Ubuntu 20.04.

Please run the following commands to install **Node.js**:
```powershell
$ curl -sL https://deb.nodesource.com/setup_18.x -o /tmp/nodesource_setup.sh
$ sudo bash /tmp/nodesource_setup.sh
$ sudo apt-get install -y nodejs
```

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
$ sudo npm install -g yo generator-office
```

### Install and use the Office JavaScript linter
Microsoft provides a JavaScript linter to help you catch common errors when using the Office JavaScript library. To install the linter, run the following two commands:
```powershell
$ sudo npm install office-addin-lint --save-dev
$ sudo npm install eslint-plugin-office-addins --save-dev
```

Run the linter with the following command either in the terminal of an editor, such as Visual Studio Code, or in a command prompt.
```powershell
$ npm run lint
```

