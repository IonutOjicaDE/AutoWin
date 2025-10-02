# AutoWin
Windows Automation using Excel + VBA

## 1. Project Overview

**AutoWin** is an open-source automation framework built entirely with **Excel + VBA**, designed to simplify repetitive Windows tasks without requiring additional software.
Because almost every Windows computer already includes Excel with VBA, AutoWin works *out of the box*, fully offline, without dependencies or installers.

The project was originally created for personal use and later opened to the public as a reusable framework. Its main value lies in making Windows automation accessible through a familiar tool (Excel), while remaining easy to extend for beginner programmers who want to add their own commands.

Typical use cases include:

* Filling in repetitive forms automatically.
* Replacing text across all Word files in a folder (and subfolders).
* Inserting structured data into Word documents.
* Moving and resizing windows based on screen setup.
* Recording actions to generate new automations directly from user activity.

**Technical context**: AutoWin runs exclusively on Windows, Excel, and VBA7 (Win64). It supports both Excel 32-bit and 64-bit installations.

---

## 2. Features

AutoWin comes with a comprehensive set of features for Windows and Office automation:

* **Mouse control**: move, click, drag operations.
* **Keyboard automation**: send keystrokes, hotkeys, and special keys.
* **Loops & conditions**: `GoToLabel`, `GoSub`, `For`, `ForEach`, and `Do...Loop` with `While` and `Until` - just like in VBA.
* **Word integration**: insert text, find & replace.
* **Outlook integration**: send emails, open existing items.
* **Excel integration**: read/write cells, evaluate formulas, copy to clipboard.
* **Window management**: search windows by title/class, get position & state, activate, move, minimize/maximize.
* **System utilities**:

  * Run any program with arguments.
  * Read/change screen resolution.
  * Detect pixel colors (cursor position or arbitrary x,y) and use them in conditions.
  * Stop execution automatically if the mouse is moved by the user.
* **Logging & error handling**: write execution logs and handle exceptions gracefully.
* **Macro recording**: capture actions and replay them as automation scripts.

**Extensibility**

* All commands are centrally registered in `RegisterCommandsCondition`.
* Adding a new command requires only mapping it in `commandMap` and implementing a VBA function that returns `True` on success or `False` on error.
* Example: the `Skip` command demonstrates how new commands can be added with minimal code.

?? **User interface**

* Press `Ctrl+Shift+N` to open the **Command Picker** (`ufCommand`), which lists commands by category, explains their arguments, and inserts them into the automation sheet.

* Press `Ctrl+Shift+M` to open the **Execute Macro** (`ufAutoWin`), which lists Subs by their name and let you choose which to execute.

**Stability**
All current features are considered stable. File operations (copy, move, delete, read) are not yet implemented but can be easily added with the same framework.

---

## 3. Project Structure

* **Main file**:
  The core of AutoWin is a single workbook: **`AutoWin.xlsm`**. This file contains both the automation engine (VBA code) and the sheets where automations are defined.

* **Worksheets**:

  * **Automation** - central sheet where all automations are defined.

    * **A - Status**: execution status of each command.
    * **B - Command**: command name.
    * **C-L - Arg1…Arg10**: up to 10 arguments per command.
    * **M - Window**: window name checked before execution.
    * **N - Color**: pixel color condition.
    * **O - Pause**: delay before execution.
    * **P - Keyboard Code**: keyboard input code.
    * **Q - OnError**: (not implemented yet).
    * **R+ - Comment**: comments, not used in execution.

  * **KeyPress** - reference sheet for keyboard codes.

    * **A - Key Name**
    * **B - Key Code (Hex)**
    * **C - Key Code (Dec)**
    * **D - Description**
    * **E - SendKeys equivalent**
    * **F - Character**
    * **G - Pressed flag**

* **Source code (`/src`)**:
  Contains exported VBA modules, useful for version control and development. End users do **not** need these files.

* **Log file**:
  When errors occur, AutoWin creates a log file named after the workbook (e.g. `AutoWin.log`).

---

## 4. Installation

1. **Download**

   * Get `AutoWin.xlsm` directly from this repository.

2. **Enable macros**

   * When opening the file in Excel, enable macros to allow automation to run.
   * The *"Trust access to VBA project"* option is only required for developers exporting VBA code, not for regular users.

3. **Compatibility**

   * Windows with **VBA7 (Win64)**.
   * Works on both Excel 32-bit and 64-bit.

4. **Initial setup**

   * No configuration is required.
   * Automations can be entered directly in the `Automation` sheet.

5. **Macro security note**

   * Only enable macros for files you trust.
   * AutoWin is open source, so you can inspect all VBA code before running.

---

## 5. Usage Guide

### Step-by-step: First automation (“Hello World” with Notepad)

1. Open **AutoWin.xlsm**.
2. Go to the **Automation** sheet.
3. Add a new automation block:

| St | Command                 | Arg1         | Arg2 | Arg3 | WindowName | Pause | Comment        |
| -- | ----------------------- | ------------ | ---- | ---- | ---------- | ----- | -------------- |
|    | Sub                     | Notepad      |      |      |            |       | Start block    |
|    | Run Program             | notepad      |      |      |            | 500   | Launch Notepad |
|    | Activate window by name | Notepad      |      |      |            | 200   | Focus window   |
|    | Send Keys               | Hello World! |      |      | Notepad    | 200   | Type text      |
|    | End Sub                 |              |      |      |            |       | End block      |

4. Run the automation:

   * Press `Ctrl+Shift+N` to open the **Command Picker** (`ufCommand`).
   * Select your automation and execute it.
   * Notepad will open, activate, and type *Hello World!* automatically.

---

### Understanding the **Automation** sheet

Each line represents a command with optional arguments.

* **St (Status)**

  * `>` pending execution
  * `+` executed successfully (`+1`, `+2`, ... in loops)
  * `!` error (automation stops)
  * `-` skipped (empty or invalid)
* **Command** - the instruction to run (e.g., *Run Program*, *Send Keys*).
* **Arg1-Arg10** - parameters depending on the command.
* **WindowName** - target window title (partial or full). If missing, AutoWin waits up to 5 seconds (configurable).
* **ColorUnderMouse** - hex color code checked before execution. AutoWin waits up to 5 seconds if mismatch (tolerance configurable).
* **Pause** - delay before execution (default: 50 ms). Can be overridden globally.
* **KeybdCode** - keyboard layout code to switch before typing (useful for multi-language input).
* **Comment** - free text, ignored at runtime.

---

### Recommended screenshots

* Full view of **Automation** sheet with a simple automation.
* **KeyPress** sheet to show keyboard mappings.
* **Command Picker (ufCommand)** form, to illustrate how commands can be inserted.
* Optionally: execution log file (`AutoWin.log`) after a successful run.

---

## 6. Examples

### Example 1: Resize Notepad window

```
Sub  Notepad Resize
Get window name from class      | Arg2=Notepad
Activate window by name         | Arg1=Notepad  | Pause=500
Get Resolution                  | Arg1=1920 Arg2=1200 | Pause=500
Set window position             | Arg1=Notepad Arg2=15 Arg3=15 Arg4=1890 Arg5=1140 | Pause=500
End Sub
```

? Opens Notepad, activates it, and resizes the window to nearly full screen.

---

### Example 2: Maximize Notepad in a loop

```
Sub  Maximize Notepad
For   Arg3=1 Arg4=10
Activate window by name | Arg1=Notepad
Window Maximize         | Pause=500
Get Window State        | Arg2=maximized | Pause=300
If Then Skip            | Arg1=FALSCH Arg2=1 | Pause=200
Next
End Sub
```

? Loops through 10 iterations, ensuring Notepad is maximized.

---

### Example 3: Taskbar menu manipulation

```
Sub  Hide Search Symbol on Taskbar
Right click   | Arg1=1850 Arg2=1180   | Comment=Taskbar menu
Left click    | Arg1=1850 Arg2=820    | Comment=Search option
Left click    | Arg1=1400 Arg2=820    | Comment=Hide option
End Sub
```

? Opens the Windows taskbar menu and hides the Search icon automatically.

---

## 7. Roadmap / To-Do

AutoWin is already stable for everyday automation, but several enhancements are planned:

* **Immediate priorities**

  * Extend **Command Picker (ufCommand)** to configure not only `Command` and `Args`, but also `WindowName`, `ColorUnderMouse`, `Pause`, `KeybdCode`, `On Error`, and `Comment`.
  * Implement full support for **On Error** handling.

* **Planned features**

  * Improve window identification (e.g. matching by class).
  * Support for addressing **controls inside windows** (buttons, textboxes, etc.).
  * Back-end actions using **SendInput** directly on objects.

* **Community input**
  Suggestions and feature requests are highly welcome. Please open an **Issue** on GitHub to share ideas or propose improvements.

---

## 8. Contributing

Contributions are welcome, whether in the form of bug reports, feature requests, or code.

* **Workflow**

  * Report bugs or suggest features in **GitHub Issues**.
  * Developers can fork the repository, create a feature branch, and submit a Pull Request.

* **Coding style**

  * All **comments and variable names in English**.
  * Follow the style used in the existing codebase (logical module grouping, descriptive function names, structured error handling).

* **Examples**

  * If you add a new command, it is ideal (but optional) to provide a sample automation in the **Automation** sheet demonstrating its usage.

* **Communication**

  * All discussions take place via GitHub Issues.

* **Acknowledgments**
  AutoWin builds on knowledge and inspiration from the VBA community. Key resources include:

  * Stack Overflow, MSDN, Microsoft Docs, VBForums, MrExcel, Rene Nyffenegger’s API reference, vbarchiv.net, CodeProject, and many others (see full list in source code comments).

---

If my work has been useful to you, do not hesitate to offer me a strawberry milk

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/ionutojica)
