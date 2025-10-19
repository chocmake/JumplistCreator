## Jumplist Creator

Jumplist Creator is a program for generating custom native Windows jumplist menus, as a way of keeping applications, directories and files organized and launchable from the taskbar via pinned items.

Such menus have been long supported since Windows 7 but are typically only seen for recent file lists or tasks specific to certain programs. However this useful Windows feature can also be used to allow for arbitrary, user-configured menus.

![Image](https://github.com/user-attachments/assets/8cc02391-3276-4b50-a0f6-7013f3748e32)

## Features

- Jumplist items can be a program/file, directory or URL.
- Items support optional arguments, custom icons and starting directories.
- Category headings can be made for visual grouping.
- Can be configured to launch a different program/file/directory when opened, so its pinned taskbar item can serve two purposes.
	- Also support for optional drag-and-drop of files/directories onto its taskbar item¹ to forward to a specified program.
- Portable² and relatively small (~1MB).
- Plaintext (INI-based) configuration files.
- Dark mode and high DPI support.
- Supports Windows 10 and 11.
- Command-line arguments `--config` and `--update` can be used to launch the configuration GUI or update the jumplist.

> [1] This drag-and-drop feature is only supported on W10 (or older) taskbars or on W11 with replacement taskbars like StartAllBack, since Windows 11 hasn't yet re-implemented drag-and-drop file support for taskbar items with its new taskbar. Windows also requires the program be in a closed state else drag-and-drop won't trigger.

> [2] Notes on portability: necessarily writes a native maximum jumplist items value to the registry and tells Windows to create jumplists (which get stored by Windows itself as `customDestinations-ms` files). Otherwise only writes to its own program directory.

---

## Screenshots

<p align="center" width="100%">
    <img src="https://github.com/user-attachments/assets/7b87cac9-b1ce-4c61-a8ca-47bf13339cb4"/>
</p>

---

## Setup

1. Download the zip of the latest release [here](https://github.com/chocmake/JumplistCreator/releases/latest/download/JumplistCreator.zip).
2. Unzip to some writable location. Eg: `C:\Jumplist Creator\Example`.
3. Drag `JumplistCreator.exe` onto your Windows taskbar to pin it.
	- Or alternatively open it then right-click its taskbar item and select *Pin to taskbar*.
4. Launch `JumplistCreator.exe` and click the *Generate List* button.
5. This will create a new `Jumplist.ini` and automatically open it in your default text editor.
6. Adjust the example jumplist menu categories and items with your own [customization](https://github.com/chocmake/JumplistCreator/blob/main/docs/Customization.md), then save the file and press the *Update List* button.
7. View your jumplist menu by right-clicking the pinned Jumplist Creator taskbar item.
	- Or alternatively, if using an older Windows version or StartAllBack's classic taskbar, can drag outward from its taskbar item.

https://github.com/user-attachments/assets/2dba201b-971d-4e83-8019-3e2f3b35333a

> **Tip:** to create multiple jumplist menus unzip to other locations and then pin those versions of the EXE to the taskbar. Eg: could have one copy in `C:\Jumplist Creator\Tools` and another in `C:\Jumplist Creator\Games`, each with separate configurations.

> Any sibling `Languages` directory is optional and only used for non-default localizations (if present). Also the `Theme.dll` is optional for the dark theme, it can be removed if wanted and the program will default to light theme.

> The program uses .NET Framework version 4.8, which is installed by default on Windows 10 and 11. If using an older Windows version you may need to install support for it.

---

## Videos

**Multiple jumplist taskbar items** with custom taskbar icons and names:

https://github.com/user-attachments/assets/2c0a1429-8106-45de-a263-ddbbb35b5438

> To achieve this create a Windows shortcut to that copy of `JumplistCreator.exe` and in the shortcut properties choose a different icon and name. Drag that shortcut onto the taskbar to pin it, then update the jumplist.

> Seen with Fluent UI Emoji icons.

**How a different program can be launched** on left-click and its taskbar items grouped in the one taskbar item:

https://github.com/user-attachments/assets/29b193b4-d9df-47c5-9863-f30e10c093a3

> Using StartAllBack's W11 taskbar which supports file drag-and-drop (and vertical alignment), Jumplist Creator here was set to launch MPC-HC Portable on click and also to accept files dropped onto the Jumplist Creator taskbar item, forwarding the file path to MPC-HC for opening.

> How the grouped taskbar windows was achieved was by modifying the Windows shortcut (LNK) file used to pin Jumplist Creator to the taskbar, so its AppId matched that of MPC-HC. This made Windows consider that pinned item as being from MPC-HC, allowing its windows to be grouped into one item and enabling the special taskbar effects like video progress colors to be applied.

> Seen with custom icons, in addition to videos from NHK's Creative Library.

---

## Documentation

- [Customization](https://github.com/chocmake/JumplistCreator/blob/main/docs/Customization.md)

- [Building](https://github.com/chocmake/JumplistCreator/blob/main/docs/Building.md)

- [Localization](https://github.com/chocmake/JumplistCreator/blob/main/docs/Localization.md)

---

## Credits

- NykUser's [Jump_List](https://github.com/NykUser/Jump_List) project for the initial basis of this hard fork ([link](https://github.com/NykUser/Jump_List/tree/ef2e97ec853aaa4c212be1a70818b945e7aadc6a) to specific repo version). For Jumplist Creator various jumplist item icon issues for W10/W11 were fixed and the functionality and UI expanded.

- [DarkModeForms](https://github.com/BlueMystical/Dark-Mode-Forms/) for the dark theme support. Added style [adjustments](https://github.com/chocmake/Dark-Mode-Forms) in a fork and compiled to DLL for the releases.

- [InsertIcons](https://github.com/einaregilsson/InsertIcons) for easily embedding multiple icon files into the EXE.
