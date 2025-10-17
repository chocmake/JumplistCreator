## Customization

General section, line and comments follow INI syntax. Comments can be added by adding a `;` to the start of a line, disabling that line from being included in the jumplist.

### Jumplist

- The jumplist configuration read by Jumplist Creator is `Jumplist.ini` and located beside the EXE.

- For category headings add a line that begins and ends with paired square brackets, like `[Example category]`.
	
- Items use `Name=Value` syntax. At the minimum the only necessary value for each jumplist item is a file/program/directory/URL target value, such as `Example program=C:\Example\app.exe`.

- Emojis and Unicode are supported for values and category headings. Examples:
	`[✨ Newest project]`
	`ラーメン=https://en.wikipedia.org/wiki/Ramen`

- Values don't require double quotes for paths but will handle them if used.

- Target/icon values will be searched both in the current directory, Windows environment variables and the absolute path.
	- Which means filenames are also supported if they exist in the Jumplist Creator directory or Windows PATH (eg: `Notepad=notepad.exe`).
		- One caveat: if using only filenames make sure to add the file extension since extensionless lookups aren't supported. So eg: `notepad.exe` will work but not `notepad`.

- Item values support optional additional parts:

	- Command-line arguments can be added by appending `| args:` followed by the arguments. Eg:
		`Example program=C:\Example\app.exe | args: --input="something"`

	- An icon can be defined to override the item's default icon by appending `| icon:` followed by the icon path and optional icon index (index only supported if it's an EXE/DLL). Examples:
		`Example program=C:\Example\app.exe | icon: C:\Icons\Tool.ico`
		`Example program=C:\Example\app.exe | icon: C:\Some tool\tool.exe`
		`Example program=C:\Example\app.exe | icon: shell32.dll,43`

	- A starting directory can be defined by appending `| startin:` followed by the path of the directory you want to the program to treat as its working directory. Eg:
		`Example program=C:\Example\app.exe | startin: D:\My documents`
		
	- All the above additional parts can be combined in any order, they only have to placed on the same line after the base target path. Eg:
		`Example program=C:\Example\app.exe | args: --input="something" | icon: C:\Icons\Tool.ico`
		
### Settings

Most settings can be changed via the GUI. However a few things are only definable via the `Settings.ini`, located beside the program, which are listed here.

- `DefaultLaunchAction`. This can be set via the GUI but only when pointed to a file/program. If adjusted via the INI you can make it point to a directory or URL, or add program arguments (see the `args` info in the jumplist section above).

- `DefaultLaunchActionDrop`. This supports the same syntax as `DefaultLaunchAction` but allows defining an action for when files/directories are drag-and-dropped onto Jumplist Creator.

	- Two placeholders are supported, which will get swapped at runtime for the actual path(s) dropped:

		- `%1` represents a *single* path dropped onto the program. Path will be wrapped in double quotes.

		- `%*` represents *all* paths dropped onto the program. Each path is wrapped in double quotes and separated by spaces.

	- So if the program you're using only supports a single path for arguments you can use `%1`, while if it supports space-separated multiple paths you can use `%*`.

	- You can define an entirely separate program from the program used for `DefaultLaunchAction` if wanted. Eg:
		`DefaultLaunchActionDrop=mspaint.exe | args: %1`

	- Or if you want to use the same program you can write the same target value or use the literal placeholder string `<default>`. Eg:
		`DefaultLaunchActionDrop=<default> | args: --input %1 --example "some other value"`

- `MetaItemsVisibility`. Supports one INI-only extra value: `0`. This will disable the program adding any meta jumplist items (the 'Configuration' and optional 'Update' items) to jumplists.
	
	- This is useful if you want to manually add your own `[Jumplist]` meta category to the `Jumplist.ini` as the very last section, with custom items and icons.
	
		- The `[Jumplist]` section, when it occurs as the very last section in the `Jumplist.ini`, is treated specially by the program as it will always make sure it's visible regardless of the maximum jumplist items value (ie: it will truncate non-meta items to make room for the meta items).
		
		- For referencing the `JumplistCreator.exe` you can use the special literal placeholder string `<JumplistExe>` (case insensitive) for the target path and it will be replaced by the current Jumplist Creatore EXE path in the jumplist. Eg:
			
			```
			[Jumplist]
			My custom name=<JumplistExe> | args: --config | icon: D:\Icons\example.ico
			```
	
	- Just be mindful if you set this option and then update the jumplist without any custom meta jumplist items and you also have a `DefaultLaunchAction` set you won't be able to open the Jumplist Creator GUI without using the command-line argument `JumplistCreator.exe --config`.
