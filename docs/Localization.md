## Localization

All default UI strings beside the program name are stored in `Languages\en-US.txt` within the source code. Translations only need to create a single file. To make and test a translation:

- Create a `Languages` directory beside the `JumplistCreator.exe`, if there isn't one.

- In it create a new TXT file named after the [locale code](https://simplelocalize.io/data/locales/) you're localizing for. Eg: `es.txt` for Spanish or for a specific dialect can use the full code like `es-AR.txt`.

- Use the same key names as the `en-US.txt` reference but with different values.

	- The name of the language itself is the `LangDisplayName` value. This is what will appear in the Settings UI menu. If that key-value pair isn't defined the language filename will be used instead.
	
	- The values of `MetaJumplistItemConfig` and `MetaJumplistItemUpdate` are used for the configuration item names which appear in the jumplists themselves.

	- Any keys you don't add to your localized version will fall back to its default English string.

- Test in-program by opening Jumplist Creator's settings UI and selecting the language from the *Languages* drop-down list. All language files added to the `Languages` sibling directory will appear here.

> If the user has their system locale set to one dialect but the only language available for the program is another dialect Jumplist Creator will find the common root, like `es` and use that by default (unless changed via the settings).

### Contributing

If you feel like contributing a translation to the project open a pull request with your translated file added to the source code's `Languages` directory, or just open an [issue](https://github.com/chocmake/JumplistCreator/issues) and attach the file.