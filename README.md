# Python2Doc
Python script to automate Illustrator and word to extract text from illustrator files to a Word file for translation and reimporting translated text back into an illustartor file.

## How to use
1. Download [Setup.exe](https://github.com/ChonkyYoshi/Python-.ai2Doc/releases/download/v.0.1/Setup.exe) and run it to install the script on your machine
2. Once installed make sure both Adobe Illustrator and Microsoft Word are installed and activated on the machine
3. Run the program from the dekstop shortcut or the Start Menu and follow the instructions.

## Restrictions
- This script is simply an automation script like VBA, as such it *will not* work if both Illustrator and Word are not installed on the machine.
- Only Text strings (i.e. text in a TextFrame) are extracted, anything else is ignored.
- The order in which the text strings are extracted and put in the resuling Word file is based on their place in the layers list. The script will go from top to bottom so you can rearrange the layers to have them be extracted in the order you want.
- The script *does not* check if the content in the Word file is ordered as it should or if it matches what's in the file already. It assumes the source file has not changed since when the text strings have been extracted and it assumes the Word file structure is still the same. It only checks that the file name matches.

# Feature requests/issues
Please open a Github issue requesting with the feature you're requesting or explaining your issue in as much detail as you can.
If you don't have a GitHub account, you can send an email to [agosta.enzowork@gmail.com](mailto:agosta.enzowork@gmail.com) doing the same.

Feel free to also check my other project [Prep Toolkit](https://github.com/ChonkyYoshi/Prep-ToolKit) if you want more scripts/automation tools regarding translation.
