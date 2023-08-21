from pathlib import Path
from win32com.client import DispatchEx
import PySimpleGUI as gui
from regex import search
from random import choice
from configparser import ConfigParser

# General note, there are a lot type: ignore and noqa E501 to shut up the formatter warnings  # noqa: E501

# Grab intructions from config.ini
config = ConfigParser()
config.read('config.ini')


# Error Handling
def FixAiFileList(AiFileList: list):
    for index, AiFile in enumerate(AiFileList):
        AiFile = Path(AiFile)
        if AiFile.is_dir():
            gui.popup_error(
                f'"{AiFile.parent.as_posix()}" {config["errors"]["No_Dir"]}',
                title='Folders are not supported')
            AiFileList.pop(index)
            continue
        if not AiFile.is_file():
            gui.popup_error(f'"{AiFile.name}" {config["errors"]["Not_file"]}',
                title='File doesn\'t exist')
            AiFileList.pop(index)
            continue
        if AiFile.suffix != '.ai':
            gui.popup_error(f'"{AiFile.name}" {config["errors"]["Not_ai"]}',
                title='File is not .ai')
            AiFileList.pop(index)
            continue
    return AiFileList


def FixWordFileList(WordFileList: list):
    for index, WordFile in enumerate(WordFileList):
        WordFile = Path(WordFile)
        if WordFile.is_dir():
            gui.popup_error(
                f'"{WordFile.parent.as_posix()}" {config["errors"]["No_Dir"]}',
                title='Folders are not supported')
            WordFileList.pop(index)
            continue
        if not WordFile.is_file():
            gui.popup_error(f'"{WordFile.name}" {config["errors"]["Not_file"]}',
                title='File doesn\'t exist')
            WordFileList.pop(index)
            continue
        if WordFile.suffix != '.docx':
            gui.popup_error(f'"{WordFile.name}" {config["errors"]["Not_docx"]}',
                title='File is not .docx')
            WordFileList.pop(index)
            continue
    return WordFileList


def SetFields(Option: str):

    window['-Start-'].update(visible=True)
    match Option:
        case 'Extract':
            # Display correct instructions and only the Illustrator file field.
            window['-Info-'].update(value=config['Instructions']['Extract'])
            window['-PBar-'].update(visible=True)
            window['-AiPath-'].update(visible=True)
            window['-AiBrowse-'].update(visible=True)
            window['-DocPath-'].update(visible=False)
            window['-DocBrowse-'].update(visible=False)
        case 'Import':
            # Display correct instructions and both the Illustrator and Word file field.
            window['-Info-'].update(value=config['Instructions']['Import'])
            window['-PBar-'].update(visible=True)
            window['-AiPath-'].update(visible=True)
            window['-AiBrowse-'].update(visible=True)
            window['-DocPath-'].update(visible=True)
            window['-DocBrowse-'].update(visible=True)
        case 'Pseudo':
             # Display correct instructions and only the Illustrator file field.
            window['-Info-'].update(value=config['Instructions']['Pseudo'])
            window['-PBar-'].update(visible=True)
            window['-AiPath-'].update(visible=True)
            window['-AiBrowse-'].update(visible=True)
            window['-DocPath-'].update(visible=False)
            window['-DocBrowse-'].update(visible=False)


def ExtractText(AiApp, WordApp, AiFile: Path,
                Hidden: bool, Locked: bool, PDF: bool):

    yield 'Opening .ai file'
    # Create new blank word file in the background and add table the with Source, Target at top and as many rows under that as there are TextFrames  # noqa: E501
    AiDoc = AiApp.Open(AiFile.as_posix())
    WordFile = WordApp.Documents.Add()
    WordFile = WordApp.ActiveDocument
    WordFile.Application.DisplayAlerts = 0
    WordFile.ShowGrammaticalErrors = False
    WordFile.ShowSpellingErrors = False
    WordFile.SpellingChecked = True
    rng = WordFile.Range()
    table = WordFile.Tables.Add(rng, len(AiDoc.TextFrames)+1, 2)
    table.Cell(1, 1).Range.Text = 'Source'
    table.Cell(1, 1).Range.font.Hidden = True
    table.Cell(1, 2).Range.Text = 'Target'
    table.Cell(1, 2).Range.font.Hidden = True
    max = AiDoc.TextFrames.Count
    count = 0
    for index, frame in enumerate(AiDoc.TextFrames):
        # For each TextFrame, if hidden or locked and option was ticked, go to next Textframe, leaving that row blank, else grab text and put it in the Word table. Row corresponds to index of the TextFrame. Yield current progress with every iterations for gui progress bar  # noqa: E501
        yield f'Extracting text, Segment {index + 1} of {max}'
        if frame.Hidden and Hidden:
            continue
        if frame.locked and Locked:
            continue
        count += 1
        table.Cell(index + 2, 1).Range.Text = frame.Contents
        table.Cell(index + 2, 1).Range.Font.Hidden = True
        table.Cell(index + 2, 2).Range.Text = frame.Contents
    if PDF:
        yield 'Exporting to PDF'
        AiDoc.ExportAsFormat(4, f'{AiFile.parent.as_posix()}/{AiFile.name}.pdf')
    AiDoc.Close()
    finalname = AiFile.name
    # Prepare final file name following option chosen
    if Hidden:
        finalname += '_NO_HIDDEN'
    if Locked:
        finalname += '_NO_LOCKED'
    # Loop over the table, and remove any blank row due to hidden/locked Textframes being skipped, only trigger if one of the two is True. Yield current progress with every iterations for gui progress bar   # noqa: E501
    if Hidden or Locked:
        for row in table.Rows:
            if row.Cells(1).Range.Text[:-2] == '':
                row.Delete()
    yield 'Saving Wor file'
    WordFile.SaveAs2(f'{AiFile.parent.as_posix()}/Strings_{finalname}.docx',
                     FileFormat=12)
    WordFile.Close()


def ImportText(AiApp, AiFile: Path, WordApp, WordFile: Path,
               Hidden:bool, Locked: bool, PDF: bool):

    WordDoc = WordApp.Documents.Open(WordFile.__str__(), Visible=False)
    WordDoc = WordApp.ActiveDocument
    AiDoc = AiApp.Open(AiFile.as_posix())
    Table = WordDoc.Tables(1)
    max = AiDoc.TextFrames.Count
    i = 2
    for index, frame in enumerate(AiDoc.TextFrames):
        # Start looping over all TextFrame, keeping track of the current row (i) and only incrementing if text is imported. Yield current progress with every iterations for gui progress bar   # noqa: E501
        yield f'Populating .ai File, segment {index} of {max}'
        # If hidden or locked (based on filename), move to next TextFrame
        if Hidden and frame.hidden:
            continue
        if Locked and frame.locked:
            continue
        # Not copying the last 2 characters as they're always the combo Ascii 13 and Ascii 10, which is what Word uses to mark the end of a cell.   # noqa: E501
        frame.Contents = Table.Cell(i, 2).Range.Text[:-2]
        i += 1
    WordDoc.Close()
    yield 'Saving Merged file'
    AiDoc.SaveAs(f'{AiFile.parent.as_posix()}/Merged_{AiFile.name}')
    if PDF:
        yield 'Exporting as PDF'
        AiDoc.ExportAsFormat(4, f'{AiFile.parent.as_posix()}/Merged_{AiFile.name}.pdf')
    AiDoc.Close()


def Pseudo(AiApp, AiFile: Path, Hidden:bool, Locked: bool, PDF: bool):
    yield 'Opening .ai file'
    AiDoc = AiApp.Open(AiFile.as_posix())
    for index, frame in enumerate(AiDoc.TextFrames):
        # Start looping over all TextFrames, pseudotranslating as needed following chosen options.  # noqa: E501
        yield f'PseudoTranslating text, Segment {index + 1} of' +\
             f' {len(AiDoc.TextFrames)}'
        if Hidden and frame.hidden:
            continue
        if Locked and frame.locked:
            continue
        # If text should be pseudotranslated, call ReplaceText on the contents
        frame.Contents = replacetext(frame.Contents)
        yield 'Saving Pseudo'
    AiDoc.SaveAs(f'{AiFile.parent.as_posix()}/Pseudo_{AiFile.name}')
    if PDF:
        yield 'Exporting as PDF'
        AiDoc.ExportAsFormat(4, f'{AiFile.parent.as_posix()}/Pseudo_{AiFile.name}.pdf')
    AiDoc.Close()


def replacetext(source: str):
    # Replaces vowels with a random accented variant using unicode. Based on default AP pseudotranslator config, can be modified as needed. Everything is done on the string in memory, before sending it back to the main loop for speed purposes.  # noqa: E501
    source = source.replace('a', choice(list('\u00e0\u00e1\u00e2\u00e3\u00e4\u00e5\u00e6')))  # noqa: E501
    source = source.replace('e', choice(list('\u00e8\u00e9\u00ea\u00eb')))
    source = source.replace('i', choice(list('\u00ec\u00ed\u00ee\u00ef')))
    source = source.replace('o', choice(list('\u00f2\u00f3\u00f4\u00f5\u00f6')))
    source = source.replace('u', choice(list('\u00f9\u00fa\u00fb\u00fc\u00fd')))
    source = source.replace('A', choice(list('\u00c0\u00c1\u00c2\u00c3\u00c4\u00c5\u00c6')))  # noqa: E501
    source = source.replace('E', choice(list('\u00c8\u00c9\u00ca\u00cb')))
    source = source.replace('I', choice(list('\u00cc\u00cd\u00ce\u00cf')))
    source = source.replace('O', choice(list('\u00d2\u00d3\u00d4\u00d5\u00d6')))
    source = source.replace('U', choice(list('\u00d9\u00da\u00db\u00dc')))
    return source


def Collapsible(layout, key, title='', arrows=(gui.SYMBOL_DOWN, gui.SYMBOL_UP),
                collapsed=False):
    # Collapsible function to have a nice options dropdown, taken straight from PySimpleGui Cookbook.  # noqa: E501
    return gui.Column([[gui.T((arrows[1] if collapsed else arrows[0]),
                      enable_events=True, k=key+'-BUTTON-'), gui.T(title,
                      enable_events=True, key=key+'-TITLE-')],
                      [gui.pin(gui.Column(layout, key=key,
                       visible=not collapsed, metadata=arrows))]], pad=(0, 0))


# Options layout, PDF is turned on by default
Options = [
    [gui.Checkbox(text='Skip Hidden', auto_size_text=True,
     key='hidden', size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='Skip Locked', auto_size_text=True,
     key='locked', size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='Export to PDF', auto_size_text=True,
     key='PDF', size=(22, 1), pad=(0, 0), metadata='option', default=True)]
]

# Main Layout, File Browser are file type restricted to either .ai or .docx. Also contains the Progress bar settings  # noqa: E501
layout = [
    [gui.Button(button_text='Extract', key='-Extract-'),
     gui.Button(button_text='Import', key='-Import-'),
     gui.Button(button_text='Pseudo', key='-Pseudo-')],
    [gui.Text(text='Select \'Extract\' or \'Import\' to start.',
              key='-Info-', size=(65, 15))],
    [gui.InputText(default_text='Path to .ai files.', key='-AiPath-',
                   visible=False, ),
     gui.FilesBrowse(button_text='Browse', target='-AiPath-', key='-AiBrowse-',
                     visible=False,
                     file_types=(('Illustrator files', '*.ai'),))],
    [gui.InputText(default_text='Path to .docx files.', key='-DocPath-',
                   visible=False),
     gui.FilesBrowse(button_text='Browse', target='-DocPath-',
                     visible=False, key='-DocBrowse-',
                     file_types=(('Word files', '*.docx'),),)],
    [Collapsible(layout=Options, key='Options',
                 title='Options', collapsed=True)],
    [gui.Submit(button_text='Start', key='-Start-', visible=False)],
    [gui.ProgressBar(max_value=1, orientation='horizontal',
                     key='-PBar-', size=(50, 5), visible=False)],
    [gui.Text(text='', key='-PFileName-')],
    [gui.Text(text='', key='-PStep-')]
]

# Create the window and start the main loop
window = gui.Window(title='Illustrator2Doc', layout=layout)
while True:
    # On any event, get all the info
    event, values = window.read()  # type: ignore
    match event:
        # On Exit, break out of the main loop to close the window
        case gui.WIN_CLOSED | 'Exit':
            break
        # if clicking on Option Arrow dropdown, show the options
        case 'Options-BUTTON-':
            window['Options'].update(visible=not window['Options'].visible)
            window['Options'+'-BUTTON-'].\
                update(window['Options'].metadata[0] if
                       window['Options'].visible else
                       window['Options'].metadata[1])
        # If clicking on one of the buttons at the top, displays the correct info and set the 'Task' variable for later  # noqa: E501
        case '-Extract-':
            SetFields('Extract')
            Task = 'Extract'
        case '-Pseudo-':
            SetFields('Pseudo')
            Task = 'Pseudo'
        case '-Import-':
            SetFields('Import')
            Task = 'Import'
        case '-Start-':
            # Make a python list of all selected files and check the value of 'Task'
            AiFileList = values['-AiPath-'].split(';')
            WordFileList = values['-DocPath-'].split(';')
            match Task: #type: ignore
                case 'Extract':
                    # Catch issues before looping
                    AiFileList = FixAiFileList(AiFileList)
                    if len(AiFileList) == 0:
                        gui.popup_error(config['errors']['No_ai_left'],
                                        title='No .ai file')
                        continue
                    window['-PStep-'].update(
                        value='Opening Illustrator and Word')
                    # Start an instance of Illustrator and Word, same instance is reused for every file and closed when processing is done  # noqa: E501
                    AiApp = DispatchEx('Illustrator.Application')
                    WordApp = DispatchEx('Word.Application')
                    # Remove Illustrator user warnings for fonts and links missing and start looping over the files  # noqa: E501
                    AiApp.UserInteractionLevel = -1
                    for AiFileindex, AiFile in enumerate(AiFileList):
                        AiFile = Path(AiFile)
                        # Display name of the file, and call Extract
                        window['-PFileName-'].update(value=AiFile.name)
                        for step in ExtractText(
                            AiApp, WordApp, AiFile,
                            window['hidden'].get(), # type: ignore
                            window['locked'].get(), # type: ignore
                            window['PDF'].get()): # type: ignore
                            # update progress bar based on the yield
                            window['-PStep-'].update(value=step)
                            window['-PBar-'].update(
                                current_count=(
                                    AiFileindex + 1)/len(AiFileList))
                    WordApp.Quit()
                    AiApp.Quit()
                case 'Import':
                    # Check number of files in both lists, if different, warn the user with a popup  # noqa: E501
                    # Catch issues before looping
                    AiFileList = FixAiFileList(AiFileList)
                    WordFileList = FixWordFileList(WordFileList)
                    if len(AiFileList) == 0:
                        gui.popup_error(config['errors']['No_ai_left'],
                                        title='No .ai file')
                        continue
                    if len(WordFileList) == 0:
                        gui.popup_error(config['errors']['No_docx_left'],
                                        title='No .docx file')
                        continue
                    if len(AiFileList) != len(WordFileList):
                        gui.popup_error(
                            config['errors']['Dif_Len'],
                            auto_close_duration=4, title='Different amount of files')
                    window['-PStep-'].update(
                        value='Opening Illustrator and Word')
                    # Start an instance of Illustrator and Word, same instance is reused for every file and closed when processing is done. Also initalize empty list for potential files that can't be matched.  # noqa: E501
                    NoMatch = list()
                    AiApp = DispatchEx('Illustrator.Application')
                    WordApp = DispatchEx('Word.Application')
                    # Remove Illustrator user warnings for fonts and links missing and start looping over the files  # noqa: E501
                    AiApp.UserInteractionLevel = -1
                    # Start looping over ai files
                    for AiFileindex, AiFile in enumerate(AiFileList):
                        AiFile = Path(AiFile)
                        window['-PFileName-'].update(value=AiFile.name)
                        Found = False
                        # Loop through the names of all Word files, looking for a match with regex. If found, open both ai word file and start importing, if not, store the name of the ai file for report and go to the next file  # noqa: E501
                        for WordIndex, WordFile in enumerate(WordFileList):
                            if search(r'Strings_' + AiFile.name +
                                      r'(_NO_HIDDEN)*(_NO_LOCKED)*-\w{2}-\w{2}',
                                      WordFile):
                                Found = True
                                Hid = False
                                lock = False
                                if search('_NO_HIDDEN', WordFile):
                                    Hid = True
                                if search('_NO_LOCKED', WordFile):
                                    lock = True
                                WordFile = Path(WordFile)
                                for step in ImportText(
                                        AiApp, AiFile, WordApp, WordFile,
                                        window['hidden'].get(), # type: ignore
                                        window['locked'].get(), # type: ignore
                                        window['PDF'].get()): # type: ignore
                                    window['-PStep-'].update(value=step)
                                    window['-PBar-'].update(
                                        current_count=(
                                            WordIndex + 1)/len(WordFileList))
                        if not Found:
                            # if no match, put the ai file name in a list for later
                            NoMatch.append(AiFile.name)
                    WordApp.Quit()
                    AiApp.Quit()
                    if len(NoMatch) != 0:
                        gui.popup(f'The following files could not be matched and were skipped:\n{NoMatch}')  # noqa: E501
                case 'Pseudo':
                    AiFileList = FixAiFileList(AiFileList)
                    if len(AiFileList) == 0:
                        gui.popup_error(config['errors']['No_ai_left'])
                        continue
                    window['-PStep-'].update(
                        value='Opening Illustrator ')
                    # Start an instance of Illustrator only, same instance is reused for every file and closed when processing is done  # noqa: E501
                    AiApp = DispatchEx('Illustrator.Application')
                    # Remove Illustrator user warnings for fonts and links missing and start looping over the files  # noqa: E501
                    AiApp.UserInteractionLevel = -1
                    for AiFileindex, AiFile in enumerate(AiFileList):
                        AiFile = Path(AiFile)
                        window['-PFileName-'].update(value=AiFile.name)
                        for step in Pseudo(
                                AiApp, AiFile,
                                window['hidden'].get(), # type: ignore
                                window['locked'].get(), # type: ignore
                                window['PDF'].get()): # type: ignore
                            window['-PStep-'].update(value=step)
                            window['-PBar-'].update(
                                current_count=(
                                    AiFileindex + 1)/len(AiFileList))
                    AiApp.Quit()
            window['-PFileName-'].update(value='')
            window['-PStep-'].update(value='Done!')
            window['-PBar-'].update(current_count=100)
window.close()
