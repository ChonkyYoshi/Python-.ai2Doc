from pathlib import Path
from win32com.client import DispatchEx
import PySimpleGUI as gui
from regex import search
from random import choice


def SetFields(Option: str):

    window['-Start-'].update(visible=True)
    match Option:
        case 'Extract':
            window['-Info-'].update(value='''To Extract text from .ai files:
- Choose the files you want to process using the \'Browse\' button then click on \'Start\'.
- The program will automatically launch an instance of Adobe Illustrator and Word then proceed to extract segments from all the selected .ai files.
- Extracted segments are saved in a Word file called Strings_<.ai file name>.docx.
''')  # noqa: E501
            window['-PBar-'].update(visible=True)
            window['-AiPath-'].update(visible=True)
            window['-AiBrowse-'].update(visible=True)
            window['-DocPath-'].update(visible=False)
            window['-DocBrowse-'].update(visible=False)
        case 'Import':
            window['-Info-'].update(value='''To create translated .ai files:
- Choose the .ai files you want to translate using the \'Browse\' button.
- Choose the translated .docx files containing the translations using the \'Browse\' button then click on \'Start\'.
- The program will automatically launch an instance of Adobe Illustrator and Word then proceed to match a word file to its corresponding .ai file to impor it.
- Merged .ai file are saved as called Merged_<.ai file name> and exported to pdf.
IMPORTANT: the translated .docx file NEEDS to be named <.ai file name>-xx-YY.docx. Else an error will occur and the file will be skipped.''')  # noqa: E501
            window['-PBar-'].update(visible=True)
            window['-AiPath-'].update(visible=True)
            window['-AiBrowse-'].update(visible=True)
            window['-DocPath-'].update(visible=True)
            window['-DocBrowse-'].update(visible=True)
        case 'Pseudo':
            window['-Info-'].update(value='''To pseudotranslate .ai files:
- Choose the .ai files you want to translate using the \'Browse\' button then click on \'Start\'.
- The program will automatically launch an instance of Adobe Illustrator then proceed to pseudotranslate all the text it can find in the file with random Chinese characters (including consonants and numbers).
- Pseudotranslated .ai file are saved as called Pseudo_<.ai file name> and exported to pdf.
IMPORTANT: This step DOES NOT extract and save a Word file.''')  # noqa: E501
            window['-PBar-'].update(visible=True)
            window['-AiPath-'].update(visible=True)
            window['-AiBrowse-'].update(visible=True)
            window['-DocPath-'].update(visible=False)
            window['-DocBrowse-'].update(visible=False)


def ExtractText(AiApp, WordApp, AiFile: Path, Hidden: bool, Locked: bool):

    yield 'Opening .ai file'
    AiDoc = AiApp.Open(AiFile.as_posix())
    WordFile = WordApp.Documents.Add()
    WordFile = WordApp.ActiveDocument
    rng = WordFile.Range()
    table = WordFile.Tables.Add(rng, len(AiDoc.TextFrames)+1, 2)
    table.Cell(1, 1).Range.Text = 'Source'
    table.Cell(1, 1).Range.font.Hidden = True
    table.Cell(1, 2).Range.Text = 'Target'
    table.Cell(1, 2).Range.font.Hidden = True
    max = AiDoc.TextFrames.Count
    count = 0
    for index, frame in enumerate(AiDoc.TextFrames):
        yield f'Extracting text, Segment {index + 1} of {max}'
        if frame.Hidden and Hidden:
            continue
        if frame.locked and Locked:
            continue
        count += 1
        table.Cell(index + 2, 1).Range.Text = frame.Contents
        table.Cell(index + 2, 1).Range.Font.Hidden = True
        table.Cell(index + 2, 2).Range.Text = frame.Contents
    AiDoc.Close()
    finalname = AiFile.name
    if Hidden:
        finalname += '_NO_HIDDEN'
    if Locked:
        finalname += '_NO_LOCKED'
    for row in table.Rows:
        if row.Cells(1).Range.Text[:-2] == '':
            row.Delete()
    WordFile.SaveAs2(f'{AiFile.parent.as_posix()}/Strings_{finalname}.docx',
                     FileFormat=12)
    WordFile.Close()


def ImportText(AiApp, AiFile: Path, WordApp, WordFile: Path,
               Hidden:bool = False, Locked: bool = False):

    WordDoc = WordApp.Documents.Open(WordFile.as_posix(), Visible=False)
    WordDoc = WordApp.ActiveDocument
    AiDoc = AiApp.Open(AiFile.as_posix())
    Table = WordDoc.Tables(1)
    max = AiDoc.TextFrames.Count
    i = 2
    for index, frame in enumerate(AiDoc.TextFrames):
        yield f'Populating .ai File, segment {index} of {max}'
        if Hidden and frame.hidden:
            continue
        if Locked and frame.locked:
            continue
        frame.Contents = Table.Cell(i, 2).Range.Text[:-2]
        i += 1
    WordDoc.Close()
    AiDoc.SaveAs(f'{AiFile.parent.as_posix()}/Merged_{AiFile.name}')
    AiDoc.ExportAsFormat(4, f'{AiFile.parent.as_posix()}/Merged_{AiFile.name}.pdf')  # noqa: E501
    AiDoc.Close()


def Pseudo(AiApp, AiFile: Path):
    yield 'Opening .ai file'
    AiDoc = AiApp.Open(AiFile.as_posix())
    for index, textframe in enumerate(AiDoc.TextFrames):
        yield f'PseudoTranslating text, Segment {index + 1} of' +\
             f' {len(AiDoc.TextFrames)}'
        textframe.Contents = replacetext(textframe.Contents)
    AiDoc.SaveAs(f'{AiFile.parent.as_posix()}/Pseudo_{AiFile.name}')
    AiDoc.ExportAsFormat(4, f'{AiFile.parent.as_posix()}/Pseudo_{AiFile.name}.pdf')  # noqa: E501
    AiDoc.Close()


def replacetext(source: str):
    source = source.replace('a', choice(list('\u00e0\u00e1\u00e2\u00e3\u00e4\u00e5\u00e6')))  # noqa: E501
    source = source.replace('e', choice(list('\u00e8\u00e9\u00ea\u00eb')))
    source = source.replace('i', choice(list('\u00ec\u00ed\u00ee\u00ef')))
    source = source.replace('o', choice(list('\u00f2\u00f3\u00f4\u00f5\u00f6')))  # noqa: E501
    source = source.replace('u', choice(list('\u00f9\u00fa\u00fb\u00fc\u00fd')))  # noqa: E501
    source = source.replace('A', choice(list('\u00c0\u00c1\u00c2\u00c3\u00c4\u00c5\u00c6')))  # noqa: E501
    source = source.replace('E', choice(list('\u00c8\u00c9\u00ca\u00cb')))
    source = source.replace('I', choice(list('\u00cc\u00cd\u00ce\u00cf')))
    source = source.replace('O', choice(list('\u00d2\u00d3\u00d4\u00d5\u00d6')))  # noqa: E501
    source = source.replace('U', choice(list('\u00d9\u00da\u00db\u00dc')))
    return source


def Collapsible(layout, key, title='', arrows=(gui.SYMBOL_DOWN, gui.SYMBOL_UP),
                collapsed=False):
    return gui.Column([[gui.T((arrows[1] if collapsed else arrows[0]),
                      enable_events=True, k=key+'-BUTTON-'), gui.T(title,
                      enable_events=True, key=key+'-TITLE-')],
                      [gui.pin(gui.Column(layout, key=key,
                       visible=not collapsed, metadata=arrows))]], pad=(0, 0))


Options = [
    [gui.Checkbox(text='Skip Hidden', auto_size_text=True,
     key='hidden', size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='Skip Locked', auto_size_text=True,
     key='locked', size=(22, 1), pad=(0, 0), metadata='option'),
     gui.Checkbox(text='No PDF Export', auto_size_text=True,
     key='noPDF', size=(22, 1), pad=(0, 0), metadata='option')]
]


layout = [
    [gui.Button(button_text='Extract', key='-Extract-'),
     gui.Button(button_text='Import', key='-Import-'),
     gui.Button(button_text='Pseudo', key='-Pseudo-')],
    [gui.Text(text='Select \'Extract\' or \'Import\' to start.',
              key='-Info-', size=(65, 8))],
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

window = gui.Window(title='Illustrator2Doc', layout=layout)
Task = ''
while True:
    event, values = window.read()  # type: ignore
    match event:
        case gui.WIN_CLOSED | 'Exit':
            break
        case 'Options-BUTTON-':
            window['Options'].update(visible=not window['Options'].visible)
            window['Options'+'-BUTTON-'].\
                update(window['Options'].metadata[0] if
                       window['Options'].visible else
                       window['Options'].metadata[1])
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
            AiFileList = values['-AiPath-'].split(';')
            WordFileList = values['-DocPath-'].split(';')
            match Task:
                case 'Extract':
                    window['-PStep-'].update(
                        value='Opening Illustrator and Word')
                    AiApp = DispatchEx('Illustrator.Application')
                    WordApp = DispatchEx('Word.Application')
                    AiApp.UserInteractionLevel = -1
                    for Aifileindex, Aifile in enumerate(AiFileList):
                        Aifile = Path(Aifile)
                        window['-PFileName-'].update(value=Aifile.name)
                        for step in ExtractText(
                            AiApp, WordApp, Aifile,
                            window['hidden'].get(), # type: ignore
                            window['locked'].get()): # type: ignore
                            window['-PStep-'].update(value=step)
                            window['-PBar-'].update(
                                current_count=(
                                    Aifileindex + 1)/len(AiFileList))
                    WordApp.Quit()
                    AiApp.Quit()
                case 'Import':
                    if len(AiFileList) != len(WordFileList):
                        gui.popup_error('''Number of files do not match!
    Please note that there isn\'t the same amount of Word files and .ai files.''', auto_close_duration=4)  # noqa: E501
                    window['-PStep-'].update(
                        value='Opening Illustrator and Word')
                    AiApp = DispatchEx('Illustrator.Application')
                    WordApp = DispatchEx('Word.Application')
                    AiApp.UserInteractionLevel = -1
                    for Aifileindex, AiFile in enumerate(AiFileList):
                        AiFile = Path(AiFile)
                        window['-PFileName-'].update(value=AiFile.name)
                        Found = False
                        for WordIndex, WordFile in enumerate(WordFileList):
                            if search(r'Strings_' + AiFile.name +
                                      r'(_NO_HIDDEN)*(_NO_LOCKED)*-\w{2}-\w{2}', WordFile):
                                Found = True
                                Hid = False
                                lock = False
                                if search('_NO_HIDDEN', WordFile):
                                    Hid = True
                                if search('_NO_LOCKED', WordFile):
                                    lock = True
                                WordFile = Path(WordFile)
                                for step in ImportText(AiApp, AiFile,
                                                       WordApp, WordFile,
                                                       Hid, lock):
                                    window['-PStep-'].update(value=step)
                                    window['-PBar-'].update(
                                        current_count=(
                                            WordIndex + 1)/len(WordFileList))
                        if not Found:
                            gui.popup_error(f'Couldn\'t find a match for {AiFile.name}. Please make sure that the translated Word file is called Strings_{AiFile.name}-xx-XX.docx')   # noqa: E501
                    WordApp.Quit()
                    AiApp.Quit()
                case 'Pseudo':
                    window['-PStep-'].update(
                        value='Opening Illustrator ')
                    AiApp = DispatchEx('Illustrator.Application')
                    AiApp.UserInteractionLevel = -1
                    for Aifileindex, Aifile in enumerate(AiFileList):
                        Aifile = Path(Aifile)
                        window['-PFileName-'].update(value=Aifile.name)
                        for step in Pseudo(AiApp, Aifile):
                            window['-PStep-'].update(value=step)
                            window['-PBar-'].update(
                                current_count=(
                                    Aifileindex + 1)/len(AiFileList))
                    AiApp.Quit()
            window['-PFileName-'].update(value='')
            window['-PStep-'].update(value='Done!')
            window['-PBar-'].update(current_count=100)
window.close()
