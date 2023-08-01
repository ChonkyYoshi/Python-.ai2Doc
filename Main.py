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


def ExtractText(AiApp, WordApp, AiFile: Path):

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
    for index, textframe in enumerate(AiDoc.TextFrames):
        yield f'Extracting text, Segment {index + 1} of' +\
             f' {len(AiDoc.TextFrames)}'
        table.Cell(index + 2, 1).Range.Text = textframe.Contents
        table.Cell(index + 2, 1).Range.Font.Hidden = True
        table.Cell(index + 2, 2).Range.Text = textframe.Contents
    AiDoc.Close()
    WordFile.SaveAs2(f'{AiFile.parent.as_posix()}/Strings_{AiFile.name}.docx',
                     FileFormat=12)
    WordFile.Close()


def ImportText(AiApp, AiFile: Path, WordApp, WordFile: Path):

    WordDoc = WordApp.Documents.Open(WordFile.__str__(), Visible=False)
    WordDoc = WordApp.ActiveDocument
    AiDoc = AiApp.Open(AiFile.as_posix())
    Table = WordDoc.Tables(1)
    for i in range(2, Table.Rows.Count + 1):
        yield f'Populating .ai File, segment {i} of {Table.Rows.Count}'
        AiDoc.TextFrames(i - 1).Contents = Table.Cell(i, 2).Range.Text[:-1]
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
    for i in source:
        if i not in list(' \t\n\r\f\v'):
            a = choice(list('\u4E10\u4e11\u4e12\u4e14\u4e15\u4e16\u4e17\u4e18\u4e19'))  # noqa: E501
            source = source.replace(i, a, 1)
    return source


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
                        for step in ExtractText(AiApp, WordApp, Aifile):
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
                    AiApp = DispatchEx('Illustrator.Application')
                    WordApp = DispatchEx('Word.Application')
                    AiApp.UserInteractionLevel = -1
                    for Aifileindex, AiFile in enumerate(AiFileList):
                        AiFile = Path(AiFile)
                        Found = False
                        for WordIndex, WordFile in enumerate(WordFileList):
                            if search(r'Strings_' + AiFile.name +
                                      r'-\w{2}-\w{2}\.docx', WordFile):
                                Found = True
                                WordFile = Path(WordFile)
                                for step in ImportText(AiApp, AiFile,
                                                       WordApp, WordFile):
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
