from pathlib import Path
from win32com.client import GetActiveObject, DispatchEx
import PySimpleGUI as gui


def SetFields(Option: str):

    window['-Start-'].update(visible=True)
    match Option:
        case 'Extract':
            window['-AiPath-'].update(visible=True)
            window['-AiBrowse-'].update(visible=True)
            window['-DocPath-'].update(visible=False)
            window['-DocBrowse-'].update(visible=False)
        case 'Import':
            window['-AiPath-'].update(visible=True)
            window['-AiBrowse-'].update(visible=True)
            window['-DocPath-'].update(visible=True)
            window['-DocBrowse-'].update(visible=True)


def ExtractText(AiApp, WordApp, File: Path):

    Strings = list()
    AiDoc = AiApp.Open(File.as_posix())
    for index, textframe in enumerate(AiDoc.TextFrames):
        yield f'Gatherings text strings, {index} of {len(AiDoc.TextFrames)}'
        Strings.append(textframe.Contents)
    AiDoc.Close()
    WordDoc = WordApp.Documents.Add(Visible=False)
    Table = WordDoc.Tables.Add()
    for string in Strings


def FillWord(WordApp, Strings: list):


layout = [
    [gui.Button(button_text='Extract', key='-Extract-'),
     gui.Button(button_text='Import', key='-Import-')],
    [gui.InputText(default_text='Path to .ai files.', key='-AiPath-',
                   visible=False),
     gui.FilesBrowse(button_text='Browse', target='-AiPath-', key='-AiBrowse-',
                     visible=False,
                     file_types=(('Illustrator files', '*.ai'),))],
    [gui.InputText(default_text='Path to .docx files.', key='-DocPath-',
                   visible=False),
     gui.FilesBrowse(button_text='Browse', target='-DocPath-',
                     visible=False, key='-DocBrowse-',
                     file_types=(('Word files', '*.doc'),),)],
    [gui.Submit(button_text='Start', key='-Start-', visible=False)],
    [gui.ProgressBar(max_value=100, orientation='horizontal',
                     key='-PBar-')],
    [gui.Text(text='', key='-PFileName-')],
    [gui.Text(text='', key='-PStep-')]
]

window = gui.Window(title='Illustrator2Doc', layout=layout)
NewAi = False
NewDoc = False
Extract = False
while True:
    event, values = window.read()  # type: ignore
    match event:
        case gui.WIN_CLOSED | 'Exit':
            break
        case '-Extract-':
            SetFields('Extract')
            window.refresh()
            Extract = True
        case 'Import':
            SetFields('Import')
            Extract = False
        case '-Start-':
            FileList = values['-AiPath-'].split(';')
            try:
                AiApp = GetActiveObject('Illustrator.Application')
            except Exception:
                AiApp = DispatchEx('Illustrator.Application')
                AiApp.Visible = False
                NewAi = True
            WordApp = DispatchEx('Word.Application')
            WordApp.Visible = False
            AiApp.UserInteractionLevel = -1
            for file in FileList:
                file = Path(file)
                window['-PFileName-'].update(value=file.name)
                if Extract:
                    for step in ExtractText(AiApp, WordApp, file):
                        window['-PStep-'].update(value=step)
            AiApp.UserInteractionLevel = 2
            WordApp.Quit()
            if NewAi:
                AiApp.Quit()
window.close()
