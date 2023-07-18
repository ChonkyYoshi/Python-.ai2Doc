from win32com.client import DispatchEx, GetActiveObject
import PySimpleGUI as gui
from regex import match, findall


def split(FullPath):
    PathOnly = ''
    for find in findall(r'[^\/]+?\/', FullPath):
        PathOnly += find
    FileOnly = match(r'(?r)[^\/]+', FullPath).group()
    return (FullPath, PathOnly.replace('/', '\\'), FileOnly)


def StartAI():
    AiApp = DispatchEx('Illustrator.Application')
    return AiApp


def Extracttext(AiApp, PathOnly, FileOnly):
    yield 'Opening ai file', 10
    AiDoc = AiApp.Open(PathOnly + FileOnly)

    strings = dict()
    for index, frame in enumerate(AiDoc.TextFrames):
        yield f'Extracting text, {index}/{len(AiDoc.TextFrames)}',\
            10 + (65*(index/len(AiDoc.TextFrames)))
        strings[index] = frame.Contents

    yield 'Opening Word', 75
    WordApp = DispatchEx('Word.Application')

    WordDoc = WordApp.Documents.Add()
    WordDoc = WordApp.ActiveDocument
    rng = WordDoc.Range()
    table = WordDoc.Tables.Add(rng, len(strings)+1, 2)
    table.Cell(1, 1).Range.Text = 'Source'
    table.Cell(1, 2).Range.Text = 'Target'

    j = 0
    for key, value in strings.items():
        yield f'Populating Word file, {j}/{len(strings)}',\
            75 + (25*(j/len(strings)))
        table.Cell(key + 2, 1).Range.Text = value
        table.Cell(key + 2, 2).Range.Text = value
        j += 1

    table.Columns(1).Select()
    WordDoc.Application.Selection.Font.Hidden = True
    table.Rows(1).Select()
    WordDoc.Application.Selection.Font.Hidden = True
    WordDoc.SaveAs(PathOnly + FileOnly + '.docx')
    WordApp.Quit()
    AiDoc.Close()


def ImportText(AiApp, AiPathOnly, AiFileOnly, WdPathOnly, WdFileOnly):
    yield 'Opening ai file', 5
    AiDoc = AiApp.Open(AiPathOnly + AiFileOnly)
    yield 'Opening Word file', 10
    WordApp = DispatchEx('Word.Application')
    WordDoc = WordApp.Documents.Open(WdPathOnly + WdFileOnly)
    WordDoc = WordApp.ActiveDocument
    table = WordDoc.Tables(1)
    table.Columns(1).Select()
    WordApp.Selection.Range.Font.Hidden = False
    j = 2
    for index, i in enumerate(AiDoc.TextFrames):
        yield f'merging {index + 1}/{len(AiDoc.TextFrames)}',\
            10 + (85*(index/len(AiDoc.TextFrames)))
        i.Contents = table.Cell(j, 2).Range.Text
        j += 1
    AiDoc.SaveAs(AiPathOnly + AiFileOnly[:-3] + '_Merged.ai')
    WordApp.Quit()
    AiApp.Quit()


layout = [
    [gui.Button(button_text='Extract Text', key='Extract'),
     gui.Button(button_text='Import Text', key='Import')],
    [gui.Text('Ai file', size=15, visible=False, key='EText'),
     gui.Input(default_text='', visible=False, key='EInput'),
     gui.FilesBrowse(visible=False, key='EBrowse', target='EInput')],
    [gui.Text('Word file', size=15, visible=False, key='IText'),
     gui.Input(default_text='', visible=False, key='IInput'),
     gui.FilesBrowse(visible=False, key='IBrowse', target='IInput')],
    [gui.Submit(visible=False, key='Run')],
    [gui.Text('', visible=False, key='Step')],
    [gui.ProgressBar(max_value=100, key='Bar', visible=False)]
]

window = gui.Window('Illustrator Extract/Merge', layout)

while True:
    event, values = window.read()
    match event:
        case 'Exit' | gui.WIN_CLOSED:
            break
        case 'Extract':
            window['EInput'].update(visible=True)
            window['EBrowse'].update(visible=True)
            window['EText'].update(visible=True)
            window['IInput'].update(visible=False)
            window['IBrowse'].update(visible=False)
            window['IText'].update(visible=False)
            window['Run'].update(visible=True)
            window['Step'].update(visible=False)
            window['Bar'].update(visible=False)
            Extract = True
        case 'Import':
            window['EInput'].update(visible=True)
            window['EBrowse'].update(visible=True)
            window['EText'].update(visible=True)
            window['IInput'].update(visible=True)
            window['IBrowse'].update(visible=True)
            window['IText'].update(visible=True)
            window['Run'].update(visible=True)
            window['Step'].update(visible=False)
            window['Bar'].update(visible=False)
            Extract = False
        case 'Run':
            window['Step'].update(visible=True)
            window['Bar'].update(visible=True)
            if Extract:
                FileList = values['EInput'].split(';')
                window['Bar'].update(0, len(FileList))
                window.Refresh()
                try:
                    AiApp = GetActiveObject('Illustrator.Application')
                except Exception:
                    AiApp = StartAI()
                for index, file in enumerate(FileList):
                    FullPath, PathOnly, FileOnly = split(file)
                    for step, prog in Extracttext(AiApp,
                                                  PathOnly, FileOnly):
                        window['Step'].update(FileOnly + '\n' + step)
                        window['Bar'].update(index + (prog/100))
                window['Step'].update('Done!')
                window['Bar'].update(len(FileList))
                # AiApp.Quit()
            else:
                AiFileList = values['EInput'].split(';')
                WordFileList = values['IInput'].split(';')
                window['Bar'].update(0, len(AiFileList))
                window.Refresh()
                MatchDict = dict()
                AiFileList = values['EInput'].split(';')
                WdFileList = values['IInput'].split(';')
                for Aifile in AiFileList:
                    AiFullPath, AiPathOnly, AiFileOnly = split(Aifile)
                    for Wdfile in WdFileList:
                        WdFullPath, WdPathOnly, WdFileOnly = split(Wdfile)
                        if AiFileOnly == WdFileOnly[:-5]:
                            MatchDict[AiFullPath] = WdFullPath
                try:
                    AiApp = GetActiveObject('Illustrator.Application')
                except Exception:
                    AiApp = StartAI()
                for Aifile, WdFile in MatchDict.items():
                    index = 0
                    AiFullPath, AiPathOnly, AiFileOnly = split(Aifile)
                    WdFullPath, WdPathOnly, WdFileOnly = split(Wdfile)
                    for step, prog in ImportText(AiApp, AiPathOnly, AiFileOnly,
                                                 WdPathOnly, WdFileOnly):
                        window['Step'].update(AiFileOnly + '\n' + step)
                        window['Bar'].update(index + (prog/100))
                        index += 1
                window['Step'].update('Done!')
                window['Bar'].update(len(MatchDict))
window.close()
