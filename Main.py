import win32com.client as com
import PySimpleGUI as gui


def Extracttext():
    window['Step'].update('Launching Illustrator')
    window['Bar'].update(10)
    IllustratorApp = com.DispatchEx('Illustrator.Application')
    window['Step'].update('Opening ai file')
    window['Bar'].update(20)
    AiDoc = IllustratorApp.Open(window['EInput'].get())

    strings = {}
    window['Step'].update('Renaming Text Frames and gathering text')
    window['Bar'].update(40)
    for index, frame in enumerate(AiDoc.TextFrames):
        frame.Name = str(index)
        strings[index] = frame.Contents

    window['Step'].update('Launching Word')
    window['Bar'].update(60)
    WordApp = com.DispatchEx('Word.Application')

    WordDoc = WordApp.Documents.Add()
    WordDoc = WordApp.ActiveDocument
    rng = WordDoc.Range()
    table = WordDoc.Tables.Add(rng, len(strings)+1, 3)
    window['Step'].update('Populating Table with strings')
    window['Bar'].update(80)
    table.Cell(1, 1).Range.Text = 'TextFrame'
    table.Cell(1, 2).Range.Text = 'Source'
    table.Cell(1, 3).Range.Text = 'Target'

    for key, value in strings.items():
        table.Cell(key + 2, 1).Range.Text = str(key)
        table.Cell(key + 2, 2).Range.Text = value
        table.Cell(key + 2, 3).Range.Text = value

    table.Columns(1).Select()
    WordDoc.Application.Selection.Font.Hidden = True
    table.Columns(2).Select()
    WordDoc.Application.Selection.Font.Hidden = True
    table.Rows(1).Select()
    WordDoc.Application.Selection.Font.Hidden = True
    window['Step'].update('Saving and quitting')
    window['Bar'].update(90)
    WordDoc.SaveAs(window['EInput'].get()[:-3] + '.docx')
    WordApp.Quit()
    AiDoc.Save()
    IllustratorApp.Quit()


def ImportText():
    window['Step'].update('Launching Illustrator')
    window['Bar'].update(0)
    IllustratorApp = com.DispatchEx('Illustrator.Application')
    window['Step'].update('Opening ai file')
    window['Bar'].update(20)
    AiDoc = IllustratorApp.Open(window['EInput'].get())

    window['Step'].update('Launching Word')
    window['Bar'].update(30)
    WordApp = com.DispatchEx('Word.Application')
    window['Step'].update('Opening Word file')
    window['Bar'].update(40)
    WordDoc = WordApp.Documents.Open(window['IInput'].get())
    WordDoc = WordApp.ActiveDocument
    WordDoc.Select()
    WordDoc.Application.Selection.Font.Hidden = False

    table = WordDoc.Tables(1)
    j = 2
    window['Step'].update('Populating ai file')
    window['Bar'].update(60)
    for i in AiDoc.TextFrames:
        i.Contents = table.Cell(j, 3).Range.Text
        i.Name = table.Cell(j, 3).Range.Text
        j += 1
    window['Step'].update('Saving and quitting')
    window['Bar'].update(90)
    AiDoc.SaveAs(window['EInput'].get()[:-3] + '_Merged.ai')

    WordApp.Quit()
    IllustratorApp.Quit(2)


layout = [
    [gui.Button(button_text='Extract Text', key='Extract'),
     gui.Button(button_text='Import Text', key='Import')],
    [gui.Text('Ai file', size=15, visible=False, key='EText'),
     gui.Input(default_text='', visible=False, key='EInput'),
     gui.FileBrowse(visible=False, key='EBrowse', target='EInput')],
    [gui.Text('Word file', size=15, visible=False, key='IText'),
     gui.Input(default_text='', visible=False, key='IInput'),
     gui.FileBrowse(visible=False, key='IBrowse', target='IInput')],
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
                Extracttext()
                window['Step'].update('Done!')
                window['Bar'].update(100)
            else:
                ImportText()
                window['Step'].update('Done!')
                window['Bar'].update(100)

window.close()
