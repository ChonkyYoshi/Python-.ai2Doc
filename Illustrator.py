from pathlib import Path
from win32com.client import DispatchEx
from random import choice

def ExtractStrings(AiApp, AiFile: Path) -> dict[str, list[str]]:
    DocStrings: dict[str, list[str]] = dict()
    AiApp.Open(AiFile.as_posix())
    AiDoc = AiApp.ActiveDocument
    for ArtboardIndex in range(0, AiDoc.Artboards.Count):
        AiDoc.Artboards.SetActiveArtboardIndex(ArtboardIndex)
        AiDoc.SelectObjectsOnActiveArtboard()
        for TextFrame in AiApp.Selection.TextFrames:
            for Paragraph in TextFrame.Paragraphs:
                Paragraph.Contents = Pseudo(Paragraph.Contents)
    return DocStrings

def Pseudo(source: str) -> str:
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


AiApp = DispatchEx('Illustrator.Application')
AiFile = Path(r'C:/Users/transformers/Documents/Python-.ai2Doc/test/TW Ophtha_NL4_Recreation_v3.0_31Jul23.ai')  # noqa: E501
ExtractStrings(AiApp, AiFile)