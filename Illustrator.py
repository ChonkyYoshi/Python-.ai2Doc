from pathlib import Path
from win32com.client import DispatchEx

def ExtractStrings(TextFrames) -> dict[str,list[str]]:
    Strings: dict[str, list[str]] = dict()
    Tops = [TextFrame.Top for TextFrame in TextFrames]
    Tops = sorted(Tops)
    for Top in Tops:
        for TextFrame in TextFrames:
            if TextFrame.Top == Top:
                if TextFrame.Uuid not in Strings:
                    Strings[TextFrame.Uuid] = list()
                for Paragraph in TextFrame.Paragraphs:
                    Strings[TextFrame.Uuid].append(Paragraph.Contents)
            
    return Strings



AiApp = DispatchEx('Illustrator.Application')
AiFile = Path(r'C:/Users/transformers/Documents/Python-.ai2Doc/test/TW Ophtha_NL4_Recreation_v3.0_31Jul23.ai')  # noqa: E501
AiApp.Open(AiFile.as_posix())
AiDoc = AiApp.ActiveDocument
AiDoc.Export(r'C:/Users/transformers/Documents/Python-.ai2Doc/test/TW Ophtha_NL4_Recreation_v3.0_31Jul23.svg', 3)  # noqa: E501

# for ArtboardIndex in range(1, AiDoc.Artboards.Count + 1):
#     AiDoc.Artboards.SetActiveArtboardIndex(ArtboardIndex)
#     AiDoc.SelectObjectsOnActiveArtboard()
#     for Object in AiDoc.Selection:
#         try:
#             if Object.GroupItem.Name == 'Ai2DocText':
#                 Str = ExtractStrings(Object.TextFrames)
#                 print(Str)
#         except Exception:
#             continue