' Written by Kenny Ma
' Last Modified: 2023-01-14
' Replace the keywords in the header with the replacement text

Sub ReplaceHeaderKeyword(replacement_dictionary As object)
    ' This maco replaces the keywords in the header with the replacement text
    Dim doc As Document
    Set doc = ActiveDocument

    For Each Key In replacement_dictionary
        doc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Find.Execute FindText:=Key, ReplaceWith:=replacement_dictionary(Key)
    Next
End Sub


Sub CombineAllCellID()
    ' copy content from file1 and file2 to this file
    files = "C:\Users\kyo_m\Documents\Code\CellLineOutput\3-95-NALM-6.docx"
    file2 = "C:\Users\kyo_m\Documents\Code\CellLineOutput\3-96-RPMI.docx"

    ListOfFiles = Array(files, file2)

    Dim doc As Document
    Set doc = ActiveDocument

    ' Combine all the files
    For Each file In ListOfFiles
        doc.Range.End = doc.Range.End + 1
        doc.Range.End.InsertFile file
    Next
End Sub