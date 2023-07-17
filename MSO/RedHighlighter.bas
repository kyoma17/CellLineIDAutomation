'Written by Kenny Ma  
'Last Modified: 2023-01-08

' This macro changes the font of a number to red if it exists in Allele columns
' This macro is used to highlight the numbers in the Allele columns that match the numbers in the DMZ_BMP column
' This is paired with the CellLineID Automation as an injection Code

' Changes the font of a number to red if exists in Allele columns
Sub ChangeMatchingToRed()
    Dim doc         As Word.Document
    Set doc = ActiveDocument
    
    Dim tbl         As Word.Table
    Set tbl = doc.Tables(2)        ' Assumes there is only one table in the document
    
    Dim row         As Word.row
    Dim Marker_List As String
    Marker_List = "D5S818, D13S317, D7S820, D16S539, vWA, TH01, AMEL, TPOX, CSF1PO, D21S11, D10S1248, D12S391, D18S51, D19S433, D1S1656, D22S1045, D2S1338, D2S441, D3S1358, D8S1179, DYS391, FGA, Penta D, Penta E"
    
    For Each row In tbl.Rows
        ' if the first cell in the row contains the word in the Marker_List then change the font to red
        Dim Marker  As String
        Marker = CleanString(row.Cells(1).Range.Text)
        
        ' If Marker is in the Marker_List then change the font to red
        If InStr(Marker_List, Marker) > 0 Then
            
            Dim Allele1, Allele2, Allele3, Allele4, DMZ_BMP As String
            Allele1 = CleanString(row.Cells(2).Range.Text)
            Allele2 = CleanString(row.Cells(3).Range.Text)
            Allele3 = CleanString(row.Cells(4).Range.Text)
            Allele4 = CleanString(row.Cells(5).Range.Text)
            DMZ_BMP = CleanString(row.Cells(6).Range.Text)
            
            ' Search DMZ_BMP for numbers, if found then change only those numbers to red, leave the rest black
            If InStr(DMZ_BMP, Allele1) > 0 Then
                ReplaceWithRed Allele1, row.Cells(6).Range
            End If
            
            If InStr(DMZ_BMP, Allele2) > 0 Then
                ReplaceWithRed Allele2, row.Cells(6).Range
            End If
            
            If InStr(DMZ_BMP, Allele3) > 0 Then
                ReplaceWithRed Allele3, row.Cells(6).Range
            End If
            
            If InStr(DMZ_BMP, Allele4) > 0 Then
                ReplaceWithRed Allele4, row.Cells(6).Range
            End If
            
        End If
    Next
    
End Sub

Sub ReplaceWithRed(Allele, target)
    With target.Find
        .ClearFormatting
        .Text = Allele
        .MatchCase = TRUE
        .Forward = TRUE
    End With
    
    If target.Find.Execute Then
        target.Font.Color = wdColorRed
    End If
End Sub

Function CleanString(str As String) As String
    ' Removes line break and bell characters from the input string
    CleanString = Replace(str, vbCr, "")
    CleanString = Replace(CleanString, vbLf, "")
    CleanString = Replace(CleanString, vbCrLf, "")
    CleanString = Replace(CleanString, Chr(13), "")
    CleanString = Replace(CleanString, Chr(7), "")
End Function

