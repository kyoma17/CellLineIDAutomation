' Written by Kenny Ma
' Last Modified: 2023-07-16
' Adds a page break after each sample set

Sub AddPageBreaks()
    Dim tbl As Table
    Dim rng As Range
    Dim i As Integer

    ' Initialize counter
    i = 1

    ' Loop through each table in the document.
    For Each tbl In ActiveDocument.Tables

        ' If i is a multiple of 3, it means we're dealing with the last table in a set.
        If i Mod 3 = 0 Then
            ' Insert a page break after the table.
            Set rng = tbl.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.InsertBreak Type:=wdPageBreak
        End If

        ' Increment the counter.
        i = i + 1
    Next tbl
End Sub


Sub PrintTableIndexes()
    Dim tbl As Table
    Dim index As String

    ' Loop through each table in the document.
    For Each tbl In ActiveDocument.Tables

        ' Get the contents of the first cell in the table.
        index = tbl.Cell(1, 1).Range.Text
        
        ' Clean up the text (remove trailing end-of-cell marker)
        index = Left(index, Len(index) - 2)

        ' Print the index to the Immediate window.
        Debug.Print index

    Next tbl
End Sub

Sub AddLineBreaksAfterEachSample()
    Dim tbl As Table
    Dim rng As Range
    Dim index As String
    Dim previousIndex As String
    Dim methodology As Boolean
    Dim isSecondNumber As Boolean

    ' Initialize variables.
    methodology = False
    isSecondNumber = False
    previousIndex = ""

    ' Loop through each table in the document.
    For Each tbl In ActiveDocument.Tables

        ' Get the contents of the first cell in the table.
        index = tbl.Cell(1, 1).Range.Text

        ' Clean up the text (remove trailing end-of-cell marker).
        index = Left(index, Len(index) - 2)

        ' Check the pattern.
        If index = "Methodology" Then
            methodology = True
        ElseIf methodology And IsNumeric(index) And IsNumeric(previousIndex) Then
            isSecondNumber = True
        End If

        ' If it is the second number in the pattern, insert line break after it.
        If isSecondNumber Then
            Set rng = tbl.Range
            rng.Collapse Direction:=wdCollapseEnd
            rng.InsertBreak Type:=wdLineBreak

            ' Reset the pattern flags.
            methodology = False
            isSecondNumber = False
        End If

        ' Remember the previous index for the next loop iteration.
        previousIndex = index
    Next tbl
End Sub
