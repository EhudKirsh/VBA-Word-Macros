Sub UpdateAll()
    Application.ScreenUpdating = False 'This improves performance

    'Firstly hide the field codes so Word doesn't need to update their display
    ActiveWindow.View.ShowFieldCodes = False 'Alt + F9

    'Update all fields in the document, including references, cross references & caption labels
    ActiveDocument.Fields.Update
    'This has to be before StyleBibliography because it resets the style and can also add rows

    StyleBibliography 'This has to be before UpdateTablesOfFiguresAndContents, because the bibliography can spill into more pages, so we start from the end

    UpdateTablesOfFiguresAndContents

    Application.ScreenUpdating = True 'Re-enable screen updating
End Sub

Sub StyleBibliography()
    Application.ScreenUpdating = False 'This improves performance

    'Style the Bibliography References Table: turn https into hyperlinks, adjust columns widths and align text to the left
    Dim T As Table: Set T = FindBibliography: T.AllowAutoFit = False
    Dim httpsPos, spacePos As Integer: Dim cols, C2 As Object

    Set cols = T.columns: Set C2 = cols(2)

    'Optional - pick how many digits of references you have:
    'cols(1).Width = 17 ': C2.Width = 420 '[9]
    'cols(1).Width = 22 ': C2.Width = 415 '[99]
    cols(1).Width = 30 ': C2.Width = 407 '[999]

    C2.AutoFit 'Width

    Dim CellsRange As Cells: Set CellsRange = C2.Cells
    Dim r As Range: Dim cellText, linkText As String
    For Each c In CellsRange
        Set r = c.Range

        r.ParagraphFormat.Alignment = wdAlignParagraphLeft 'Align Left

        'Hyperlinks
        cellText = r.Text: cellText = Left(cellText, Len(cellText) - 2)
        httpsPos = InStr(cellText, "https")
        If httpsPos > 0 Then
            spacePos = InStr(httpsPos, cellText, " ") 'Find the first space after "https"
            If spacePos = 0 Then spacePos = Len(cellText) + 1 'Use text length if no space is found

            'Extract the link text (URL)
            linkText = Mid(cellText, httpsPos, spacePos - httpsPos - 1) 'Assuming there's a dot '.' just before thespace ' '

            r.Start = r.Start + httpsPos - 1 'Assuming there's a dot '.' just before thespace ' '
            r.End = r.Start + Len(linkText)

            ActiveDocument.Hyperlinks.Add Anchor:=r, Address:=linkText
        End If
    Next c

    Application.ScreenUpdating = True 'Re-enable screen updating
End Sub

Function FindBibliography() As Table
    Application.ScreenUpdating = False 'This improves performance

    'There is no in-built syntax to find the Bibliography References Table. This function finds it and attempts to do so in the most quick and efficient way.
    Dim p As Integer ' n 'p = page, n = new page
    Dim RangeFields As Fields

    Dim DocTables As Tables: Set DocTables = ActiveDocument.Tables
    Dim rng As Range: Set rng = ActiveDocument.Range

    For Each T In DocTables 'T = Table
        n = T.Range.Information(wdActiveEndPageNumber): If p = n Then GoTo SkipLoop 'If a single page has multiple tables, skip because we already checked this page
        'I search all the field codes in the page because I can't figure the relation (parent, sibling, etc) because the Bibliography table and field code.
        'But at least I narrow down the search only to pages that have ActiveDocument.Tables.

        rng.Start = rng.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=n).Start
        rng.End = rng.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=n + 1).Start

        Set RangeFields = rng.Fields 'Fields on the page of the table

        For Each fld In RangeFields
            If fld.Type = wdFieldBibliography Then
                Set FindBibliography = fld.Result.Tables(1)
                MsgBox "Bibliography is (ends) on page " & FindBibliography.Range.Information(wdActiveEndPageNumber)
                Application.ScreenUpdating = True 'Re-enable screen updating
                Exit Function 'Increase efficiency and stop searching, assuming there is only 1 Bibliography References Table
            End If
        Next fld
        p = n
SkipLoop:
    Next T 'Many bibliographies are followed by an appendix with many tables, so it's not obvious that the Bibliography table is in the latter half of ActiveDocument.Tables.
    'Also, For-In is faster than For-To-Step in VBA, so for both reasons it makes sense to search through the tables with For-In as opposed to from the end.

    MsgBox "No Bibliography Found"
    Application.ScreenUpdating = True 'Re-enable screen updating
End Function

Sub UpdateTablesOfFiguresAndContents()
    Application.ScreenUpdating = False 'This improves performance

    Dim ToFs As TablesOfFigures: Dim ToCs As TablesOfContents: Dim Paras As Paragraphs
    Set ToFs = ActiveDocument.TablesOfFigures: Set ToCs = ActiveDocument.TablesOfContents

    Dim p, n, Change, i, j As Integer  'p = #pages, n = new #pages, Change = change in #page, i = #Loop iterations
    n = ActiveDocument.ComputeStatistics(wdStatisticPages)
    i = 0: Change = 0

    Do 'The Do-Until Loop is in case of potentially spilling ToCs and ToFs.
        i = i + 1: p = n: j = 0

        For Each ToF In ToFs
            ToF.Update
        Next ToF

        For Each ToC In ToCs
            ToC.Update 'Update first, because it resets the indentation
            
            j = j + 1: Set Paras = ToC.Range.Paragraphs
            If j = 1 Then '1st ToC is a special case
                For Each para In Paras
                    para.LeftIndent = (Val(Right(para.Style, 1)) - 1) * 21
                Next para
            Else 'Indent all to the left except in the 1st ToC
                For Each para In Paras
                    para.LeftIndent = (Val(Right(para.Style, 1)) - 1) * 21 - 20
                Next para
            End If
        Next ToC

        n = ActiveDocument.ComputeStatistics(wdStatisticPages)
        Change = Change + n - p 'postive means increase, negative means decrease
    Loop Until p = n

    If Change > 0 Then
        MsgBox "# iterations: " & i & vbCrLf & "# pages increased: " & Change
    ElseIf Change < 0 Then
        MsgBox "# iterations: " & i & vbCrLf & "# pages decreased: " & -1 * Change 'Abs(Change)
    End If 'No need to MsgBox if Change = 0 because this is typically the case
    
    Application.ScreenUpdating = True 'Re-enable screen updating
End Sub

Sub PasteAsText() '(Ctrl+Shift+V)
    On Error Resume Next 'Prevent an error and simply do nothing in case of an empty clipboard or image
    Selection.PasteAndFormat (wdFormatPlainText) 'Selection.PasteSpecial DataType:=wdPasteText, but faster
End Sub

Sub ShowAllHeadingsInNavigationPane()
    ActiveWindow.DocumentMap = True
End Sub

Sub TodaysDate()
    MsgBox "Today's date is: " & Format(Date, "dddd, mmmm d, yyyy"), vbInformation, "Today's Date"
End Sub 'e.g. Thursday, October 24, 2024

Sub CountToCs() '#Tables of Contents
    MsgBox "Number of Tables of Contents: " & ActiveDocument.TablesOfContents.Count
End Sub

Sub CountToFs() '#Tables of Figures
    MsgBox "Number of Tables of Figures: " & ActiveDocument.TablesOfFigures.Count
End Sub

Sub CountTables() '#Tables, excluding ToCs & ToFs, but includes Bibliography
    MsgBox "Number of Tables: " & ActiveDocument.Tables.Count
End Sub

Sub CountFields() 'Including field codes, but not only
    MsgBox "Number of fields: " & ActiveDocument.Fields.Count
End Sub
