In Microsoft Word, to update all references and cross references, typically people use Ctrl+A to select all and then press F9. This however pops up an annoying window that maunally needs to be clicked. It's stupid, annoying and wastes time. This window pops up for every Table of Contents and Table of Figures.
<!-- To Do: Add 2 screenshots here: TOC & Table of Figures -->
The purpose of this VBA Macro is to update all of these automatically with a single click without any window like this poping up. Specifically, it updates all of the following:

- Caption Labels - equations, figures and tables by default, but also custom
- Cross-References, including of the above caption labels, but also of headings
- Citations [#]
- Tables of Figures (of Caption Labels)
- Tables of Contents, with custom indentations for different levels of headings
- Bibliography References Table, aligning text to the left, fit widths of cols, HyperLink URLs 🔗, and display typically hidden details like DOI
```VBA
Sub UpdateAll()
    Application.ScreenUpdating = False 'This improves performance

    'Firstly hide the field codes so Word doesn't need to update their display
    ActiveWindow.View.ShowFieldCodes = False 'AltLeft + F9

    'Update all fields in the document, including references, cross references & caption labels
    ActiveDocument.Fields.Update
    'This has to be before StyleBibliography because it resets the style and can also add rows

    StyleBibliography 'This has to be before UpdateTablesOfFiguresAndContents, because the bibliography can spill into more pages, so we start from the end

    UpdateTablesOfFiguresAndContents

    Application.ScreenUpdating = True 'Re-enable screen updating
End Sub

Sub StyleBibliography()
    Application.ScreenUpdating = False 'This improves performance

    'Style the Bibliography References Table: turn http into hyperlinks, adjust columns widths and align text to the left
    Dim T As Table: Set T = FindBibliography: T.AllowAutoFit = False
    Dim httpPos, spacePos, refs As Integer: Dim cols, C2 As Object
    Set cols = T.columns: Set C2 = cols(2)

    refs = ActiveDocument.Bibliography.Sources.Count
    'Width of 1st col based on how many digits of references you have:
    If refs <= 9 Then '[9]
        cols(1).Width = 17 ': C2.Width = 420
    ElseIf refs <= 99 Then '[99]
        cols(1).Width = 22 ': C2.Width = 415
    Else '[999]
        cols(1).Width = 30 ': C2.Width = 407
    End If
    C2.AutoFit 'Width

    Dim CellsRange As Cells: Set CellsRange = C2.Cells
    Dim r As Range: Dim cellText, linkText As String
    For Each c In CellsRange
        Set r = c.Range

        r.ParagraphFormat.Alignment = wdAlignParagraphLeft 'Align Left

        'Hyperlinks
        cellText = r.Text: cellText = Left(cellText, Len(cellText) - 2)
        httpPos = InStr(cellText, "http") 'some links don't have the 's' in 'https', but 'http' works for both
        If httpPos > 0 Then
            spacePos = InStr(httpPos, cellText, " ") 'Find the first space after "http"
            If spacePos = 0 Then spacePos = Len(cellText) + 1 'Use text length if no space is found

            'Extract the link text (URL)
            linkText = Mid(cellText, httpPos, spacePos - httpPos - 1) 'Assuming there's a dot '.' just before thespace ' '

            r.Start = r.Start + httpPos - 1 'Assuming there's a dot '.' just before thespace ' '
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
```
<ins>Steps to use these macros in Word:</ins>
1) Open the 'Microsoft Visual Basic for Applications' window by pressing AltLeft + F11 OR click on 'Developer' tab & 'Visual Basic' under 'Code' -> click on 'Insert' -> 'Module' -> copy the above macros and paste into this module -> Save by pressing Ctrl+S OR Click on the Save icon 💾
2) Run them to check they work by pressing F5 OR clicking on the Run icon ▷ when the caret stands on whichever macro you'd like to test
3) Add the UpdateAll macro to the Quick Access Toolbar: click on 'File' -> 'Options' -> 'Quick Access Toolbar' -> '<ins>C</ins>hoose commands from:' -> 'Macros' -> click on the macro you created -> '<ins>A</ins>dd >>' -> click on this macro that you just added to the right -> '<ins>M</ins>odify...' -> Pick a nice Display name and icon, I like 'UpdateAll' and the update document symbol 📄🔄 -> OK & OK
Now by simply clicking on this icon at the top left on your screen runs this macro every time. You can also run it with a custom hotkey sequence in the 'Customize Ribbon' tab in the 'Options' next to the 'Quick Access Toolbar', but I didn't bother with it.

<!-- To Do: Add screenshots here & link to my PhD thesis to show how the bibliography hyperlinks 🔗 and ToC indentations look like -->
---
<ins>Paste as text:</ins>
```VBA
Sub PasteAsText() '(Ctrl+Shift+V)
    On Error Resume Next 'Prevent an error and simply do nothing in case of an empty clipboard or image
    Selection.PasteAndFormat (wdFormatPlainText) 'Selection.PasteSpecial DataType:=wdPasteText, but faster
End Sub
```
Techinically, PasteAsText can be set to be the default paste in the Options, but it tends to not work. Also, this macro also helps to create a hot key shortcut as shown in the steps below:

1) Add this macro to the Quick Access Toolbar with the 3 steps above. I use the clipboard 📋 symbol and 'PasteAsText (Ctrl+Shift+V)' Display name.
2) Now click on 'Customize Ribbon' in the Options -> 'Keyboard shortcuts: Cus<ins>t</ins>omize...' -> <ins>C</ins>ategories: 'Macros' -> Click on the macro on the right you want to assign a shortcut hot key to, PasteAsText in this case -> Look at 'C<ins>u</ins>rrent keys:' to see what the current shortcut hot keys are for it, but if it's empty or not to your liking, record your new shortcut hot keys by clicking on 'Press <ins>n</ins>ew shortcut key:', I use Ctrl+Shift+V -> <ins>A</ins>ssign -> Close & OK
---
<ins>More small interesting and useful macros:</ins>

Counters:
```VBA
Sub CountImages()
    Dim i, f As Integer 'i = inlineImages, f = floatingImages

    i = ActiveDocument.InlineShapes.Count: f = ActiveDocument.Shapes.Count

    MsgBox "Inline Images: " & i & vbCrLf & _
           "Floating Images: " & f & vbCrLf & "Total Images: " & i + f
End Sub

Sub CountBookmarks() 'These allow forming custom TOCs for each chapter
    MsgBox "Number of Bookmarks: " & ActiveDocument.Bookmarks.Count
End Sub

Sub CountToCs() '#Tables of Contents
    MsgBox "Number of Tables of Contents: " & ActiveDocument.TablesOfContents.Count
End Sub
'Note that I have 1 bookmark for every ToC besides the main ToC, so for me: #ToCs = #Bookmarks + 1.
'For you it might be different if you use bookmarks for other purposes as well.

Sub CountToFs() '#Tables of Figures
    MsgBox "Number of Tables of Figures: " & ActiveDocument.TablesOfFigures.Count
End Sub

Sub CountCaptionLabels()
    Application.ScreenUpdating = False 'This improves performance
 
    Dim i, Total, LabelsCount As Integer
    Total = 0: LabelsCount = CaptionLabels.Count
    Dim obj: Set obj = CreateObject("Scripting.Dictionary")
    Dim Name, Msg As String: Msg = ""

    For i = 1 To LabelsCount
        obj.Add CaptionLabels(i).Name, 0
    Next

    Dim fld, flds As Fields: Set flds = ActiveDocument.Fields
    For Each fld In flds
        If fld.Type = wdFieldSequence Then
            Total = Total + 1
            Name = Trim(Split(fld.Code.Text, " ")(2))
            obj(Name) = obj(Name) + 1
        End If
    Next

    i = 0
    For Each Name In obj.keys
        i = i + 1
        Msg = Msg & i & ") " & Name & ": " & obj(Name) & vbCrLf
    Next

    MsgBox "Number of Labels: " & LabelsCount & vbCrLf & vbCrLf & _
        Msg & vbCrLf & "Total Number of Captions: " & Total

    Application.ScreenUpdating = True 'Re-enable screen updating
End Sub

Sub CountTables() '#Tables, excluding ToCs & ToFs, but includes Bibliography
    MsgBox "Number of Tables: " & ActiveDocument.Tables.Count
End Sub

Sub CountFields() 'Including field codes, but not only
    MsgBox "Number of Fields: " & ActiveDocument.Fields.Count
End Sub

Sub CountCitationsAndReferences()
    Application.ScreenUpdating = False 'This improves performance

    Dim c, r As Integer: c = 0: r = ActiveDocument.Bibliography.Sources.Count
    '#References = Length of current list of sources or updated Bibliography

    Dim flds As Fields: Set flds = ActiveDocument.Fields
    For Each fld In flds
        If fld.Type = wdFieldCitation Then
            c = c + 1 '#Citations = Occurances of citations throughout the document
        End If
    Next
    MsgBox "Number of Citations: " & c & vbCrLf & "Number of References: " & _
    r & vbCrLf & "Citations/References Ratio: " & Round(c / r, 2)

    Application.ScreenUpdating = True 'Re-enable screen updating
End Sub

Sub CountCrossReferences()
    Application.ScreenUpdating = False 'This improves performance

    Dim c As Integer: c = 0
    Dim flds As Fields: Set flds = ActiveDocument.Fields
    For Each fld In flds
        If fld.Type = wdFieldRef Then
            c = c + 1
        End If
    Next
    MsgBox "Number of Cross-References: " & c

    Application.ScreenUpdating = True 'Re-enable screen updating
End Sub

Sub CountHyperlinksURLs()
    Application.ScreenUpdating = False 'This improves performance

    Dim c As Integer: c = 0
    Dim flds As Fields: Set flds = ActiveDocument.Fields
    For Each fld In flds
        If fld.Type = wdFieldHyperlink Then
            c = c + 1
        End If
    Next
    MsgBox "Number of Hyperlinks URLs: " & c

    Application.ScreenUpdating = True 'Re-enable screen updating
End Sub
```
Note that screen updating is only disabled in subs and functions that are not instant-quick.

Non-Counters, Other MsgBox:
```VBA
Sub TodaysDate()
    MsgBox "Today's date is: " & Format(Date, "dddd, mmmm d, yyyy")
End Sub

Sub DocumentFolderPath() 'Where it's saved to
    Dim p As String: p = ActiveDocument.Path
    If p <> "" Then
        MsgBox "Document Path: " & p
    Else
        MsgBox "This document hasn't been saved yet"
    End If
End Sub
```
Non-MsgBox:
```VBA
Sub SaveDocument()
    ActiveDocument.Save
End Sub
```
ToggleShow:
```VBA
Sub ToggleShowHeadingsNavigationPane()
    If ActiveWindow.DocumentMap Then
        ActiveWindow.DocumentMap = False
    Else
        ActiveWindow.DocumentMap = True
    End If
End Sub

Sub ToggleShowFieldCodes() 'AltLeft + F9
    If ActiveWindow.View.ShowFieldCodes Then
        ActiveWindow.View.ShowFieldCodes = False
    Else
        ActiveWindow.View.ShowFieldCodes = True
    End If
End Sub
```
