In Microsoft Word, to update all references and cross references, typically people use Ctrl+A to select all and then press F9. This however pops up an annoying window that maunally needs to be clicked. It's stupid, annoying and wastes time. This window pops up for every Table of Contents and Table of Figures.
<!-- To Do: Add 2 screenshots here: TOC & Table of Figures -->
The purpose of this VBA Macro is to update all of these automatically with a single click without any window like this poping up. Specifically, it updates all of the following:

- All Caption Labels (equations, figures and tables by default, but also custom)
- All Cross-References, including of the above caption labels, but also of headings
- All References
- All Tables of Contents, Tables of Figures and Bibliography References Table
```VBA
Sub UpdateTablesOfFiguresAndContents()
    Dim ToFs As TablesOfFigures: Dim ToCs As TablesOfContents: Dim Paras As Paragraphs
    Set ToFs = ActiveDocument.TablesOfFigures: Set ToCs = ActiveDocument.TablesOfContents

    Dim p, n, Change, i As Integer  'p = #pages, n = new #pages, Change = change in #page, i = #Loop iterations
    n = ActiveDocument.ComputeStatistics(wdStatisticPages) '#Pages
    i = 0: Change = 0

    Do 'The Do-Until Loop is in case of potentially spilling ToCs and ToFs.
        i = i + 1: p = n

        For Each ToF In ToFs
            ToF.Update
        Next ToF

        For Each ToC In ToCs
            ToC.Update
            Set Paras = ToC.Range.Paragraphs
            For Each para In Paras
                para.LeftIndent = (Val(Right(para.Style, 1)) - 1) * 21
            Next para
        Next ToC

        n = ActiveDocument.ComputeStatistics(wdStatisticPages)
        Change = Change + n - p 'postive means increase, negative means decrease
    Loop Until p = n

    If Change > 0 Then
        MsgBox "# iterations: " & i & vbCrLf & "# pages increased: " & Change
    ElseIf Change < 0 Then
        MsgBox "# iterations: " & i & vbCrLf & "# pages decreased: " & -1 * Change 'Abs(Change)
    End If 'No need to MsgBox if Change = 0 because this is typically the case
End Sub

Sub StyleBibliography()
    'Style the Bibliography References Table: turn https into hyperlinks, adjust columns widths and align text to the left
    Dim T As Table
    Dim f As field
    Dim FieldsCount, httpsPos, spacePos As Integer

    FieldsCount = ActiveDocument.Fields.Count
    For i = FieldsCount To 1 Step -1 'Searching from the end, because the bibliography is most likely in the latter half of the document
        Set f = ActiveDocument.Fields(i)
        If f.Type = wdFieldBibliography Then 'Find the bibliography
            Dim cols, C2 As Object
            Set T = f.Result.Tables(1): T.AllowAutoFit = False
            Set cols = T.columns: Set C2 = cols(2)

            'Optional - pick how many digits of references you have:
            'cols(1).Width = 17 ': C2.Width = 420 '[9]
            'cols(1).Width = 22 ': C2.Width = 415 '[99]
            cols(1).Width = 30 ': C2.Width = 407 '[999]

            C2.AutoFit 'Width

            Dim CellsRange As Cells: Set CellsRange = C2.Cells

            Dim r As Range
            Dim cellText, linkText As String

            For Each c In CellsRange
                Set r = c.Range
                r.ParagraphFormat.Alignment = wdAlignParagraphLeft

                cellText = r.Text: cellText = Left(cellText, Len(cellText) - 2)
                httpsPos = InStr(cellText, "https")
                If httpsPos > 0 Then
                    'Find the first space after "https"
                    spacePos = InStr(httpsPos, cellText, " ")
                    If spacePos = 0 Then spacePos = Len(cellText) + 1 'Use text length if no space is found

                    'Extract the link text (URL)
                    linkText = Mid(cellText, httpsPos, spacePos - httpsPos - 1) 'Assuming there's a dot '.' just before thespace ' '

                    r.Start = r.Start + httpsPos - 1 'Assuming there's a dot '.' just before thespace ' '
                    r.End = r.Start + Len(linkText)

                    'Insert the hyperlink
                    ActiveDocument.Hyperlinks.Add Anchor:=r, Address:=linkText
                End If
            Next c

            Exit For 'Increase efficiency and stop searching, assume only 1 bibliography
        End If
    Next i
End Sub

Sub UpdateAll()
    Application.ScreenUpdating = False 'This improves performance

    'Firstly hide the field codes so Word doesn't need to update their display
    ActiveWindow.View.ShowFieldCodes = False 'Alt + F9

    'Update all fields in the document, including references, cross references & caption labels
    ActiveDocument.Fields.Update
    'This has to be before StyleBibliography because it resets the style and can add rows

    StyleBibliography 'This has to be before UpdateTablesOfFiguresAndContents, because the bibliography can spill into more pages, so we start from the end

    UpdateTablesOfFiguresAndContents

    Application.ScreenUpdating = True 'Re-enable screen updating
End Sub
```
<ins>Steps to put this macro in Word:</ins>
1) Open the 'Microsoft Visual Basic for Applications' window by pressing Alt+F11 OR click on 'Developer' tab -> 'Visual Basic' under 'Code' -> click on 'Insert' -> 'Module' -> copy the above program and paste into this module -> Save by pressing Ctrl+S OR Click on the Save icon 💾
2) Run it to check it works by pressing F5 OR clicking on the Run icon ▷
3) Add this macro to the Quick Access Toolbar: click on 'File' -> 'Options' -> 'Quick Access Toolbar' -> '<ins>C</ins>hoose commands from:' -> 'Macros' -> click on the macro you created -> '<ins>A</ins>dd >>' -> click on this macro that you just added to the right -> '<ins>M</ins>odify...' -> Pick a nice Display name and icon, I like 'UpdateAll' and the update document symbol 📄🔄 -> OK & OK
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
More small interesting and useful macros:
```VBA
Sub TodaysDate()
    MsgBox "Today's date is: " & Format(Date, "dddd, mmmm d, yyyy"), vbInformation, "Today's Date"
End Sub 'e.g. Thursday, October 24, 2024
```
```VBA
Sub ShowAllHeadingsInNavigationPane()
    ActiveWindow.DocumentMap = True
End Sub
```
```VBA
Sub CountFields() 'Including field codes, but not only
    MsgBox "Number of fields: " & ActiveDocument.Fields.Count
End Sub
```
```VBA
Sub CountToCs() '#Tables of Contents
    MsgBox "Number of Tables of Contents: " & ActiveDocument.TablesOfContents.Count
End Sub
```
```VBA
Sub CountToFs() '#Tables of Figures
    MsgBox "Number of Tables of Figures: " & ActiveDocument.TablesOfFigures.Count
End Sub
```
```VBA
Sub CountTables() '#Tables, excluding ToCs & ToFs, but includes Bibliography
    MsgBox "Number of Tables: " & ActiveDocument.Tables.Count
End Sub
```
