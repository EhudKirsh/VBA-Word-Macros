In Microsoft Word, to update all references and cross references, typically people use Ctrl+A to select all and then press F9. This however pops up an annoying window that maunally needs to be clicked. It's stupid, annoying and wastes time. This window pops up for every Table of Content (TOC) and every Table of caption labels.
<!-- To Do: Add 2 screenshots here: TOC & Table of Figures -->
The purpose of this VBA Macro is to update all of these automatically with a single click without any window like this poping up. Specifically, it updates all of the following:

- All Caption Labels (equations, figures and tables by default, but also custom)
- All Cross-References, including of the above caption labels, but also of headings
- All References
- All TOCs, Tables of caption labels and Table of References/Bibliography
```VBA
Sub UpdateAll()
    ' Disable screen updating and events for performance
    Application.ScreenUpdating = False
    
    ' Update all fields in the document (includes captions)
    ActiveDocument.Fields.Update
    'This has to be first to re-index captions correctly later in the tables

    ' Update all Tables of Figures
    Dim fig As TableOfFigures
    For Each fig In ActiveDocument.TablesOfFigures
        fig.Update
    Next fig
    
    ' Update all Tables of Contents (TOCs)
    Dim toc As TableOfContents
    For Each toc In ActiveDocument.TablesOfContents
        toc.Update
    Next toc
    'Keep TOCs last, because the others might change page numbers

    ActiveWindow.View.ShowFieldCodes = False

    ' Re-enable screen updating and events
    Application.ScreenUpdating = True
End Sub
```
<ins>Steps to put this macro in Word:</ins>
1) Open the 'Microsoft Visual Basic for Applications' window by pressing Alt+F11 OR click on 'Developer' tab -> 'Visual Basic' under 'Code' -> click on 'Insert' -> 'Module' -> copy the above program and paste into this module -> Save by pressing Ctrl+S OR Click on the Save icon 💾
2) Run it to check it works by pressing F5 OR clicking on the Run icon ►
3) Add this macro to the Quick Access Toolbar: click on 'File' -> 'Options' -> 'Quick Access Toolbar' -> '<ins>C</ins>hoose commands from:' -> 'Macros' -> click on the macro you created -> '<ins>A</ins>dd >>' -> click on this macro that you just added to the right -> '<ins>M</ins>odify...' -> Pick a nice Display name and icon, I like 'UpdateAll' and the Run symbol ▷ -> OK & OK
<!-- To Do: Add screenshots here -->
Now by simply clicking on this icon at the top left on your screen runs this macro every time. You can also run it with a custom hotkey sequence in the 'Customize Ribbon' tab in the 'Options' next to the 'Quick Access Toolbar', but I didn't bother with it.

---
<ins>Paste as text:</ins>
```VBA
Sub PasteAsText()
    On Error GoTo ErrorHandler
    'Prevent an error and do nothing in case of an empty clipboard or image
    
    Selection.PasteAndFormat (wdFormatPlainText)
    'Selection.PasteSpecial DataType:=wdPasteText

    Exit Sub

ErrorHandler:
    Err.Clear
End Sub
```
Techinically, PasteAsText can be set to be the default paste in the Options, but it tends to not work. Also, this macro also helps to create a hot key shortcut as shown in the steps below:

1) Add this macro to the Quick Access Toolbar with the 3 steps above. I use the clipboard 📋 symbol and 'PasteAsText (Ctrl+Shift+V)' Display name.
2) Now click on 'Customize Ribbon' in the Options -> 'Keyboard shortcuts: Cus<ins>t</ins>omize...' -> <ins>C</ins>ategories: 'Macros' -> Click on the macro on the right you want to assign a shortcut hot key to, PasteAsText in this case -> Look at 'C<ins>u</ins>rrent keys:' to see what the current shortcut hot keys are for it, but if it's empty or not to your liking, record your new shortcut hot keys by clicking on 'Press <ins>n</ins>ew shortcut key:', I use Ctrl+Shift+V -> <ins>A</ins>ssign -> Close & OK
