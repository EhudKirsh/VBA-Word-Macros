In Microsoft Word, to update all references and cross references, typically people use Ctrl + A to select all and then press F9. This however pops up an annoying window that maunally needs to be clicked. It's stupid, annoying and time consuming. This window pops up for every Table of Content (TOC) and every Table of caption labels.

The purpose of this VBA Macro is to update all of these automatically at a single press of a button without a single window like this poping up. Specifically, it updates all of the following:

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

    ' Update all Table of Figures
    Dim fig As TableOfFigures
    For Each fig In ActiveDocument.TablesOfFigures
        fig.Update
    Next fig
    
    ' Update all Table of Contents (TOCs)
    Dim toc As TableOfContents
    For Each toc In ActiveDocument.TablesOfContents
        toc.Update
    Next toc
    'Keep TOCs last, because the others might change page numbers

    ' Re-enable screen updating and events
    Application.ScreenUpdating = True
End Sub

```
