Sub UpdateAll()
    ' Disable screen updating and events for performance
    Application.ScreenUpdating = False
    
    ' Update all fields in the document (includes captions)
    ActiveDocument.Fields.Update
    'This has to be first to re-index captions correctly later in the tables
'----------------------------------------------------------------------------

    'Update Bibliography References Table Style
    Dim F As Field
    Dim found As Boolean
    found = False

    For Each F In ActiveDocument.Fields
        If F.Type = wdFieldBibliography Then
            Dim C As Object
            Set C = F.Result.Tables(1).columns

            'Optional - pick how many digits of references you have:
            'C(1).Width = 17 '[9]
            'C(1).Width = 22 '[99]
            C(1).Width = 30 '[999]

            C(2).Width = AutoFit

            Exit For
        End If
    Next F
'----------------------------------------------------------------------------

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



Sub PasteAsText() '(Ctrl+Shift+V)
    On Error Resume Next 'Prevent an error and simply do nothing in case of an empty clipboard or image
    Selection.PasteAndFormat (wdFormatPlainText) 'Selection.PasteSpecial DataType:=wdPasteText, but faster
End Sub
