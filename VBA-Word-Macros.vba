Sub UpdateAll()
    ' Disable screen updating and events for performance
    Application.ScreenUpdating = False
    
    ' Update all fields in the document (includes captions)
    ActiveDocument.Fields.Update
    'This has to be first to re-index captions correctly later in the tables
'----------------------------------------------------------------------------

    'Update Bibliography References Table Style. This has to be before the TOCs, because resizing the Bibliography stops spilling into more pages
    Dim T As Table
    Dim F As Field
    Dim FieldsCount As Long
    FieldsCount = ActiveDocument.Fields.Count
    For i = FieldsCount To 1 Step -1 'Searching from the end, because the bibliography is most likely in the latter half of the document
        Set F = ActiveDocument.Fields(i)
        If F.Type = wdFieldBibliography Then 'Find the bibliography
            Dim cols, C2 As Object
            Set cols = F.Result.Tables(1).columns

            'Optional - pick how many digits of references you have:
            'cols(1).Width = 17 '[9]
            'cols(1).Width = 22 '[99]
            cols(1).Width = 30 '[999]

            Set C2 = cols(2)
            C2.AutoFit 'Width

            Dim CellsRange As Cells
            Set CellsRange = C2.Cells

            Dim c As Cell
            For Each c In CellsRange
                c.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            Next c

            Exit For 'Increase efficiency and stop searching, assume only 1 bibliography
        End If
    Next i

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
