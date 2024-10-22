Sub UpdateAll()
    ' Disable screen updating and events for performance
    Application.ScreenUpdating = False
    
'-------------------------------------------------------------------------------------------------

    'Firstly hide the field codes so Word doesn't need to update their display
    ActiveWindow.View.ShowFieldCodes = False
    
    ' Update all fields in the document, including references, cross references & caption labels
    ActiveDocument.Fields.Update
    'This has to be here at the start before updating TOCs & TOFs to re-index captions correctly later in the tables

'-------------------------------------------------------------------------------------------------

    'Update Bibliography References Table Style. This has to be before the TOCs, it can spill into more pages, so we start from the end
    Dim T As Table
    Dim F As field
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

'-------------------------------------------------------------------------------------------------

    ' Update all Tables of Figures & Contents
    Dim ToF As TableOfFigures
    For Each ToF In ActiveDocument.TablesOfFigures
        ToF.Update
    Next ToF

    Dim toc As TableOfContents
    For Each toc In ActiveDocument.TablesOfContents
        toc.Update
    Next toc
    
'-------------------------------------------------------------------------------------------------
    
    ' Re-enable screen updating and events
    Application.ScreenUpdating = True
End Sub



Sub PasteAsText() '(Ctrl+Shift+V)
    On Error Resume Next 'Prevent an error and simply do nothing in case of an empty clipboard or image
    Selection.PasteAndFormat (wdFormatPlainText) 'Selection.PasteSpecial DataType:=wdPasteText, but faster
End Sub
