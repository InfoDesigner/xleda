Sub ChangePivotSource()
    
    ' Changes the source table of a pivot table and places up to 5 slicers for the new source    
    Application.ScreenUpdating = False
    
    Call ConfigurePivot
    
    ' Force Excel to pause and process background data calculations quietly in memory
    DoEvents
    Application.Wait (Now + TimeValue("00:00:01"))
    DoEvents
    
    Call ConfigureSlicers
    
    Application.ScreenUpdating = True
    
End Sub

Function TableExists(sheetName As String, tableName As String) As Boolean

    ' Checks whether a table exists

    Dim ws As Worksheet, t As ListObject
    
    Set ws = Nothing
    Set t = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then Set t = ws.ListObjects(tableName)
    On Error GoTo 0
    TableExists = Not t Is Nothing
End Function

Sub SyncTableWorksheets()

    ' Validates the current worksheet name for a list of tables

    Dim metaTable As ListObject, tblRow As ListRow
    Dim ws As Worksheet, targetTable As ListObject
    Dim tableName As String, sheetName As String
    Dim i As Long, foundInWorkbook As Boolean
    Set metaTable = ThisWorkbook.Worksheets("meta").ListObjects("tbl_Tables")
    For i = metaTable.ListRows.Count To 1 Step -1
        Set tblRow = metaTable.ListRows(i)
        tableName = Intersect(tblRow.Range, metaTable.ListColumns("TableName").Range).Value
        sheetName = Intersect(tblRow.Range, metaTable.ListColumns("WorksheetName").Range).Value
        If Not TableExists(sheetName, tableName) Then
            foundInWorkbook = False
            For Each ws In ThisWorkbook.Worksheets
                If TableExists(ws.Name, tableName) Then
                    Intersect(tblRow.Range, metaTable.ListColumns("WorksheetName").Range).Value = ws.Name
                    foundInWorkbook = True
                    Exit For
                End If
            Next ws
            If Not foundInWorkbook Then tblRow.Delete
        End If
    Next i
End Sub

Sub ConfigurePivot()
    
    ' Configures a pivot table to use a new source

    Dim PT As PivotTable
    Dim SourceTable As ListObject
    Dim i As Long, j As Long
    Dim FieldName As String
    Dim AddedFieldsCount As Long
    Dim pf As PivotField
    Dim ExcludedList As Variant
    Dim newCache As PivotCache
    
    
    ' Set variables
    Set SourceTable = ThisWorkbook.Worksheets(Range("SelectedSheet").Value).ListObjects(Range("SelectedTable").Value)
    Set PT = ThisWorkbook.Worksheets("Pivot").PivotTables("pvt_Pivot")
    ExcludedList = Array("Record List", "index", "HasBlank", "Record Hash")
    

    ' Change pivot source data
    Set newCache = ThisWorkbook.PivotCaches.Create(xlDatabase, SourceTable.Range.Address(True, True, xlR1C1, True))
    PT.ChangePivotCache newCache
    PT.PivotCache.Refresh
    
    ' Remove all pivot fields and filters
    PT.ClearAllFilters
    For i = PT.RowFields.Count To 1 Step -1
        PT.RowFields(i).Orientation = xlHidden
    Next i
    
    PT.ManualUpdate = True
    
    
    ' Add up to 10 row fields with no subtotals
    AddedFieldsCount = 0
    For i = 1 To SourceTable.ListColumns.Count
        If AddedFieldsCount = 10 Then Exit For
        FieldName = SourceTable.ListColumns(i).Name
        If IsError(Application.Match(FieldName, ExcludedList, 0)) Then
            Set pf = Nothing
            On Error Resume Next
            Set pf = PT.PivotFields(FieldName)
            On Error GoTo 0
            If Not pf Is Nothing Then
                AddedFieldsCount = AddedFieldsCount + 1
                With pf
                    .Orientation = xlRowField
                    .Position = AddedFieldsCount
                    For j = 1 To 12
                        .Subtotals(j) = False
                    Next j
                End With
            End If
        End If
    Next i
    
    ' Reset pivot table settings
    PT.ManualUpdate = False
    PT.Update
    PT.RefreshTable
    
    
    'Collapse all fields
    If PT.RowFields.Count > 1 Then
        For i = 1 To PT.RowFields.Count - 1
            PT.RowFields(i).ShowDetail = False
        Next i
    End If
    
End Sub


Sub DeleteOtherSlicers()
    
    ' Deletes all slicers except the one named 'TableSlicer'

    Dim sCache As SlicerCache
    Dim sItem As Slicer
    Dim shp As Shape
    
    ' 1. Delete all slicers on the active sheet except "TableSlicer"
    For Each sCache In ActiveWorkbook.SlicerCaches
        For Each sItem In sCache.Slicers
            If sItem.Shape.Parent Is ActiveSheet And sItem.Name <> "TableSlicer" Then
                sItem.Delete
            End If
        Next sItem
    Next sCache
    
    ' 2. Delete shape groups with "group" in their name (case-insensitive)
    For Each shp In ActiveSheet.Shapes
        If shp.Type = msoGroup Then
            If InStr(1, shp.Name, "group", vbTextCompare) > 0 Then
                shp.Delete
            End If
        End If
    Next shp
End Sub




Sub ConfigureSlicers()

    'Adds the HasBlank and first 4 columns of a data source to a pivot table

    Dim PT As PivotTable
    Dim SourceTable As ListObject
    Dim PivotWS As Worksheet
    Dim StartCell As Range
    Dim TargetRange As Range
    Dim SC As SlicerCache
    Dim FieldName As String
    Dim i As Long
    Dim sSize As Double: sSize = 144
    Dim CurrentLeft As Double
    Dim CurrentTop As Double
    Dim Gap As Double: Gap = 10
    Dim SlicerNames() As String
    Dim SlicerCount As Long: SlicerCount = 0
    Dim WasHidden As Boolean
    Dim ProcessingFields() As String
    Dim TargetFieldCount As Long: TargetFieldCount = 0
    Dim HasBlankIndex As Long: HasBlankIndex = 0
    
    ' Initialize variables
    Set SourceTable = ThisWorkbook.Worksheets(Range("SelectedSheet").Value).ListObjects(Range("SelectedTable").Value)
    Set PT = ThisWorkbook.Worksheets("Pivot").PivotTables("pvt_Pivot")
    Set PivotWS = ThisWorkbook.Worksheets("Pivot")
    Set StartCell = ThisWorkbook.Names("PivotFiltersStart").RefersToRange
    Set TargetRange = ThisWorkbook.Names("PivotTableFilters").RefersToRange
    
    
    ' Start a collection of fields to add
    For i = 1 To SourceTable.ListColumns.Count
        If SourceTable.ListColumns(i).Name = "HasBlank" Then
            TargetFieldCount = TargetFieldCount + 1
            ReDim Preserve ProcessingFields(1 To TargetFieldCount)
            ProcessingFields(TargetFieldCount) = "HasBlank"
            Exit For
        End If
    Next i
    
    
    ' Refine the fields to add
    For i = 1 To SourceTable.ListColumns.Count
        FieldName = SourceTable.ListColumns(i).Name
        If FieldName <> "index" And FieldName <> "Record Hash" And FieldName <> "HasBlank" Then
            If TargetFieldCount < 5 Then
                TargetFieldCount = TargetFieldCount + 1
                ReDim Preserve ProcessingFields(1 To TargetFieldCount)
                ProcessingFields(TargetFieldCount) = FieldName
            Else
                Exit For
            End If
        End If
    Next i
    
    
    ' Capture whether the pivot filters are hidden and unhide them
    WasHidden = TargetRange.Rows(1).EntireRow.Hidden
    If WasHidden Then
        TargetRange.EntireRow.Hidden = False
    End If
    
    ' Anchor for slicer placement
    CurrentLeft = StartCell.Left
    CurrentTop = StartCell.Top
    
    ' Delete all slicers from previous table
    Call DeleteOtherSlicers
    
    
    ' Loop through ProcessingFields and add the expected slicers
    For i = 1 To TargetFieldCount
        FieldName = ProcessingFields(i)
       
        Set SC = ThisWorkbook.SlicerCaches.Add2(PT, FieldName)
        SC.Slicers.Add PivotWS, , , FieldName
        
        ' Note: This remains because it passes background object references to Windows/Mac
        ' graphic memory, but since screen updating is dead-locked, nothing paints on screen.
        DoEvents
        
        ' Add the slicer
        With PivotWS.Shapes(FieldName)
            .Top = CurrentTop
            .Left = CurrentLeft
            .Width = sSize
            .Height = sSize
            .Placement = xlMoveAndSize
        End With
        
        ' Increment anchor placement
        SlicerCount = SlicerCount + 1
        ReDim Preserve SlicerNames(1 To SlicerCount)
        SlicerNames(SlicerCount) = FieldName
        CurrentLeft = CurrentLeft + sSize + Gap
    Next i
    
    ' Group the added slicers and set xlMoveAndSize
    If SlicerCount > 1 Then
        Dim ShpGroup As Shape
        On Error Resume Next
        Set ShpGroup = PivotWS.Shapes.Range(SlicerNames).Group
        If Not ShpGroup Is Nothing Then
            ShpGroup.Placement = xlMoveAndSize
        End If
        On Error GoTo 0
    End If
    
    ' Restore the Pivot filter range
    If WasHidden Then
        TargetRange.EntireRow.Hidden = True
    End If
    
    ' Set pivot table formatting for header/left column
    PT.TableRange1.ColumnWidth = 15
    With PT.TableRange1.Rows(1)
        .RowHeight = 30
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    With PT.TableRange1.Columns(1)
        .HorizontalAlignment = xlLeft
    End With
    
    
End Sub