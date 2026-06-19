' VBA Code Used in Workbook

Sub ToggleSection()
'   
    Dim MenuRange As Range
    Dim SubSections As Variant
    Dim n As Variant

    ' Pause Screen Updating
    Application.ScreenUpdating = False


    SubSections = Array("Data_Description", "Composition", "Summary_Stats", "Percentiles", "Field_Lists", "Compiled_Lists")

    'Pull section name from name of selected shape
    SectionName = ActiveSheet.Shapes(Application.Caller).name


    'Toggle section visibility
    Hidden = Range(Replace(SectionName, " ", "_")).EntireRow.Hidden
    Range(Replace(SectionName, " ", "_")).EntireRow.Hidden = Not Hidden


    ' If section is "Field Analysis", collapse subsections and set menu icons
    If SectionName = "Field Analysis" Then

        'Set toggles for all sections
        Range("Toggles").Orientation = 0
        ' Range("TopToggle").Orientation = -90
        Range("TopToggle").Orientation = (90 * Hidden)
        
        For Each n In SubSections
            Range(n).EntireRow.Hidden = True
            Range(n).Offset(-2, -2).Orientation = 0
        Next n
        
    Else
        
        'Set menu icon orientation for selected section
        ActiveSheet.Shapes(SectionName).TopLeftCell.Offset(0, -1).Orientation = (90 * Hidden)

    End if

    ' Restore Screen Updating
    Application.ScreenUpdating = True


End Sub


Function PythonDict(rng As Range) As String
    Dim cell As Range
    Dim dictStr As String
    
    dictStr = "{"
    For Each cell In rng.Columns(1).Cells
        ' Format as "Key": "Value" with proper escaping
        dictStr = dictStr & """" & cell.Value & """: """ & cell.Offset(0, 1).Value & """, "
    Next cell
    
    ' Remove the trailing comma and space, then close the dictionary
    If Right(dictStr, 2) = ", " Then
        dictStr = Left(dictStr, Len(dictStr) - 2)
    End If
    dictStr = dictStr & "}"
    
    PythonDict = dictStr
End Function



Function PythonList(rng As Range) As String
    Dim cell As Range
    Dim result As String
    Dim val As String
    
    ' Loop through each cell in the provided range
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            val = CStr(cell.Value)
            
            ' Format strings with quotes, leave numbers/booleans unquoted
            If Not IsNumeric(val) Then
                result = result & "'" & val & "', "
            Else
                result = result & val & ", "
            End If
        End If
    Next cell
    
    ' Clean up the trailing comma and wrap in brackets
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 2)
        PythonList = "[" & result & "]"
    Else
        PythonList = "[]"
    End If
End Function



Sub ChangePivotSource()

    Dim pt As PivotTable
    Dim SourceTable As ListObject
    Dim i As Long, j As Long
    Dim FieldName As String
    Dim AddedFieldsCount As Long
    Dim pf As PivotField
    Dim ExcludedList As Variant
    
    ' Configure Variables
    Set SourceTable = ThisWorkbook.Worksheets(Range("SelectedSheet").Value).ListObjects(Range("SelectedTable").Value)
    Set pt = ThisWorkbook.Worksheets("Pivot").PivotTables("pvt_Pivot")
    ExcludedList = Array("Record List", "index", "HasBlank", "Record Hash")
    
    ' Pause Screen Updating
    Application.ScreenUpdating = False
    pt.ManualUpdate = True
    
    ' Change Source Data using R1C1 for Mac/Win
    pt.SourceData = SourceTable.Range.Address(True, True, xlR1C1, True)
    pt.PivotCache.Refresh
    
    
    ' Clear Existing Slicers/Filters/Row Fields
    pt.ClearAllFilters
    For i = pt.RowFields.Count To 1 Step -1
        pt.RowFields(i).Orientation = xlHidden
    Next i
    
    
    ' --------------------------------------------------------------
    ' Add up to 10 Fields
    
    AddedFieldsCount = 0
    
    For i = 1 To SourceTable.ListColumns.Count
        If AddedFieldsCount = 10 Then Exit For
        
        FieldName = SourceTable.ListColumns(i).Name

        If Not IsNumeric(Application.Match(FieldName, ExcludedList, 0)) Then
        
            ' Add extra error pass for the pivot field
            On Error Resume Next
            Set pf = pt.PivotFields(FieldName)
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
                Set pf = Nothing
            End If
        End If
    Next i
    
    ' Safely Collapse Fields
    If pt.RowFields.Count > 1 Then
        For i = 1 To pt.RowFields.Count - 1
            pt.RowFields(i).ShowDetail = False
        Next i
    End If
    
    ' Push updates to table
    pt.ManualUpdate = False
    pt.RefreshTable
    
    ' Allow Mac UI thread to register layout changes before sizing
    DoEvents 
    
    ' ------------------------------------------------
    ' Format pivot table
    
    ' Format column widths/header Row/left column
    pt.TableRange1.ColumnWidth = 20
    
    With pt.TableRange1.Rows(1)
        .RowHeight = 30
        .WrapText = True
        .VerticalAlignment = xlCenter ' Bonus: keeps text centered neatly in the 30pt height
    End With
    
    With pt.TableRange1.Columns(1)
        .HorizontalAlignment = xlLeft
        .AutoFit
    End With
    
    ' Restore Screen Updating
    Application.ScreenUpdating = True
    
End Sub


Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)

    ' Static variables hold their value in memory between slicer clicks
    Static OldTable As String
    Dim NewTable As String
    
    ' Check current value of the named range changed by the slicer
    NewTable = Me.Range("SelectedSheet").Value
    OldTable = ThisWorkbook.Worksheets("Pivot").Range("Name").Value
    
    
    ' Only run if the table name has actually changed to prevent endless loops
    If NewTable <> OldTable Then
        
        ' Temporarily disable events to prevent the macro from re-triggering this event
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        
        ' Run your main pivot table update macro
        Call ChangePivotSource
        
        ' Update the static memory to the new table name
        ThisWorkbook.Worksheets("Pivot").Range("Name").Value = NewTable
                
        ' Restore Excel event and calculation settings
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True

        
    End If
End Sub





Private Sub Worksheet_Activate()
    Dim PT As PivotTable
    
    ' Temporarily turn off events to prevent any calculation or selection loops
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    ' 1. Run your metadata verification sync macro
    Call SyncTableWorksheets
    
    ' 2. Safely find and refresh the selector pivot table
    On Error Resume Next
    Set PT = ThisWorkbook.Worksheets("meta").PivotTables("pvt_TableSelector")
    On Error GoTo 0
    
    If Not PT Is Nothing Then
        PT.RefreshTable
    End If
    
    ' Turn events back on
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


