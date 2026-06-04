' VBA Code Used in Workbook

Sub ToggleSection()
'   
    Dim MenuRange As Range
    Dim SubSections As Variant
    Dim n As Variant


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