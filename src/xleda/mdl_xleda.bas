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



Function PythonList(rng As Variant) As String
    Dim Element As Variant
    Dim result As String

    If TypeName(rng) = "Range" Then
        For Each Element In rng.Cells
            If Not IsEmpty(Element.Value) Then result = result & PythonListFormat(Element.Value) & ", "
        Next
    ElseIf IsArray(rng) Then
        For Each Element In rng
            If Not IsEmpty(Element) Then result = result & PythonListFormat(Element) & ", "
        Next
    Else
        If Not IsEmpty(rng) Then result = result & PythonListFormat(rng) & ", "
    End If

    If Len(result) > 0 Then result = Left(result, Len(result) - 2)
    PythonList = "[" & result & "]"
End Function



Private Function PythonListFormat(v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        PythonListFormat = "''"
        Exit Function
    End If

    Select Case VarType(v)
        Case vbBoolean
            PythonListFormat = IIf(v, "True", "False")
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            PythonListFormat = CStr(v)
        Case Else
            PythonListFormat = "'" & Replace(CStr(v), "'", "\'") & "'"
    End Select
End Function