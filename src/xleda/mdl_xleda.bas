' VBA Code Used in Workbook

Sub ToggleSection()
'   
    Dim MenuRange As Range
    Dim SubSections As Variant
    Dim n As Variant


    SubSections = Array("Field_Notes", "Composition", "Summary_Stats", "Percentiles", "Field_Lists", "Compiled_Lists")

    'Pull section name from name of selected shape
    SectionName = ActiveSheet.Shapes(Application.Caller).name


    'Toggle section visibility
    Hidden = Range(Replace(SectionName, " ", "_")).EntireRow.Hidden
    Range(Replace(SectionName, " ", "_")).EntireRow.Hidden = Not Hidden


    'Rotate menu icon
    Set MenuRange = ActiveSheet.Shapes(SectionName).TopLeftCell.Offset(0, -1)
    MenuRange.Orientation = (90 * Hidden)

    ' If section is "Field Analysis", collapse subsections and set menu icons
    If SectionName = "Field Analysis" Then
        
        For Each n In SubSections
            Range(n).EntireRow.Hidden = True
            Range(n).Offset(-2, -2).Orientation = 0
        Next n

    End if

End Sub