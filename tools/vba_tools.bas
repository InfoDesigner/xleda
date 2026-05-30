Sub MoveNamedRange()
    
    ' Set your variables
    Set sourceWB = Workbooks("xleda_format_reference.xlsm")
    Set targetWB = Workbooks("template.xlsm")
    
    On Error Resume Next

    For Each nm In sourceWB.Names
        
        targetWB.Names.Add name:=nm.name, RefersTo:=nm.value

    Next
       
End Sub


Sub ListHiddenNames()
    Dim nm As name
    Dim ws As Worksheet
    Dim rowCount As Long
    
    ' Create a new sheet to hold the list
    Set ws = Worksheets.Add
    ws.name = "Hidden_Names_List"
    
    ' Set headers
    ws.Range("A1:C1").Value = Array("Name", "Refers To", "Visible?")
    rowCount = 2
    
    ' Loop through all names in the workbook
    For Each nm In ThisWorkbook.Names
        If nm.Visible = False Then
            
            ws.Cells(rowCount, 1).Value = nm.name
            ws.Cells(rowCount, 2).Value = "'" & nm.RefersTo ' Leading quote prevents formula execution
            ws.Cells(rowCount, 3).Value = "Hidden"
            rowCount = rowCount + 1
        End If
    Next nm
    
    MsgBox "Found " & rowCount - 2 & " hidden names.", vbInformation
End Sub



Sub DeleteHiddenNames()
    Dim xName As name
    Dim counter As Long
    counter = 0
    
    ' Loop through all names in the workbook
    For Each xName In ActiveWorkbook.Names
        If xName.Visible = False Then
            On Error Resume Next ' Skip names that can't be deleted (e.g., corrupted or internal)
            xName.Delete
            On Error GoTo 0
            counter = counter + 1
        End If
    Next xName
    
    MsgBox counter & " hidden names were deleted.", vbInformation
End Sub
