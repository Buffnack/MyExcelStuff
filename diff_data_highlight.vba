Option Explicit

Dim local_list_index As String
Dim remote_list_index As String
Dim local_table_index As String
Dim remote_table_index As String

Private Sub initWithDefault()
    
    local_list_index = "HalloWelt"
    local_table_index = "Tabelle1"
    remote_list_index = "Remote"
    remote_table_index = "Tabelle13"

End Sub

Private Sub ClearBackground_Click()

    If local_list_index = "" Then
        
        initWithDefault
    End If
    
    
    Dim local_tab As ListObject
    Dim local_tab_rows As ListRow
    Dim local_tab_rows_range As Range
    Dim msgBoxObj
    
    
    Set local_tab = ActiveWorkbook.Worksheets(local_list_index).ListObjects(local_table_index)
    
    If local_tab Is Nothing Then
        MsgBox "Tabular not found", vbCritical, "Init required"
    End If
    
    msgBoxObj = MsgBox("Reset Background-Color of Elements?", vbYesNo)
    
    If msgBoxObj = vbNo Then
        Exit Sub
    End If
    
    
    
    For Each local_tab_rows In local_tab.ListRows
        
        For Each local_tab_rows_range In local_tab_rows.Range
            local_tab_rows_range.Interior.Color = RGB(255, 255, 255)
        Next local_tab_rows_range
        
    Next local_tab_rows
    

End Sub

Private Sub Compare_Click()
   
    Dim ws As Object
    Dim local_tab As ListObject, remote_tab As ListObject
    ' Dim local_list_index As String, remote_list_index As String, local_table_index As String, remote_table_index As String
    
    If local_list_index = "" Then
        initWithDefault
    End If
    

    'From here nothing has to be changed
    
    'get the tables
    Set local_tab = ActiveWorkbook.Worksheets(local_list_index).ListObjects(local_table_index)
    Set remote_tab = ActiveWorkbook.Worksheets(remote_list_index).ListObjects(remote_table_index)
    'Remote table may be a subset of the local OR larger
    
    Dim remote_list_row As ListRow
    Dim local_list_row As ListRow
    
    Dim not_found_list_row As New Collection, found_list_row_diff As New Collection
    
    Dim not_found_row As ListRow
    Dim not_found_flag As Boolean, different_value_flag As Boolean
    Dim tmp_id As String
    
    'prevent working on the sheet itself
    Dim local_data As Variant
    Dim remote_data As Variant
    
    With local_tab.DataBodyRange
        local_data = .Value
    End With
    
    With remote_tab.DataBodyRange
        remote_data = .Value
    End With

    
    Dim i As Long
    Dim k As Long
    Dim tmp_range As Range
  
    For i = 1 To UBound(remote_data)
        not_found_flag = True
       ' remote_tab.ListRows(i).Range().Select
       ' Set tmp_local_value = remote_data(i)
        For k = LBound(local_data) To UBound(local_data)
           ' tmp_local_value = local_data(k)
            'what if empty value detected?
            
            If remote_data(i, 1) = local_data(k, 1) Then
                'local_tab.ListRows(k).Range().Select
                
                not_found_flag = False
        
                If StrComp(remote_data(i, 10), local_data(k, 10)) Then
                    local_tab.ListRows(k).Range(1, 10).Interior.Color = RGB(0, 204, 153)
                    'Set tmp_range = local_tab.ListRows(k).Range(1, 10)
                End If
                
                If StrComp(Trim(remote_data(i, 9)), Trim(local_data(k, 9))) <> 0 Then
                   local_tab.ListRows(k).Range(1, 9).Interior.Color = RGB(153, 204, 255)
                   'local_tab.ListRows(k).Range(1, 9).Select
                   'Set tmp_range = local_tab.ListRows(k).Range(1, 9)
                End If
                
                Exit For
            End If
        Next k
        
        If not_found_flag Then
            not_found_list_row.Add Item:=remote_tab.ListRows(i).Range
        End If
        
    Next i
    
     If not_found_list_row.Count > 0 Then
        Dim msgBoxResponse
        msgBoxResponse = MsgBox(CStr(not_found_list_row.Count) + " not found in list. Want to add them?", vbYesNo, "Add missing items?")
        
        If msgBoxResponse = vbYes Then
            update_missing_values local_tab, not_found_list_row
        End If
    
    End If

        
End Sub





Private Sub update_missing_values(ByVal table As ListObject, ByVal collect As Collection)

    Dim tmp_range As Range
    Dim tmp_list_row As ListRow

    Dim i As Integer, j As Integer
    
    For i = 1 To collect.Count
        'Set tmp_list_row = local_tab.ListRows.Add
        
        Set tmp_range = collect.Item(i)
        
        If Application.CountA(tmp_range.EntireRow) = 0 Then
            Next
        End If
        
        Set tmp_list_row = table.ListRows.Add
        
        tmp_range.Copy Destination:=tmp_list_row.Range
    
    Next
    
End Sub


Private Sub ShowInternalNames_Click()


    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim str As String
    Dim TableStrings As String
    
    'Loop through each sheet and table in the workbook
      For Each sht In ThisWorkbook.Worksheets
            
        For Each tbl In sht.ListObjects
          MsgBox "Sheet Name: " + sht.Name + " and Table: " + tbl.Name
          
            
        Next tbl
      Next sht


End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)



End Sub
