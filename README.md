Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng As Range
    Dim c As Range
    Dim sTemp As String
    Dim MyPos
    Dim DR1, DR2, DR3 As Worksheet
    'Dim arr() As String
    
    Set DR1 = ThisWorkbook.Worksheets("DR1")
    Set DR2 = ThisWorkbook.Worksheets("DR2")
    Set DR3 = ThisWorkbook.Worksheets("DR3")
    
    'Count hide row
    Set rngDR1 = DR1.Range("A1:A1000")
    Set rngDR2 = DR2.Range("A1:A1000")
    Set rngDR3 = DR3.Range("A1:A1000")
    'sTemp = ""

    For Each c In rngDR1
        If c.EntireRow.Hidden Then
            sTempDR1 = sTempDR1 & "L" & c.Row & "A" & vbCrLf
        End If
    Next c
    
    For Each c In rngDR2
        If c.EntireRow.Hidden Then
            sTempDR2 = sTempDR2 & "L" & c.Row & "A" & vbCrLf
        End If
    Next c
    
    For Each c In rngDR3
        If c.EntireRow.Hidden Then
            sTempDR3 = sTempDR3 & "L" & c.Row & "A" & vbCrLf
        End If
    Next c

    'If sTemp > "" Then
        'sTemp = "The following rows are hidden:" & vbCrLf & _
          'vbCrLf & sTemp
        'MsgBox sTemp
    'Else
        'MsgBox "There are no hidden rows"
    'End If
     
    
    If Not Intersect(Range("L55:L150"), Target) Is Nothing Then
        Application.EnableEvents = False
        Me.Unprotect Password:="secret"
        
        Application.CutCopyMode = False
        Application.OnKey "^c", ""
        Application.CellDragAndDrop = False
        
        Application.CutCopyMode = False
        Application.OnKey "^c", ""
        Application.CellDragAndDrop = False


         Select Case Range("L55").Value
            Case "○"
                'MsgBox sTempDR2
                For Each Stooge In Array("9", "10", "11", "12", "13")
                    'ReDim Preserve arr(UBound(arr) + 1)
                    arr(UBound(arr)) = Stooge
                'MsgBox Stooge
                        Transfrom = "L" + Stooge + "A"
                        MyPos = InStr(sTempDR2, Transfrom)
                        'MsgBox MyPos
                        'MsgBox Stooge
                        
                        'If duplicate
                        If MyPos <> 0 Then
                            'MsgBox "duplicate"
                            DR2.Range("A" + Stooge).EntireRow.Hidden = False
                            
                        'If = 0 (no duplicate)
                        Else
                            'MsgBox "no duplicate"
                            DR2.Range("A" + Stooge).EntireRow.Hidden = True
                            
                        End If
                Next Stooge
                MsgBox arr
            Case ""
                For Each Stooge In Array("9", "10", "11", "12", "13")
                    DR2.Range("A" + Stooge).EntireRow.Hidden = False
                Next Stooge
        End Select
        
        Select Case Range("L59").Value
            Case "○"
                'MsgBox sTempDR2
                For Each Stooge In Array("9", "10", "11", "12", "13")
                'MsgBox Stooge
                        Transfrom = "L" + Stooge + "A"
                        MyPos = InStr(sTempDR2, Transfrom)
                        'MsgBox MyPos
                        'MsgBox Stooge
                        
                        'If duplicate
                        If MyPos <> 0 Then
                            MsgBox "duplicate"
                            DR2.Range("A" + Stooge).EntireRow.Hidden = False
                            
                        'If = 0 (no duplicate)
                        Else
                            'MsgBox "no duplicate"
                            DR2.Range("A" + Stooge).EntireRow.Hidden = True
                            
                        End If
                Next Stooge
            Case ""
                For Each Stooge In Array("9", "10", "11", "12", "13")
                    DR2.Range("A" + Stooge).EntireRow.Hidden = False
                Next Stooge
        End Select
        
        Select Case Range("L60").Value
            Case "○"
                'MsgBox sTempDR2
                For Each Stooge In Array("9", "10", "16", "15", "17")
                'MsgBox Stooge
                        Transfrom = "L" + Stooge + "A"
                        'MsgBox sTempDR2
                        'MsgBox Transfrom
                        MyPos = InStr(sTempDR2, Transfrom)
                        'If hide
                        If MyPos <> 0 Then
                            DR2.Range("A" + Stooge).EntireRow.Hidden = False
                            'MsgBox MyPos
                        Else
                            DR2.Range("A" + Stooge).EntireRow.Hidden = True
                            'MsgBox Stooge
                        End If
                Next Stooge
            Case ""
                For Each Stooge In Array("9", "10", "16", "15", "17")
                    DR2.Range("A" + Stooge).EntireRow.Hidden = False
                Next Stooge
         End Select
        
        
        Me.Protect Password:="secret"
        Application.EnableEvents = True
    End If
End Sub





Private Sub Workbook_Activate()
Application.CutCopyMode = False
Application.OnKey "^c", ""
Application.CellDragAndDrop = False
End Sub

Private Sub Workbook_Deactivate()
Application.CellDragAndDrop = True
Application.OnKey "^c"
Application.CutCopyMode = False
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
Application.CutCopyMode = False
Application.OnKey "^c", ""
Application.CellDragAndDrop = False
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
Application.CellDragAndDrop = True
Application.OnKey "^c"
Application.CutCopyMode = False
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
Application.CutCopyMode = False
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
Application.OnKey "^c", ""
Application.CellDragAndDrop = False
Application.CutCopyMode = False
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
Application.CutCopyMode = False
End Sub









Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rng As Range
    Dim c As Range
    Dim sTemp As String
    Dim MyPos
    Dim Stooge  As Variant
    
    Set rng = Range("A1:A1000")
    sTemp = ""

    For Each c In rng
        If c.EntireRow.Hidden Then
            sTemp = sTemp & "L" & c.Row & "A" & vbCrLf
        End If
    Next c

    If sTemp > "" Then
        sTemp = "The following rows are hidden:" & vbCrLf & _
          vbCrLf & sTemp
          
            arr = Split(sTemp, ",")
            hidelist = ""
            For i = 0 To UBound(arr)
                hidelist = hidelist + vbNewLine + vbNewLine + arr(i)
            Next i

        'MsgBox hidelist
    Else
        MsgBox "There are no hidden rows"
    End If
     
    
    If Not Intersect(Range("L5:L20"), Target) Is Nothing Then
        Application.EnableEvents = False
        'Me.Unprotect Password:="secret"
        'Range("A6:A61").EntireRow.Hidden = True
        Select Case Range("L5").Value
            Case "○"
                For Each Stooge In Array("L" + "18" + "A", "L" + "1" + "A", "L" + "21" + "A")
                   MyPos = InStr(sTemp, Stooge)
                   'MsgBox hidelist
                        If MyPos <> 0 Then
                            MsgBox "ok"
                        Else
                            MsgBox "no"
                        End If
                Next Stooge
            Case ""
                For Each Stooge In Array("18", "2", "21")
                    'Range("L" + Stooge).EntireRow.Hidden = False
                Next Stooge
        End Select
        
        
        'Me.Protect Password:="secret"
        Application.EnableEvents = True
    End If
End Sub
























