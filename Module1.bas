Attribute VB_Name = "Module1"
'variables globales
Dim globalBooleanRequestDateDifferentThanToday As Boolean
Dim globalStringDateDifferentThanToday As String

Sub ToExecuteWhenRequestDateNotToday()

    globalBooleanRequestDateDifferentThanToday = True
    globalStringDateDifferentThanToday = InputBox("fecha a poner como request date  --  MM/DD/YYYY")

    inputStr = InputBox("1: varias ordenes moradas" + vbCrLf + "2:una sola orden morada" + vbCrLf + "3:marcar trabajos para hoy" + vbCrLf + vbCrLf + vbCrLf)
    If inputStr = 1 Then
        Call z1_updateSeveralLastRequestDatesSeguiditasUnaTrasOtra
    ElseIf inputStr = 2 Then
        Call a1_updateLastRequestDate
    ElseIf inputStr = 3 Then
        Call markTodayWorks
    Else
        MsgBox "ERROR: INPUT INVALIDA"
        Stop
    End If
    
    globalBooleanRequestDateDifferentThanToday = False
    
End Sub

Sub z1_updateSeveralLastRequestDatesSeguiditasUnaTrasOtra()

    Dim inputStr As String
    Dim inputInteger As Integer
    
   ' MsgBox globalBooleanRequestDateDifferentThanToday
    
    inputStr = InputBox("cuantas ordenes seguiditas estan repetidas en morado ? ")
    inputInteger = CInt(inputStr)
    
    For i1 = 1 To inputInteger
        a1_updateLastRequestDate
    Next i1

End Sub
Sub a1_updateLastRequestDate()

    Dim currRow, currCol, resultRow, resultCol  As Integer
    Dim colLastRequestDate, colWORefNumber, colTriggerDate As Integer
    Dim WOToSearch As String
    Dim questionAnswer As Integer
    Dim colorIndexOfCurrSearch As Integer
  
     ' MsgBox globalBooleanRequestDateDifferentThanToday

    CLEARFILTERS
        
    colLastRequestDate = 1
    colWORefNumber = 8
    colTriggerDate = 2
    
    
    'GET CURRENT VALUE TO SEARCH
    currRow = ActiveCell.Row
    currCol = ActiveCell.Column
    WOToSearch = Cells(currRow, currCol).Value
    'questionAnswer = MsgBox("urgente?", vbDefaultButton2 + vbYesNo + vbQuestion, "rojo?")
    colorIndexOfCurrSearch = GetABIReferenceColorIndex(currRow, currCol)
    If colorIndexOfCurrSearch = 3 Then
        questionAnswer = 6
    ElseIf colorIndexOfCurrSearch = 46 Then
        questionAnswer = 46
    ElseIf colorIndexOfCurrSearch = 1 Or colorIndexOfCurrSearch = -4105 Then
        questionAnswer = 7
    Else
        Stop
        'not considered color
    End If
    
    'GET CURRENT RESULT
    Columns(colWORefNumber).Select
    Cells(1, colWORefNumber).Activate
    Selection.Find(What:=WOToSearch, After:=ActiveCell, LookIn:=xlFormulas2 _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    resultRow = ActiveCell.Row
    resultCol = ActiveCell.Column
    If resultRow = currRow Then
        i = MsgBox("ERROR, SIN DUPLICADO ENCONTRADO APARENTE", vbCritical + vbOKOnly, "ERROR DE BUSQUEDA")
        Exit Sub
    End If
    '6 si ----  7 no
    If questionAnswer = 6 Then
        Rows(resultRow).Select
        With Selection.Font
            .Color = -16776961
            .TintAndShade = 0
        End With
    ElseIf questionAnswer = 7 Then
        Rows(resultRow).Select
        With Selection.Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
    ElseIf questionAnswer = 46 Then
        Rows(resultRow).Select
        With Selection.Font
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0
        End With
    Else
        Stop
        'option not considered
    End If
    Cells(resultRow, resultCol).Select
    
    
    If resultCol = colWORefNumber And resultRow <> currRow Then
                
                
                
                If globalBooleanRequestDateDifferentThanToday = False Then
                    Cells(resultRow, colLastRequestDate).FormulaR1C1 = CStr(Date)
                ElseIf globalBooleanRequestDateDifferentThanToday = True Then
                    Cells(resultRow, colLastRequestDate).FormulaR1C1 = globalStringDateDifferentThanToday
                Else
                    MsgBox "SYSTEM ERROR"
                    Stop
                End If
                
                
    
        Rows(currRow).Select
        Selection.Delete Shift:=xlUp
        currRow = currRow
        Call FormatConditionalFormatingDuplicateAndTodayTrigger(colWORefNumber, colTriggerDate)
        Cells(currRow, currCol).Select
    Else
        MsgBox "ERROR"
        Stop
    End If
    'Stop
        
'    If Cells(currRow.currCol).Value = "" Then
'        Exit Sub
'    End If
    
    
End Sub

'Sub TESTERÑASDLKJÑASDF()
'    'NEW AMARILLO  = 53
'    'NEW ROJO = 9
'        colorIndexNum = GetABIReferenceColorIndex(1341, 8)
'        Stop
'End Sub

Sub markTodayWorks()
    
    Dim colLastRequestDate, startRow, endRow, colLastTrigger, daysToResaltarFromThePast, i1 As Integer
    Dim tempStr As String
    Dim inputSelectionBox, colMarked, colABIReferenceNumber As Integer
    Dim flgAdditionalPrevDays As Boolean
    Dim cntResaltadas As Integer
    Dim colorIndexNum, colIDInternal As Integer
    Dim inputSelectionBox2 As String
    Dim additionalDaysToMark As Integer
    
   '     MsgBox globalBooleanRequestDateDifferentThanToday

    
    'prep worksheet
    CLEARFILTERS
    quitarMarkToday
    
    'initializers
    colLastRequestDate = 1
    colLastTrigger = 2
    startRow = 2
    flgAdditionalPrevDays = False
    colMarked = 16
    cntResaltadas = 0
    colABIReferenceNumber = 8
    colIDInternal = 7
    additionalDaysToMark = 0
    
    'days to actually go back and rellenar as pending
    daysToResaltarFromThePast = 1
    'Add missing IDs for new added reference numbers
    Call addTheMissingIDsForTheNewReferencesReceived(colIDInternal, colABIReferenceNumber, colLastRequestDate)

ReDoResaltarDeFilas:
    cntResaltadas = cntResaltadas + 1
    'get endRow
    Cells(startRow, colIDInternal).Select
    Selection.End(xlDown).Select
    endRow = Selection.Row
    
    'mark all rows inside rules
    For i1 = startRow To endRow
        
        'Cells(i1, colLastRequestDate).Select
        If cntResaltadas = 1 Then
            Cells(i1, colMarked) = ""
        End If
        'due today
        If Cells(i1, colLastRequestDate) = Date Then
            Rows(i1).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            Cells(i1, colMarked) = True
            'MARK CAFES SIEMPRE TRUE HASTA REALIZADOS (PUESTOS EN NEGRO O ROJO)
            colorIndexNum = GetABIReferenceColorIndex(i1, colABIReferenceNumber)
            If colorIndexNum = 53 Then
                Cells(i1, colMarked) = True
            End If
        'due yesterday undone
        ElseIf Cells(i1, colLastRequestDate) = (Date - daysToResaltarFromThePast) Then
            'Stop
            If flgAdditionalPrevDays Then
                If Cells(i1, colLastTrigger) <> (Date - daysToResaltarFromThePast) Then
                    'Stop
                    Rows(i1).Select
                    If flgAdditionalPrevDays Then
                        If Cells(i1, colLastTrigger).Value > Cells(i1, colLastRequestDate).Value Then
                            'skip formating
                            GoTo JustSkipFormatingForPrevDays
                        Else
                            'continue formating
                        End If
                    End If
                    With Selection.Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorAccent1
                        .TintAndShade = 0.599993896298105
                        .PatternTintAndShade = 0
                    End With
                    Cells(i1, colMarked) = True
JustSkipFormatingForPrevDays:
                    'Stop
                End If
            End If
        End If
    Next i1
    
    'Set same color of text for row as the ABI Reference Font Color
    For i1 = startRow To endRow
            '    'ROW#        COLOR            COLORINDEX
            '    '22 = NEGRO (AUTOMATIC)         -4105  ' this is default black
            '    '36 = NEGRO                       1    ' this is the black we want on all
            '    '37 = ROJO                        3
            '    '340 = CAFÉ                       53
        'Format Row Text Color
        colorIndexNum = GetABIReferenceColorIndex(i1, colABIReferenceNumber)
        If colorIndexNum = -4105 Then
            colorIndexNum = 1
        End If
        If colorIndexNum = -4105 Or colorIndexNum = 1 Or colorIndexNum = 3 Or colorIndexNum = 53 Or colorIndexNum = 46 Then
            ' do nothing
        Else
            Stop
        End If
        Call SetFullRowColorIndex(i1, colorIndexNum)
        
        'test
        If GetABIReferenceColorIndex(i1, 1) = GetABIReferenceColorIndex(i1, colABIReferenceNumber) Then
            ' do nothing
        Else
            Stop
        End If
        
    Next i1
    
    
    'ADDITIONAL DAYS TO FORMAT
    If additionalDaysToMark = 0 Then
          '6 yes ----  7 no
        inputSelectionBox = MsgBox("Necesitas dias adicionales para marcar pendientes de dejar status?", vbQuestion + vbDefaultButton2 + vbYesNo, "Marcar otros aparte de ayer?")
        If inputSelectionBox = 6 Then
            MsgBox "indique cuantos dias va a solicitar"
            inputSelectionBox2 = InputBox("Cuantos dias adicionales necesitas en bucle?", "cuantos dias aditionales?")
            additionalDaysToMark = CInt(inputSelectionBox2)
        End If
    End If
    If inputSelectionBox = 6 Then
        tempStr = InputBox("hace cuantos dias fue el dia adicional a marcar?", "pending to status to client", "lunes poner 3 y 4 // martes a viernes 2 dias")
        daysToResaltarFromThePast = CInt(tempStr)
        flgAdditionalPrevDays = True
        additionalDaysToMark = additionalDaysToMark - 1
        'Stop ' checar que el contador regresivo funcione para los 5 dias que estas pidiendo adicionales a hoy
        GoTo ReDoResaltarDeFilas
    ElseIf inputSelectionBox = 7 Then
        ' just continue and do nothing
    Else
        MsgBox "ERROR FATAL"
        Stop
    End If
    
    'format row height
    tempStr = "1:" & CStr(endRow)
    Rows(tempStr).Select
    Selection.RowHeight = 15
    'filter only marked as highlighted
    Range("A1").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=16, Criteria1:= _
        "<>"
    ActiveWindow.ScrollRow = 2
    MsgBox "se muestran los resaltados (filtro)"
    
    
End Sub

Private Sub quitarMarkToday()

    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Private Sub CLEARFILTERS()
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub

Sub FormatConditionalFormatingDuplicateAndTodayTrigger(ByVal colWOnumber As Integer, ByVal colTriggerDate As Integer)

    Application.CutCopyMode = False
    Cells.FormatConditions.Delete
    Columns(colWOnumber).Select
    Cells(2, colWOnumber).Activate
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Columns(colTriggerDate).Select
    Cells(2, colTriggerDate).Activate
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=TODAY()"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub

'    SetFullRowColorIndex (rowNumberToChangeColor)
            '    'ROW#        COLOR            COLORINDEX
            '    '22 = NEGRO (AUTOMATIC)         -4105  ' this is default black
            '    '36 = NEGRO                       1    ' this is the black we want on all
            '    '37 = ROJO                        3
            '    '340 = CAFÉ                       53
            '   NEW AMARILLO  = 53
            '   NEW ROJO = 9
            
Function GetABIReferenceColorIndex(ByVal rowNumber As Integer, ByVal columnNumber As Integer) As Integer
    
    Dim columnLetter, rangeFormat As String
    
    columnLetter = NumeroALetra(columnNumber)
    If columnLetter = "Error" Then
        Stop
    End If
    
    rangeFormat = columnLetter & CStr(rowNumber)
    GetABIReferenceColorIndex = Range(rangeFormat).Font.ColorIndex
    
End Function

Sub SetFullRowColorIndex(ByVal rowNumber As Integer, ByVal colorIndexNumber As Integer)

    Rows(rowNumber).Font.ColorIndex = colorIndexNumber

End Sub




Function NumeroALetra(ByVal numero As Integer) As String
    Select Case numero
    Case 1
    NumeroALetra = "A"
    Case 2
    NumeroALetra = "B"
    Case 3
    NumeroALetra = "C"
    Case 4
    NumeroALetra = "D"
    Case 5
    NumeroALetra = "E"
    Case 6
    NumeroALetra = "F"
    Case 7
    NumeroALetra = "G"
    Case 8
    NumeroALetra = "H"
    Case 9
    NumeroALetra = "I"
    Case 10
    NumeroALetra = "J"
    Case 11
    NumeroALetra = "K"
    Case 12
    NumeroALetra = "L"
    Case 13
    NumeroALetra = "M"
    Case 14
    NumeroALetra = "N"
    Case 15
    NumeroALetra = "O"
    Case 16
    NumeroALetra = "P"
    Case 17
    NumeroALetra = "Q"
    Case 18
    NumeroALetra = "R"
    Case 19
    NumeroALetra = "S"
    Case 20
    NumeroALetra = "T"
    Case 21
    NumeroALetra = "U"
    Case 22
    NumeroALetra = "V"
    Case 23
    NumeroALetra = "W"
    Case 24
    NumeroALetra = "X"
    Case 25
    NumeroALetra = "Y"
    Case 26
    NumeroALetra = "Z"
    Case Else
    NumeroALetra = "Error" ' Valor de error para números no válidos
    End Select
End Function

Sub addTheMissingIDsForTheNewReferencesReceived(ByVal colIDInternal As Integer, ByVal colABIRefNumber As Integer, ByVal colLastRequestDate As Integer)
'
' Macro1 Macro
'

'

    Dim lastIDUsed, currRow, currCol As Integer
    'CLEAR FILTERS
    CLEARFILTERS
    'ASCENDING ID INTERNAL NUMBER
    Range("A2").Select
    ActiveWorkbook.Worksheets("to-do").ListObjects("Table1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("to-do").ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[[#All],[ID]]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("to-do").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("G2").Select
    Selection.End(xlDown).Select
    lastIDUsed = Selection.Value
    currRow = Selection.Row
    currCol = colIDInternal
    Do While Cells(currRow + 1, colABIRefNumber) <> ""
        lastIDUsed = lastIDUsed + 1
        currRow = currRow + 1
        Cells(currRow, colIDInternal).Select
        Selection.Value = lastIDUsed
        If Cells(currRow, colLastRequestDate).Value = "" Then
        
                If globalBooleanRequestDateDifferentThanToday = False Then
                    Cells(currRow, colLastRequestDate).Value = CStr(Date)
                ElseIf globalBooleanRequestDateDifferentThanToday = True Then
                   Cells(currRow, colLastRequestDate).Value = globalStringDateDifferentThanToday
                Else
                    MsgBox "SYSTEM ERROR"
                    Stop
                End If
           
        End If
    Loop
    'ASCENDING FACILITY NAME COLUMN
    Range("A2").Select
    ActiveWorkbook.Worksheets("to-do").ListObjects("Table1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("to-do").ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[[#All],[FACILITY NAME]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("to-do").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'END WITH FIRST ROW SELECT
    Range("A2").Select
    ActiveWindow.ScrollRow = 2
End Sub

