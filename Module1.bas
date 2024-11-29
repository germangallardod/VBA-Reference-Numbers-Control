Attribute VB_Name = "Module1"
'variables globales
Dim globalBooleanRequestDateDifferentThanToday As Boolean
Dim globalStringDateDifferentThanToday As String
Dim globalBooleana1UpdateLastRequestDateAssignedRowAndColumn As Boolean
Dim globalRowAssignedToBeUsedInA1Procedure, globalColAssignedToBeUsedInA1Procedure As Integer

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

Private Sub z1_updateSeveralLastRequestDatesSeguiditasUnaTrasOtra()

    Dim inputStr As String
    Dim inputInteger As Integer
    
   ' MsgBox globalBooleanRequestDateDifferentThanToday
    
    inputStr = InputBox("cuantas ordenes seguiditas estan repetidas en morado ? ")
    inputInteger = CInt(inputStr)
    
    For i1 = 1 To inputInteger
        a1_updateLastRequestDate
    Next i1

End Sub
Private Sub a1_updateLastRequestDate()

    Dim currRow, currCol, resultRow, resultCol  As Integer
    Dim colLastRequestDate, colWORefNumber, colTriggerDate As Integer
    Dim WOToSearch As String
    Dim questionAnswer As Integer ' THIS IS ACTUALLY THE TYPE OF REQUEST URGENT(RED)6 NORMAL (BLACK)7 OTHER (BROWN)46
    Dim colorIndexOfCurrSearch As Integer
    Dim lastStatusDate, lastStatusText, locationAging, typeOfRecords, facilityUSState As Variant
    Dim colLastStatusDate, colLastStatusText, colLocationAging, colTypeOfRecords, colFacilityUSState As Integer
      
     ' MsgBox globalBooleanRequestDateDifferentThanToday

    CLEARFILTERS
        
    colLastRequestDate = 1
    colWORefNumber = 8
    colTriggerDate = 2
    colLocationAging = 15
    colLastStatusText = 14
    colLastStatusDate = 13
    colTypeOfRecords = 16
    colFacilityUSState = 11
    
    
    'GET CURRENT VALUE TO SEARCH
    If globalBooleana1UpdateLastRequestDateAssignedRowAndColumn Then
        currRow = globalRowAssignedToBeUsedInA1Procedure
        currCol = globalColAssignedToBeUsedInA1Procedure
    Else
        currRow = ActiveCell.Row
        currCol = ActiveCell.Column
    End If
    WOToSearch = Cells(currRow, colWORefNumber).Value
    lastStatusDate = Cells(currRow, colLastStatusDate).Value
    lastStatusText = Cells(currRow, colLastStatusText).Value
    locationAging = Cells(currRow, colLocationAging).Value
    typeOfRecords = Cells(currRow, colTypeOfRecords).Value
    facilityUSState = Cells(currRow, colFacilityUSState).Value
    'questionAnswer = MsgBox("urgente?", vbDefaultButton2 + vbYesNo + vbQuestion, "rojo?")
    colorIndexOfCurrSearch = GetABIReferenceColorIndex(currRow, currCol)
    If colorIndexOfCurrSearch = 3 Or colorIndexOfCurrSearch = 9 Then  ' rojo
        questionAnswer = 6
    ElseIf colorIndexOfCurrSearch = 46 Then ' cafe
        questionAnswer = 46
    ElseIf colorIndexOfCurrSearch = 1 Or colorIndexOfCurrSearch = -4105 Then ' black
        questionAnswer = 7
    Else
        Stop
        'not considered color
    End If
    
    'GET CURRENT RESULT
    'Columns(colWORefNumber).Select
    Cells(1, colWORefNumber).Activate
    Columns(colWORefNumber).Find(What:=WOToSearch, After:=ActiveCell, LookIn:=xlFormulas2 _
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
        'Rows(resultRow).Select
        With Rows(resultRow).Font
            .Color = -16776961
            .TintAndShade = 0
        End With
    ElseIf questionAnswer = 7 Then
        'Rows(resultRow).Select
        With Rows(resultRow).Font
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
        End With
    ElseIf questionAnswer = 46 Then
        'Rows(resultRow).Select
        With Rows(resultRow).Font
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0
        End With
    Else
        Stop
        'option not considered
    End If
    'Cells(resultRow, resultCol).Select
    
    
    If resultCol = colWORefNumber And resultRow <> currRow Then
                
                
        'CHANGE REQUEST DATE
        If globalBooleanRequestDateDifferentThanToday = False Then
            Cells(resultRow, colLastRequestDate).FormulaR1C1 = CStr(Date)
        ElseIf globalBooleanRequestDateDifferentThanToday = True Then
            Cells(resultRow, colLastRequestDate).FormulaR1C1 = globalStringDateDifferentThanToday
        Else
            MsgBox "SYSTEM ERROR"
            Stop
        End If
        
        'CHANGE STATUS DATE, TEXT AND LOCATION AGING
        Cells(resultRow, colLastStatusDate).Value = lastStatusDate
        Cells(resultRow, colLastStatusText).Value = lastStatusText
        Cells(resultRow, colLocationAging).Value = locationAging
        Cells(resultRow, colLocationAging).Value = locationAging
        Cells(resultRow, colTypeOfRecords).Value = typeOfRecords
        Cells(resultRow, colFacilityUSState).Value = facilityUSState
        
        'DELETE ROW OF NEW CALL REQUEST (DUPLICATE)
'        Rows(currRow).Select
'        Cells(currRow + 2, 1).Select
        Rows(currRow).Delete Shift:=xlUp
        currRow = currRow
        Call FormatConditionalFormatingDuplicateAndTodayTrigger(colWORefNumber, colTriggerDate)
        'Cells(currRow, currCol).Select
    Else
        MsgBox "ERROR"
        Stop
    End If
    'Stop
        
'    If Cells(currRow.currCol).Value = "" Then
'        Exit Sub
'    End If
    
    
End Sub


Sub markTodayWorks()
    
    Dim colLastRequestDate, startRow, endRow, colLastTrigger, daysToResaltarFromThePast, i1 As Integer
    Dim tempStr As String
    Dim inputSelectionBox, colMarked, colABIReferenceNumber As Integer
    Dim flgAdditionalPrevDays As Boolean
    Dim cntResaltadas As Integer
    Dim colorIndexNum, colIDInternal As Integer
    Dim inputSelectionBox2 As String
    Dim additionalDaysToMark As Integer
    Dim colLocationAging, colLastStatusText, colLastStatusDate As Integer
    Dim colCalculatedDaysFromLastStatus As Integer
    Dim daysFromLastVisibleStatusVBA As Integer
    
   '     MsgBox globalBooleanRequestDateDifferentThanToday
    
    'initializers
    colLastRequestDate = 1
    colLastTrigger = 2
    startRow = 2
    flgAdditionalPrevDays = False
    colMarked = 18
    cntResaltadas = 0
    colABIReferenceNumber = 8
    colIDInternal = 7
    additionalDaysToMark = 0
    colLocationAging = 15
    colLastStatusText = 14
    colLastStatusDate = 13
    colCalculatedDaysFromLastStatus = 17
    
    'prep worksheet
    CLEARFILTERS
    ChangeToRedUrgents (colABIReferenceNumber)
    CLEARFILTERS
    quitarMarkToday 'quitar rellenos de las celdas
    
    'stop automatic formula calculation
    Application.Calculation = xlManual
    
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
'            Rows(i1).Select
'            With Selection.Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorAccent2
'                .TintAndShade = 0.599993896298105
'                .PatternTintAndShade = 0
'            End With
'            With Selection.Font
'                .Name = "Calibri"
'                .Strikethrough = False
'                .Superscript = False
'                .Subscript = False
'                .OutlineFont = False
'                .Shadow = False
'                .Underline = xlUnderlineStyleNone
'                .TintAndShade = 0
'                .ThemeFont = xlThemeFontMinor
'            End With
'            With Selection.Font
'                .Name = "Calibri"
'                .Size = 11
'                .Strikethrough = False
'                .Superscript = False
'                .Subscript = False
'                .OutlineFont = False
'                .Shadow = False
'                .Underline = xlUnderlineStyleNone
'                .TintAndShade = 0
'                .ThemeFont = xlThemeFontMinor
'            End With
            With Rows(i1).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            With Rows(i1).Font
                .Name = "Calibri"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            Cells(i1, colMarked) = True
            
            'MARK A LITTLE DARKER IF BLACK 'and' DAYS TO OVERDUE IS CLOSE TO TOMORROW (WORKDAY)
            'ONLY FOR TASKS ASSIGNED FOR TODAY
            'IF NEEDED FOR PREV TO WORK DAYS, GO TO NEXT ELSE IF AND ADD THIS LINES TOO (ADAPTED)
            'Stop
                If GetABIReferenceColorIndex(i1, colABIReferenceNumber) = 1 Or GetABIReferenceColorIndex(i1, colABIReferenceNumber) = -4105 Then
                    
                    daysFromLastVisibleStatusVBA = Fix(Date - Cells(i1, colLastStatusDate).Value)
                    'MsgBox daysFromLastVisibleStatusVBA
                    
                    'Stop
                    'Calculate
                    '* reembplazar esta variable contra la revision manual de la celda para evitar calcular manualmente toda la hoja
                    
                    
                    If daysFromLastVisibleStatusVBA >= 20 Then
                        'Stop
                        'Rows(i1).Select
                        With Rows(i1).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent2
                            .TintAndShade = 0.399975585192419
                            .PatternTintAndShade = 0
                        End With
                    End If
                End If
            'Stop
                
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
            '    '340 = CAF?                       53
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
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=colMarked, Criteria1:= _
        "<>"
    ActiveWindow.ScrollRow = 2
    MsgBox "se muestran los resaltados (filtro)"
    
    'REACTIVA CALCULOS AUTOMATICOS
    Calculate
    Application.Calculation = xlAutomatic
    
    
End Sub

Private Sub quitarMarkToday()

    Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub

Public Sub CLEARFILTERS()
    On Error Resume Next
    
    Dim selectedRow, selectedCol As Integer
    
    selectedRow = Selection.Row
    selectedCol = Selection.Column
    Range("h1").Select
    ActiveSheet.ShowAllData
    Cells(selectedRow, selectedCol).Select
End Sub

Sub FormatConditionalFormatingDuplicateAndTodayTrigger(ByVal colWOnumber As Integer, ByVal colTriggerDate As Integer)

    Application.CutCopyMode = False
    Cells.FormatConditions.Delete
    'Columns(colWOnumber).Select
    'Cells(2, colWOnumber).Activate
    Columns(colWOnumber).FormatConditions.AddUniqueValues
    Columns(colWOnumber).FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Columns(colWOnumber).FormatConditions(1).DupeUnique = xlDuplicate
    With Columns(colWOnumber).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10498160
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
'    Columns(colTriggerDate).Select
'    Cells(2, colTriggerDate).Activate
    Columns(colTriggerDate).FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=TODAY()"
    Columns(colTriggerDate).FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Columns(colTriggerDate).FormatConditions(1).Interior
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
            '    '340 = CAF?                       53
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
    NumeroALetra = "Error" ' Valor de error para n?meros no v?lidos
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
'    Range("G2").Select
'    Selection.End(xlDown).Select
'    lastIDUsed = Selection.Value
    With Range("G2").End(xlDown)
        lastIDUsed = .Value
        currRow = .Row
    End With
    currCol = colIDInternal
    Do While Cells(currRow + 1, colABIRefNumber) <> ""
        lastIDUsed = lastIDUsed + 1
        currRow = currRow + 1
'        Cells(currRow, colIDInternal).Select
'        Selection.Value = lastIDUsed
        Cells(currRow, colIDInternal).Value = lastIDUsed
        
        
'        tempvalwo = Cells(currRow, colABIRefNumber).Value
'        MsgBox tempvalwo
'        temprowmatchresult = Application.WorksheetFunction.Match(Cells(currRow, colABIRefNumber).Value, Range(NumeroALetra(colABIRefNumber) & CStr(1) & ":" & NumeroALetra(colABIRefNumber) & CStr(currRow)), 0)
'        MsgBox temprowmatchresult
        
        
        
        
        'Stop
        
        If (Application.WorksheetFunction.Match(Cells(currRow, colABIRefNumber).Value, Range(NumeroALetra(colABIRefNumber) & CStr(1) & ":" & NumeroALetra(colABIRefNumber) & CStr(currRow)), 0)) = currRow Then
            'IF THE ORDER IS A NEW ORDER
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
        Else
            'IF THE ORDER IS A DUPLICATE ORDER
            'Cells(currRow, colABIRefNumber).Select
            globalBooleana1UpdateLastRequestDateAssignedRowAndColumn = True
            globalRowAssignedToBeUsedInA1Procedure = currRow
            globalColAssignedToBeUsedInA1Procedure = colABIRefNumber
            Call a1_updateLastRequestDate
            globalBooleana1UpdateLastRequestDateAssignedRowAndColumn = False
            lastIDUsed = lastIDUsed - 1
            currRow = currRow - 1
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

'
'Private Function esOrdenDuplicada(ByVal currRow As Integer, ByVal colWOlocation As String) As Boolean
'
'    'if((
'        MATCH(N44,$C$7:$C$847,0)+XXXXXXXX
'        =ROW(N44),1,0)
'
'End Function


Sub ChangeToRedUrgents(ByVal colABIReference As Integer)
'
' Macro1 Macro
'

'
    Dim inputStr, rangeStr As String

    'CLEAR FILTERS
    CLEARFILTERS
    'filter new reds to give format
    Range("h2").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=9, Criteria1:=RGB _
        (255, 199, 206), Operator:=xlFilterCellColor
ReAskQuestion:
    inputStr = InputBox("cual es la primer row de rojos? - elija ninguno si todos son negros", "first filtered row")
    If inputStr = "ninguno" Then
        Exit Sub
    ElseIf inputStr = "" Then
        MsgBox "numero de celda inicial de rojos no puede ser vacio"
        GoTo ReAskQuestion
    ElseIf Not IsNumeric(inputStr) Then
        MsgBox "numero de fila requiere ser numerico"
        GoTo ReAskQuestion
        If InStr(inputStr, ".") > 0 Then
            MsgBox "numero de fila no puede tener decimales"
            GoTo ReAskQuestion
            ' Check if the numeric value is an integer
            If CLng(inputStr) <> Val(inputStr) Then
                MsgBox "this is not a valid value"
                GoTo ReAskQuestion
            End If
        End If
'    Else
'        MsgBox "this is not a valid value"
'        GoTo ReAskQuestion
    End If
    rangeStr = NumeroALetra(colABIReference) & inputStr
    
    
    Range(rangeStr).Select
    i1 = MsgBox("celda seleccionada es la correcta?", vbQuestion + vbYesNo + vbDefaultButton1, "celda seleccionada correcta?")
    If i1 = 6 Then 'yes
        ' do nothing
    ElseIf i1 = 7 Then 'no
        'ask question again
        GoTo ReAskQuestion
    Else
        'error
        Stop
    End If
    
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Range(rangeStr).Select
    
End Sub

