Attribute VB_Name = "Module2"

Public Sub addTouchToLog()
Attribute addTouchToLog.VB_ProcData.VB_Invoke_Func = "U\n14"
    
    Dim colListWORefNumber, colListTriggerDate, colListFormat1 As Integer
    Dim callsListSheetName, touchesLogSheetName As String
    Dim colLogDate, colLogTime, colLogReference As Integer
    Dim referenceCell As Range
    Dim referenceToLog As String
    
    'initializers
    colListWORefNumber = 8
    colListTriggerDate = 2
    colListFormat1 = 21
    callsListSheetName = "to-do"
    touchesLogSheetName = "outreachesLog"
    colLogDate = 1
    colLogTime = 2
    colLogReference = 3
    
    ' al hacer control + Shift + U,
    
    ' verifica hoja y cantidad de celdas seleccionadas
    If ActiveSheet.Name <> callsListSheetName Then
        MsgBox "No se encuentra en la hoja de llamadas", vbCritical + vbDefaultButton1 + vbOKOnly, "Error"
        End
    End If
    If Selection.Cells.Count > 1 Then
        MsgBox "Only one cell can be selected.", vbCritical + vbDefaultButton1 + vbOKOnly, "Error"
        End
    End If
    '/***********************need select of reference column? or only the row is ok **************************************/
    
    ' se haga registro de la orden seleccionada y se agrega a outreaches log
    referenceToLog = Cells(ActiveCell.Row, colListWORefNumber).Value
    Sheets(touchesLogSheetName).Select
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(endRow + 1, colLogDate).Value = Date
    Cells(endRow + 1, colLogTime).Value = Time
    Cells(endRow + 1, colLogReference).Value = referenceToLog
    ' /*****************************no colocar work orders duplicadas o  borrar duplicadas***************************/
    ' /*****************************   -> buscar del dia de hoy, y si ya existe no agregar registro al log********************************/
    Sheets(callsListSheetName).Select
    
    ' agregar fecha hoy a trigger column,  de List Sheet.
    Cells(ActiveCell.Row, colListTriggerDate).Value = Date
    With Cells(ActiveCell.Row, colListWORefNumber).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Cells(ActiveCell.Row, colListFormat1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub

Public Sub addManualTouch()
Attribute addManualTouch.VB_ProcData.VB_Invoke_Func = "T\n14"

    Dim colListWORefNumber, colListTriggerDate, colListFormat1 As Integer
    Dim callsListSheetName, touchesLogSheetName As String
    Dim colLogDate, colLogTime, colLogReference As Integer
    Dim referenceCell As Range
    Dim referenceToLog As String
    Dim prevActiveRow, prevActiveCol As Integer
    
    'initializers
    colListWORefNumber = 8
    colListTriggerDate = 2
    colListFormat1 = 21
    callsListSheetName = "to-do"
    touchesLogSheetName = "outreachesLog"
    colLogDate = 1
    colLogTime = 2
    colLogReference = 3
    
    '/********************BEFORE HERE IS THE COPY PASTE OF THE PREV SUB WITH EDIT FROM ACTIVECELL.ROW*********************/
    ' al hacer control + Shift + T,
    
    ' verifica hoja correcta
    If ActiveSheet.Name <> callsListSheetName Then
        MsgBox "No se encuentra en la hoja de llamadas", vbCritical + vbDefaultButton1 + vbOKOnly, "Error"
        End
    End If
    
    prevActiveRow = ActiveCell.Row
    prevActiveCol = ActiveCell.Column
    
    ' se haga registro de la orden tecleada manualmente y se agrega a outreaches log
    referenceToLog = InputBox("Enter the reference to log:")
    
    CLEARFILTERS
    
    ' Find the cell with the reference
    Set referenceCell = Sheets(callsListSheetName).Columns(colListWORefNumber).Find(What:=referenceToLog, LookIn:=xlValues, LookAt:=xlWhole)
    
    If referenceCell Is Nothing Then
        inputNotFoundBox = MsgBox("Reference not found!" & vbCrLf & "Deseas loguear la referencia manual " & referenceToLog & " ?", vbQuestion + vbDefaultButton1 + vbYesNo, "Log Manual Text?")
        If inputNotFoundBox = vbYes Then
            Sheets(touchesLogSheetName).Select
            endRow = Cells(Rows.Count, 1).End(xlUp).Row
            Cells(endRow + 1, colLogDate).Value = Date
            Cells(endRow + 1, colLogTime).Value = Time
            Cells(endRow + 1, colLogReference).Value = referenceToLog
            Sheets(callsListSheetName).Select
        ElseIf inputnotgfoundbox = vbNo Then
            'just end
        End If
        ThisWorkbook.Save
        End
    Else
        If referenceCell.Row > Cells(Rows.Count, 1).End(xlUp).Row Then
            MsgBox referenceCell.Row
            MsgBox Cells(Rows.Count, 1).End(xlUp).Row
            MsgBox "no information was entered.  input box is empty", vbCritical
            End
        End If
        
    End If
    
    '/********************AFTER HERE IS THE COPY PASTE OF THE PREV SUB WITH EDIT FROM ACTIVECELL.ROW*********************/
    Sheets(touchesLogSheetName).Select
    endRow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(endRow + 1, colLogDate).Value = Date
    Cells(endRow + 1, colLogTime).Value = Time
    Cells(endRow + 1, colLogReference).Value = referenceToLog
    ' /*****************************no colocar work orders duplicadas o  borrar duplicadas***************************/
    ' /*****************************   -> buscar del dia de hoy, y si ya existe no agregar registro al log********************************/
    Sheets(callsListSheetName).Select
    
    ' agregar fecha hoy a trigger column,  de List Sheet.
    Cells(referenceCell.Row, colListTriggerDate).Value = Date
    With Cells(referenceCell.Row, colListWORefNumber).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Cells(referenceCell.Row, colListFormat1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Cells(referenceCell.Row, colListFormat1).Select
    'ThisWorkbook.Save
    
    '/********************SELECT CELL WHERE WE WERE WORKING BEFORE*******************/
    '*********ONLY WHEN REFERENCE WAS FOUND ON THE SEARCH IT FILTERS TO ACTIVE CALL LIST*****************
    '**************ELSE: KEEPS TO-DO LIST WITHOUT FILTERS WHEN HAVING A MANUAL ADD OUTREACH***************************
    Cells(prevActiveRow, prevActiveCol).Activate
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=18, Criteria1:= _
        "<>"
    
    ThisWorkbook.Save
    
End Sub

Private Sub test012837401928374()

'    'endRow = Cells(Rows.Count, 1).End(xlUp).Row
'    'MsgBox endRow

'    'initializers
'    colListWORefNumber = 8
'    callsListSheetName = "to-do"
'    touchesLogSheetName = "outreachesLog"
'    colLogDate = 1
'    colLogTime = 2
'    colLogReference = 3
'    referenceToLog = "testvalue" 'Cells(ActiveCell.Row, colListWORefNumber).Value
'    Sheets(touchesLogSheetName).Select
'    endRow = Cells(Rows.Count, 1).End(xlUp).Row
'    Cells(endRow + 1, colLogDate).Value = Date
'    Cells(endRow + 1, colLogTime).Value = Time
'    Cells(endRow + 1, colLogReference).Value = referenceToLog
'    Sheets(callsListSheetName).Select
    
End Sub
