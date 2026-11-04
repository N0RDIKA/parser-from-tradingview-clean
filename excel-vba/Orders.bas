Attribute VB_Name = "Orders"
'===========================================================
' Формирование файла транзакций .tri для торговой системы.
' Одна строка на инструмент: рыночная заявка NEW_ORDER.
'===========================================================
Option Explicit

Private lastTransId As Long

' TRANS_ID: десятые доли секунды с начала дня, но всегда больше предыдущего.
' Раньше ID собирался из строки времени плюс номер строки и мог совпадать
' или выходить за пределы целого числа.
Private Function NextTransId() As Long
    Dim candidate As Long
    candidate = CLng(Timer * 10)
    If candidate <= lastTransId Then candidate = lastTransId + 1
    lastTransId = candidate
    NextTransId = candidate
End Function

Private Function Opposite(ByVal side As String) As String
    If side = "Buy" Then Opposite = "Sell" Else Opposite = "Buy"
End Function

' ws    - лист вывода с настройками в столбце R
' side  - "Buy" или "Sell"
' x1,x2 - диапазон строк с заявками
Public Function WriteOrders(ByVal ws As Worksheet, ByVal side As String, _
                            ByVal x1 As Long, ByVal x2 As Long) As String
    Dim i As Long, sym As String, vol As Double
    Dim rowSide As String, classcode As String, account As String
    Dim clientCode As String, triPath As String, orders As String
    Dim isHedge As Boolean, f As Integer

    On Error GoTo Fail

    triPath = Trim$(CStr(ws.Cells(CFG_ROW_TRI_PATH, CFG_COL).Value))
    If triPath = "" Then Err.Raise vbObjectError + 1, , "Не указан путь к файлу заявок в R" & CFG_ROW_TRI_PATH & " на листе " & ws.Name

    clientCode = Trim$(CStr(ws.Cells(CFG_ROW_CLIENT_CODE, CFG_COL).Value))

    For i = x1 To x2
        sym = Trim$(CStr(ws.Cells(i, 2).Value))
        vol = Val(CStr(ws.Cells(i, 10).Value))

        If sym <> "" And vol <> 0 Then
            isHedge = (UCase$(sym) = HEDGE_SECCODE)

            ' Хедж фьючерсом на индекс открывается в обратную сторону
            If isHedge Then
                rowSide = Opposite(side)
                classcode = HEDGE_CLASSCODE
                account = HEDGE_ACCOUNT
            Else
                rowSide = side
                classcode = Trim$(CStr(ws.Cells(CFG_ROW_CLASSCODE, CFG_COL).Value))
                account = Trim$(CStr(ws.Cells(CFG_ROW_ACCOUNT, CFG_COL).Value))
            End If

            orders = orders & "ACTION=NEW_ORDER;TRANS_ID=" & NextTransId() & _
                     ";CLASSCODE=" & classcode & _
                     ";SECCODE=" & sym & _
                     ";ACCOUNT=" & account

            If Not isHedge And clientCode <> "" Then
                orders = orders & ";CLIENT_CODE=" & clientCode
            End If

            orders = orders & ";TYPE=M;OPERATION=" & IIf(rowSide = "Buy", "B", "S") & _
                     ";QUANTITY=" & Abs(vol) & ";PRICE=0;" & vbCrLf
        End If
    Next i

    If orders = "" Then
        Debug.Print "Нет заявок для записи на листе " & ws.Name
        Exit Function
    End If

    f = FreeFile
    Open triPath For Output Shared As #f
    Print #f, orders
    Close #f

    Debug.Print "Записаны заявки (" & side & ") в " & triPath & ":" & vbCrLf & orders
    If OPEN_TRI_IN_NOTEPAD Then Shell "notepad.exe """ & triPath & """", vbNormalFocus

    WriteOrders = orders
    Exit Function

Fail:
    ' Раньше ошибки скрывались через On Error Resume Next, и заявка могла молча не уйти
    Debug.Print "Ошибка записи заявок: " & Err.Description & " (" & Now & ")"
    ' Без MsgBox: модальное окно остановило бы таймеры до нажатия кнопки
    Application.StatusBar = "ОШИБКА заявок " & Format(Now, "HH:MM:SS") & ": " & Err.Description
    On Error Resume Next
    If f <> 0 Then Close #f
End Function
