Attribute VB_Name = "SignalWatcher"
'===========================================================
' Проверка новых сигналов каждые 5 секунд и формирование
' заявок по стратегиям: фьючерсы, акции, опционы.
'===========================================================
Option Explicit

Public NextCheck As Date

Public Sub StartSignalWatcher()
    NextCheck = Now + TimeSerial(0, 0, CHECK_INTERVAL_SEC)
    Application.OnTime NextCheck, "CheckSignals"
End Sub

Public Sub StopSignalWatcher()
    On Error Resume Next
    If NextCheck <> 0 Then
        Application.OnTime EarliestTime:=NextCheck, Procedure:="CheckSignals", Schedule:=False
    End If
    NextCheck = 0
End Sub

Public Sub CheckSignals()
    On Error GoTo Cleanup
    Application.EnableEvents = False

    ProcessSheet SHEET_SIGNALS_FUTURES, SHEET_POS_FUTURES, SHEET_OUT_FUTURES_LE, SHEET_OUT_FUTURES_LX
    ProcessSheet SHEET_SIGNALS_STOCKS, SHEET_POS_STOCKS, SHEET_OUT_STOCKS_LE, SHEET_OUT_STOCKS_LX
    ProcessSheet SHEET_SIGNALS_OPTIONS, SHEET_POS_OPTIONS, SHEET_OUT_OPTIONS_LE, SHEET_OUT_OPTIONS_LX

Cleanup:
    If Err.Number <> 0 Then Debug.Print "Ошибка CheckSignals: " & Err.Description & " (" & Now & ")"
    Application.EnableEvents = True
    NextCheck = Now + TimeSerial(0, 0, CHECK_INTERVAL_SEC)
    Application.OnTime NextCheck, "CheckSignals"
End Sub

' Обрабатывает ОДИН необработанный сигнал на листе за проход.
' Каждый сигнал перезаписывает файл заявок, поэтому несколько сигналов подряд
' за один проход могли затереть друг друга до того, как торговая система их прочитает.
' Следующий сигнал будет обработан через CHECK_INTERVAL_SEC.
Private Sub ProcessSheet(ByVal signalsSheet As String, ByVal posSheet As String, _
                         ByVal outLE As String, ByVal outLX As String)
    Dim ws As Worksheet, cell As Range
    Dim lastRow As Long, signal As String

    Set ws = ThisWorkbook.Sheets(signalsSheet)
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For Each cell In ws.Range("C2:C" & lastRow)
        signal = UCase$(Trim$(CStr(cell.Value)))

        If signal <> "" And Right$(signal, 5) <> "_DONE" Then
            If IsStale(ws.Cells(cell.Row, 1).Value) Then
                ' Старый сигнал (например, пришел, пока Excel был закрыт) не исполняем по рынку
                cell.Value = signal & "_STALE_DONE"
                Debug.Print "Пропущен устаревший сигнал " & signal & " на " & signalsSheet
            Else
                Select Case signal
                    Case "LE", "LE_ADD"
                        GenerateOrders posSheet, outLE, "Buy"
                        cell.Value = signal & "_DONE"
                        Exit Sub
                    Case "LX"
                        GenerateOrders posSheet, outLX, "Sell"
                        cell.Value = signal & "_DONE"
                        Exit Sub
                    Case Else
                        cell.Value = signal & "_UNKNOWN_DONE"
                        Debug.Print "Неизвестный сигнал: " & signal & " на " & signalsSheet
                End Select
            End If
        End If
    Next cell
End Sub

Private Function IsStale(ByVal received As Variant) As Boolean
    If MAX_SIGNAL_AGE_MIN <= 0 Then Exit Function
    If Not IsDate(received) Then Exit Function
    IsStale = (Now - CDate(received)) * 24 * 60 > MAX_SIGNAL_AGE_MIN
End Function

' Переносит позиции стратегии на лист вывода и формирует рыночные заявки.
' Одна процедура вместо шести одинаковых для каждого типа инструмента и направления.
Public Sub GenerateOrders(ByVal posSheet As String, ByVal outSheet As String, ByVal side As String)
    Dim wsMain As Worksheet, wsOut As Worksheet
    Dim lastRowMain As Long, outRow As Long, j As Long
    Dim stamp As String, firstCol As String

    Set wsMain = ThisWorkbook.Sheets(posSheet)
    Set wsOut = ThisWorkbook.Sheets(outSheet)

    wsOut.Range("A2:B100").ClearContents
    wsOut.Range("J2:J100").ClearContents

    lastRowMain = wsMain.Cells(wsMain.Rows.Count, "A").End(xlUp).Row
    outRow = 1
    stamp = "Обновлено: " & Format(Now, "dd.mm.yyyy HH:MM:SS")

    For j = 2 To lastRowMain
        firstCol = UCase$(Trim$(CStr(wsMain.Cells(j, 1).Value)))
        If firstCol <> "" And firstCol <> "ИТОГО" Then
            outRow = outRow + 1
            wsOut.Cells(outRow, 1).Value = wsMain.Cells(j, 1).Value    ' дата или время
            wsOut.Cells(outRow, 2).Value = wsMain.Cells(j, 2).Value    ' тикер
            wsOut.Cells(outRow, 10).Value = wsMain.Cells(j, 7).Value   ' объем
            wsOut.Cells(outRow, 12).Value = stamp
        End If
    Next j

    If outRow >= 2 Then
        WriteOrders wsOut, side, 2, outRow
    Else
        Debug.Print "Нет позиций для заявки на " & posSheet
    End If
End Sub
