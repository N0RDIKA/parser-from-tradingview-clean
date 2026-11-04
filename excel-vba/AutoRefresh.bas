Attribute VB_Name = "AutoRefresh"
'===========================================================
' Обновление данных из Google Sheets (Power Query) и раскладка
' новых сигналов по листам стратегий. Запуск каждые 30 секунд.
'===========================================================
Option Explicit

Public NextRefreshTime As Date

Public Sub StartAutoRefresh()
    RefreshAndDistribute
End Sub

Public Sub StopAutoRefresh()
    On Error Resume Next
    If NextRefreshTime <> 0 Then
        Application.OnTime EarliestTime:=NextRefreshTime, Procedure:="RefreshAndDistribute", Schedule:=False
    End If
    NextRefreshTime = 0
End Sub

Public Sub RefreshAndDistribute()
    Dim wsSrc As Worksheet
    Dim i As Long, lastSrc As Long
    Dim key As String, target As String

    On Error GoTo Fail

    Set wsSrc = ThisWorkbook.Sheets(SHEET_SOURCE)
    RefreshQueries wsSrc

    lastSrc = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastSrc
        If InStr(1, wsSrc.Cells(i, 3).Value, "_DONE", vbTextCompare) = 0 Then
            key = Trim$(CStr(wsSrc.Cells(i, 2).Value))

            Select Case key
                Case SIGNAL_KEY_FUTURES: target = SHEET_SIGNALS_FUTURES
                Case SIGNAL_KEY_STOCKS:  target = SHEET_SIGNALS_STOCKS
                Case SIGNAL_KEY_OPTIONS: target = SHEET_SIGNALS_OPTIONS
                Case Else:               target = ""
            End Select

            If target <> "" Then CopyIfNew wsSrc, i, ThisWorkbook.Sheets(target)
        End If
    Next i

    Debug.Print "Обновление завершено: " & Now
    GoTo Reschedule

Fail:
    ' Раньше при любой ошибке таймер не перезапускался и автоматизация молча останавливалась
    Debug.Print "Ошибка обновления: " & Err.Description & " (" & Now & ")"

Reschedule:
    NextRefreshTime = Now + TimeSerial(0, 0, REFRESH_INTERVAL_SEC)
    Application.OnTime NextRefreshTime, "RefreshAndDistribute"
End Sub

Private Sub RefreshQueries(ByVal ws As Worksheet)
    Dim qt As QueryTable, lo As ListObject
    For Each qt In ws.QueryTables
        qt.Refresh BackgroundQuery:=False
    Next qt
    For Each lo In ws.ListObjects
        On Error Resume Next
        lo.QueryTable.Refresh BackgroundQuery:=False
        On Error GoTo 0
    Next lo
End Sub

' Ключ сигнала: время получения | символ | время сигнала из TradingView.
' Столбец C (сигнал) в ключ не входит, потому что после обработки к нему дописывается _DONE.
Private Function SignalKey(ByVal ws As Worksheet, ByVal r As Long) As String
    SignalKey = CStr(ws.Cells(r, 1).Value) & "|" & CStr(ws.Cells(r, 2).Value) & "|" & CStr(ws.Cells(r, 5).Value)
End Function

' Копирует строку сигнала на лист стратегии, если ее там еще нет
Private Sub CopyIfNew(ByVal wsSrc As Worksheet, ByVal srcRow As Long, ByVal wsDest As Worksheet)
    Dim j As Long, lastDest As Long, key As String

    key = SignalKey(wsSrc, srcRow)
    lastDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row

    For j = 2 To lastDest
        If SignalKey(wsDest, j) = key Then Exit Sub
    Next j

    wsDest.Rows(lastDest + 1).Value = wsSrc.Rows(srcRow).Value
End Sub
