'===========================================================
' Код модуля ЭтаКнига (ThisWorkbook).
' Вставляется вручную: VBA-редактор > ЭтаКнига > вставить код.
'===========================================================
Option Explicit

Private Sub Workbook_Open()
    Application.EnableEvents = True
    StartSignalWatcher
    StartAutoRefresh
End Sub

' Без остановки таймеров Excel после закрытия сам заново открывает книгу,
' чтобы выполнить запланированный Application.OnTime
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    StopSignalWatcher
    StopAutoRefresh
End Sub
