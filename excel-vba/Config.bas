Attribute VB_Name = "Config"
'===========================================================
' Настройки. Здесь все, что зависит от конкретного счета
' и стратегий. В публичном репозитории только заглушки.
'===========================================================
Option Explicit

' --- Интервалы опроса, секунд ---
Public Const REFRESH_INTERVAL_SEC As Long = 30   ' обновление Power Query
Public Const CHECK_INTERVAL_SEC As Long = 5      ' проверка новых сигналов

' --- Сигналы старше этого возраста не исполняются, минут (0 - без проверки) ---
Public Const MAX_SIGNAL_AGE_MIN As Long = 5

' --- Листы ---
Public Const SHEET_SOURCE As String = "Лист8"        ' выгрузка из Google Sheets через Power Query

' Входящие сигналы по стратегиям
Public Const SHEET_SIGNALS_FUTURES As String = "Лист1"
Public Const SHEET_SIGNALS_STOCKS As String = "Лист9"
Public Const SHEET_SIGNALS_OPTIONS As String = "Лист14"

' Состав позиций по стратегиям
Public Const SHEET_POS_FUTURES As String = "Лист2"
Public Const SHEET_POS_STOCKS As String = "Лист6"
Public Const SHEET_POS_OPTIONS As String = "Лист13"

' Листы вывода заявок: вход (LE) и выход (LX)
Public Const SHEET_OUT_FUTURES_LE As String = "Лист5"
Public Const SHEET_OUT_FUTURES_LX As String = "Лист3"
Public Const SHEET_OUT_STOCKS_LE As String = "Лист10"
Public Const SHEET_OUT_STOCKS_LX As String = "Лист7"
Public Const SHEET_OUT_OPTIONS_LE As String = "Лист11"
Public Const SHEET_OUT_OPTIONS_LX As String = "Лист12"

' --- Как опознать стратегию по столбцу B (ticker из алерта TradingView) ---
' Сюда подставляется формула корзины в том виде, в каком ее присылает TradingView.
Public Const SIGNAL_KEY_FUTURES As String = "FUTURES_BASKET"
Public Const SIGNAL_KEY_STOCKS As String = "STOCKS_BASKET"
Public Const SIGNAL_KEY_OPTIONS As String = "TP"

' --- Хедж фьючерсом на индекс: направление заявки инвертируется ---
Public Const HEDGE_SECCODE As String = "IMOEXF"
Public Const HEDGE_CLASSCODE As String = "SPBFUT"
Public Const HEDGE_ACCOUNT As String = "YOUR_FUTURES_ACCOUNT"

' --- Ячейки настроек на каждом листе вывода (столбец R) ---
' R2 - путь к файлу заявок .tri, R3 - CLASSCODE, R4 - ACCOUNT, R5 - CLIENT_CODE
Public Const CFG_COL As Long = 18
Public Const CFG_ROW_TRI_PATH As Long = 2
Public Const CFG_ROW_CLASSCODE As Long = 3
Public Const CFG_ROW_ACCOUNT As Long = 4
Public Const CFG_ROW_CLIENT_CODE As Long = 5

' --- Открывать файл заявок в блокноте после записи (для отладки) ---
Public Const OPEN_TRI_IN_NOTEPAD As Boolean = False
