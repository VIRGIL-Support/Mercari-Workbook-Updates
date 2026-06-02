Attribute VB_Name = "modConstants"
Option Explicit

' =====================================================
' APPLICATION INFORMATION
' =====================================================

Public Const APP_NAME As String = "Mercari System"

Public Const APP_VERSION As String = "1.0.0"

' =====================================================
' WORKSHEET NAMES
' =====================================================

Public Const WS_HOME As String = "HOME"
Public Const WS_POLICIES As String = "POLICIES"
Public Const WS_INVENTORY As String = "INVENTORY"
Public Const WS_DATA As String = "DATA"
Public Const WS_TABLES As String = "TABLES"
Public Const WS_SETTINGS As String = "SETTINGS"
Public Const WS_LOOKUPS As String = "LOOKUPS"
Public Const WS_SOLD_ITEMS As String = "SOLD ITEMS"

' =====================================================
' TABLE NAMES
' =====================================================

Public Const TBL_LOOKUPS As String = "tblLookups"

' =====================================================
' INVENTORY SHEET COLUMNS
' =====================================================

Public Const COL_ITEM_NUMBER As Long = 1
Public Const COL_ITEM_NAME As Long = 2
Public Const COL_ITEM_PRICE As Long = 3
Public Const COL_DATE_SOLD As Long = 4

Public Const COL_START_EDIT As Long = 5
Public Const COL_COPY_FOR_AI As Long = 6
Public Const COL_DETAILS As Long = 7
Public Const COL_VIEW_DETAILS As Long = 8
Public Const COL_DETAILS_FOLDER As Long = 9
Public Const COL_VIEW_ALL_FOLDERS As Long = 10

Public Const COL_STATUS As Long = 11

' =====================================================
' STATUS VALUES
' =====================================================

Public Const STATUS_NEW As String = "NEW"

Public Const STATUS_IN_PROGRESS As String = "IN PROGRESS"

Public Const STATUS_PROMPT_GENERATED As String = "PROMPT GENERATED"

Public Const STATUS_DETAILS_CREATED As String = "DETAILS FILE CREATED"

Public Const STATUS_SOLD As String = "SOLD"

' =====================================================
' FILE NAMES
' =====================================================

Public Const ERROR_LOG_FILE As String = "ErrorLog.txt"

' =====================================================
' APPLICATION SETTINGS
' =====================================================

Public Const MAX_ITEM_TITLE_LENGTH As Long = 80

Public Const DEFAULT_MAX_PHOTOS As Long = 12
