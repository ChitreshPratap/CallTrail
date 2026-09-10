Attribute VB_Name = "modConfig"
Option Explicit

' =====================================================================
' modConfig
' All environment-specific settings live here ONLY.
' Change this one module when you move the DB file or add users -
' nothing else in the app should hard-code a path or a name.
' =====================================================================

' >>> UPDATE THIS to the real UNC/share path before distributing <<<
Public Const DB_PATH As String = "C:\Users\pc\Documents\GitHub\Caller_Tracker\DB\Claim_Calling_Tracker.xlsx"

Public Const SHEET_CLAIMS As String = "Claims"
Public Const SHEET_HISTORY As String = "History"
Public Const SHEET_USERS As String = "Users"
Public Const SHEET_CONFIG As String = "Config"

Public Const STATUS_PENDING As String = "Pending"
Public Const STATUS_CLOSED As String = "Closed"

' Retry behaviour when the shared workbook is locked by another user
Public Const LOCK_MAX_RETRIES As Long = 5
Public Const LOCK_RETRY_WAIT_SEC As Long = 3
