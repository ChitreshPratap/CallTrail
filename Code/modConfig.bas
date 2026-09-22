Attribute VB_Name = "modConfig"
Option Explicit

' =====================================================================
' modConfig
' All environment-specific settings live here ONLY.
' Change this one module when you move the DB file or add users -
' nothing else in the app should hard-code a path or a name.
' =====================================================================

' >>> UPDATE THIS to the real UNC/share path before distributing <<<
Public Const DB_PATH As String = "\\SharedDrive\ClaimTracker\Claim_Calling_Tracker.xlsx"

' ---------------------------------------------------------------------
' DATABASE PASSWORD
'
' The central workbook is saved with a file-open password, so Excel
' encrypts it (AES) and nobody can open it by double-clicking - Excel
' asks for the password and there is nothing to read without it.
'
' >>> CHANGE THIS before running modSecurity.SetDatabasePassword <<<
'
' Kept Private and exposed through DbPassword() so there is exactly one
' place it lives. Read the honest limits in modSecurity: anyone who can
' open THIS macro file's VBA editor can read this line, so lock the VBA
' project before distributing.
' ---------------------------------------------------------------------
Private Const DB_PASSWORD As String = "Change-Me-Before-Use-2026!"

Public Const SHEET_CLAIMS As String = "Claims"
Public Const SHEET_HISTORY As String = "History"
Public Const SHEET_USERS As String = "Users"
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_ARC_CLAIMS As String = "ArchivedClaims"
Public Const SHEET_ARC_HISTORY As String = "ArchivedHistory"

Public Const STATUS_PENDING As String = "Pending"
Public Const STATUS_CLOSED As String = "Closed"

' Retry behaviour when the shared workbook is locked by another user
Public Const LOCK_MAX_RETRIES As Long = 5
Public Const LOCK_RETRY_WAIT_SEC As Long = 3

' Accessor for the password constant above. Placed last because VBA
' requires every Const/variable declaration to sit above the first
' procedure in a module.
Public Function DbPassword() As String
    DbPassword = DB_PASSWORD
End Function
