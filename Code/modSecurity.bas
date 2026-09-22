Attribute VB_Name = "modSecurity"
Option Explicit

' =====================================================================
' modSecurity
' Applies and checks the file-open password on the central database.
'
' HOW THE PROTECTION WORKS
' The database is saved with a file-open password. For .xlsx files
' Excel then ENCRYPTS the whole file (AES). Double-clicking it on the
' share drive brings up a password prompt; without the password there
' is nothing to read - not in Excel, not by unzipping the file.
'
' Every open the app makes goes through modUtils.OpenDBCore, which
' supplies DbPassword() from modConfig. Saves made by the app keep the
' password - Excel re-encrypts on every Save of a workbook opened with
' one - so the file is protected at rest at all times, not just once.
'
' THE HONEST LIMITS - read these before relying on it
'
'  1. The password is in the macro file. Anyone who can open that file's
'     VBA editor can read modConfig. LOCK THE VBA PROJECT before you
'     distribute (VBA editor > Tools > VBAProject Properties >
'     Protection > "Lock project for viewing" + a DIFFERENT password).
'
'  2. VBA project locking is weak. Tools that strip it are freely
'     available. So this design stops people opening the database
'     casually or by accident; it does not stop a determined person who
'     has the macro file and wants in. For that you need real access
'     control - share-folder permissions, or moving the data into a
'     database server with per-user logins.
'
'  3. Anything the app COPIES OUT is no longer protected. The DL_ sheets
'     from modDownload sit unencrypted in the user's own macro workbook.
'
'  4. Lose the password and the data is gone. There is no recovery for
'     Excel's AES encryption. Keep the password somewhere safe that is
'     NOT the macro file - a password manager, or sealed with IT.
' =====================================================================

' ---------------------------------------------------------------------
' Admin-only. Sets the database password to whatever DbPassword()
' currently returns in modConfig.
'
' Use it:
'   - ONCE at setup, to encrypt a database that has no password yet
'     (leave the "current password" prompt blank), and
'   - whenever you ROTATE the password: change DB_PASSWORD in modConfig,
'     run this, enter the OLD password when asked, then send every user
'     the new macro file. Old copies will get a clear "password does not
'     match" message rather than a cryptic failure.
' ---------------------------------------------------------------------
Public Sub SetDatabasePassword()
    Dim wb As Workbook
    Dim currentPwd As String
    Dim resp As VbMsgBoxResult

    If DbPassword() = "Change-Me-Before-Use-2026!" Then
        MsgBox "DB_PASSWORD in modConfig is still the placeholder." & vbCrLf & vbCrLf & _
               "Set a real password there first, then run this again.", _
               vbExclamation, "Placeholder Password"
        Exit Sub
    End If

    resp = MsgBox("This encrypts the central database with the password set in modConfig:" & vbCrLf & vbCrLf & _
                  DB_PATH & vbCrLf & vbCrLf & _
                  "Before continuing:" & vbCrLf & _
                  "  - make sure nobody else has the app or the file open" & vbCrLf & _
                  "  - take a backup copy of the file" & vbCrLf & _
                  "  - store the password somewhere safe OUTSIDE this macro file." & vbCrLf & _
                  "    A lost password cannot be recovered." & vbCrLf & vbCrLf & _
                  "Continue?", vbYesNo + vbExclamation, "Set Database Password")
    If resp = vbNo Then Exit Sub

    currentPwd = InputBox("Enter the database's CURRENT password." & vbCrLf & vbCrLf & _
                          "Leave blank if the file has no password yet (first-time setup).", _
                          "Current Password")
    ' InputBox returns "" for both Cancel and an empty entry - both mean
    ' "no current password", which is the right reading for first setup.

    On Error GoTo OpenFail
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(FileName:=DB_PATH, UpdateLinks:=0, readOnly:=False, _
                            Notify:=False, Password:=currentPwd)
    On Error GoTo Fail

    ' Someone else holds it: Excel opened a read-only copy, and SaveAs
    ' over the original would fail or clash with their save.
    If wb.readOnly Then
        wb.Close saveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Someone else has the database open. Ask everyone to close the app, then try again.", _
               vbExclamation, "Database In Use"
        Exit Sub
    End If

    If Not IsCurrentUserAdmin(wb) Then
        wb.Close saveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "Only Admin users can change the database password.", vbExclamation, "Access Denied"
        Exit Sub
    End If

    ' SaveAs over the same path with a Password is what encrypts it.
    ' xlOpenXMLWorkbook keeps it .xlsx.
    Application.DisplayAlerts = False
    wb.SaveAs FileName:=DB_PATH, FileFormat:=xlOpenXMLWorkbook, Password:=DbPassword()
    Application.DisplayAlerts = True
    wb.Close saveChanges:=False

    Application.ScreenUpdating = True

    ' Prove it round-trips before telling anyone it worked.
    If VerifyDatabaseAccessSilent() Then
        MsgBox "Database password set and verified." & vbCrLf & vbCrLf & _
               "Next steps:" & vbCrLf & _
               "  1. Lock this macro file's VBA project (Tools > VBAProject Properties > Protection)." & vbCrLf & _
               "  2. Distribute this macro file to users." & vbCrLf & _
               "  3. Any older copies of the macro file will now stop working - replace them.", _
               vbInformation, "Password Set"
    Else
        MsgBox "The password was applied but the app could not re-open the database with it." & vbCrLf & vbCrLf & _
               "Restore your backup and check DB_PASSWORD in modConfig.", _
               vbCritical, "Verification Failed"
    End If
    Exit Sub

OpenFail:
    Application.ScreenUpdating = True
    MsgBox "Could not open the database with the current password you entered." & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Open Failed"
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close saveChanges:=False
    On Error GoTo 0
    MsgBox "Setting the password failed: " & Err.Description & vbCrLf & vbCrLf & _
           "Check the file still opens with its previous password before retrying.", _
           vbCritical, "Password Change Failed"
End Sub

' ---------------------------------------------------------------------
' Anyone can run this. Confirms this copy of the macro file can open
' the database - the first thing to try when a user reports errors
' after a password rotation.
' ---------------------------------------------------------------------
Public Sub VerifyDatabaseAccess()
    If VerifyDatabaseAccessSilent() Then
        MsgBox "This copy of CallTrail can open the database.", vbInformation, "Access OK"
    Else
        MsgBox "This copy of CallTrail could NOT open the database." & vbCrLf & vbCrLf & _
               "Most likely the database password has changed and this macro file is out of date. " & _
               "Ask your administrator for the current version.", vbExclamation, "Access Failed"
    End If
End Sub

Private Function VerifyDatabaseAccessSilent() As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = OpenCentralDBReadOnly()
    If Not wb Is Nothing Then
        CloseCentralDB wb, False
        VerifyDatabaseAccessSilent = True
    End If
    On Error GoTo 0
End Function
