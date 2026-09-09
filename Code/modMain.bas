Attribute VB_Name = "modMain"
Option Explicit

' =====================================================================
' modMain  (CallTrail app)
' The only module that should be wired to buttons/ribbon/shortcut keys.
' Each Sub here just shows a form; the forms call modDataAccess.
' Add new features by adding one Sub here + one form, without touching
' modDataAccess or modUtils.
' =====================================================================

Public Sub ShowAddClaimForm()
    frmAddClaim.Show
End Sub

Public Sub ShowUpdateClaimForm()
    frmUpdateClaim.Show
End Sub

Public Sub ShowAdminSiteForm()
    Dim wb As Workbook
    Dim isAdmin As Boolean

    Set wb = OpenCentralDB()
    isAdmin = IsCurrentUserAdmin(wb)
    CloseCentralDB wb, False

    If Not isAdmin Then
        MsgBox "This feature is restricted to Admin users.", vbExclamation, "Access Denied"
        Exit Sub
    End If
    frmAdminSite.Show
End Sub

Public Sub ShowClaimHistoryViewer()
    frmClaimHistory.Show
End Sub

' Browse/filter all claims with full detail + call history
Public Sub ShowViewClaims()
    frmViewClaims.Show
End Sub

' --- Bulk claim entry via the bulkClaimAdd sheet (see modBulkImport) ---

' Creates/clears the bulkClaimAdd sheet
Public Sub BulkSheetSetup()
    SetupBulkSheet
End Sub

' Optional: load a CSV/xlsx of claims into bulkClaimAdd
Public Sub BulkLoadFromFile()
    LoadClaimsFromFile
End Sub

' Dry run - flags problems, changes nothing
Public Sub BulkCheck()
    CheckBulkClaims
End Sub

' MAIN action: validate, import valid rows (deleting them from the
' sheet), leave invalid rows behind with their error text
Public Sub BulkAddClaims()
    ProcessBulkClaims
End Sub
