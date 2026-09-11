Attribute VB_Name = "modMain"
Option Explicit

' =====================================================================
' modMain  (CallTrail app)
' The only module that should be wired to buttons/ribbon/shortcut keys.
' Each Sub here just shows a form; the forms call modDataAccess.
' Add new features by adding one Sub here + one form, without touching
' modDataAccess or modUtils.
' =====================================================================

' PRIMARY ENTRY POINT - the tabbed application shell.
' Wire your main button to this.
' =====================================================================
Public Sub ShowApp()
    UserFormMain.Show
End Sub


' ---------------------------------------------------------------------
' The individual entry points below still work and are kept deliberately:
' they let you wire a button straight to one screen, and they are what
' you would use if you ever want a single-purpose shortcut without the
' whole shell. The tabbed app does not depend on them.
' ---------------------------------------------------------------------

'Public Sub ShowAddClaimForm()
'    frmAddClaim.Show
'End Sub

' Search ONE claim: full detail, read-only call history, and editable
' detail fields for corrections
Public Sub ShowSearchClaim()
    frmSearchClaim.Show
End Sub



'Show the form to add claim manually
Public Sub ShowAddClaimForm()
    
    'frmAddClaim.Show
    UserFormMain.Show
    
End Sub

Public Sub ShowAddClaimForm_tab(tabNumber As Integer)
    
    'frmAddClaim.Show
    Dim uform As UserFormMain
    Set uform = New UserFormMain
    uform.setPage tabNumber
    uform.Show
    Unload uform
    Set uform = Nothing
    
End Sub


Public Sub ShowUpdateClaimForm()
    frmUpdateClaim.Show
End Sub

'' Search ONE claim: full detail, read-only call history, and editable
'' detail fields for corrections
'Public Sub ShowSearchClaim()
'
'    frmSearchClaim.Show
'
'End Sub

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
    Frmviewclaims.Show
End Sub

' --- Bulk claim entry via the bulkClaimAdd sheet (see modBulkImport) ---

' Creates/clears the bulkClaimAdd sheet
Public Sub BulkSheetSetup()
    modBulkImport.SetupBulkSheet
End Sub

' Optional: load a CSV/xlsx of claims into bulkClaimAdd
Public Sub BulkLoadFromFile()
    modBulkImport.LoadClaimsFromFile
End Sub

' Dry run - flags problems, changes nothing
Public Sub BulkCheck()
    modBulkImport.CheckBulkClaims
End Sub

' MAIN action: validate, import valid rows (deleting them from the
' sheet), leave invalid rows behind with their error text
Public Sub BulkAddClaims()
    modBulkImport.ProcessBulkClaims
End Sub
