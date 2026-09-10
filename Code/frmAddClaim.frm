VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddClaim 
   Caption         =   "Add Claim"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   OleObjectBlob   =   "frmAddClaim.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    
    If Trim(txtClaimID.Value) = "" Or Trim(txtClaimSite.Value) = "" _
       Or Trim(txtProviderName.Value) = "" Or Trim(txtClaimQuery.Value) = "" _
       Or Trim(dtCreationDate.Value) = "" Then
        MsgBox "ClaimID, Site, Provider Name, Query and Creation Date are all required.", vbExclamation
        Exit Sub
    End If
    
    If Not IsDate(dtCreationDate.Value) Then
        MsgBox "Creation Date is not a valid date.", vbExclamation
        Exit Sub
    End If

    If AddClaim(Trim(txtClaimID.Value), Trim(txtClaimSite.Value), _
                Trim(txtProviderName.Value), Trim(txtClaimQuery.Value), _
                CDate(dtCreationDate.Value)) Then
        MsgBox "Claim added.", vbInformation
        
        resetAddClaim
        
        'Unload Me
        
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub resetAddClaim()
        
    txtClaimID.Value = ""
    txtClaimSite.Value = ""
    txtProviderName.Value = ""
    txtClaimQuery.Value = ""
    dtCreationDate.Value = ""
    
End Sub

