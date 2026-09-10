VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdateClaim 
   Caption         =   "Update Claim"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10770
   OleObjectBlob   =   "frmUpdateClaim.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUpdateClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    cboStatus.List = Array("Pending", "Closed")
    cboStatus.ListIndex = 0
End Sub

Private Sub cmdSubmit_Click()
    If Trim(txtClaimID.Value) = "" Or Trim(txtComment.Value) = "" Then
        MsgBox "Claim ID and comment are required.", vbExclamation
        Exit Sub
    End If
    If LogCallAndUpdateStatus(Trim(txtClaimID.Value), Trim(txtComment.Value), cboStatus.Value) Then
        MsgBox "Call logged.", vbInformation
        txtComment.Value = ""
    End If
    
End Sub
