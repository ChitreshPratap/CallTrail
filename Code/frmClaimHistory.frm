VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClaimHistory 
   Caption         =   "Call History"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10140
   OleObjectBlob   =   "frmClaimHistory.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClaimHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
    Dim data As Variant
    data = GetHistoryForClaim(Trim(txtClaimID.Value))
    lstHistory.Clear
    If IsEmpty(data) Then
        lstHistory.AddItem "No Data Found."
        MsgBox "No history found for this claim.", vbInformation
    Else
        Dim i As Long
        For i = 1 To UBound(data, 1)
            
            lstHistory.AddItem data(i, 1)
            lstHistory.List(lstHistory.ListCount - 1, 1) = data(i, 2)
            lstHistory.List(lstHistory.ListCount - 1, 2) = data(i, 3)
            lstHistory.List(lstHistory.ListCount - 1, 3) = data(i, 4)
        
        Next i

    End If
    
End Sub
