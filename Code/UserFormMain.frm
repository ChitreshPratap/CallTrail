VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMain 
   Caption         =   "UserForm1"
   ClientHeight    =   10095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16185
   OleObjectBlob   =   "UserFormMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public btnStyleCollection As New Collection
Public buttonHome As tsLabelHE
Public buttonProcessing As tsLabelHE
Public buttonAboutUs As tsLabelHE
Public mainWindowHeight As Long
Public mainWindowWidth As Long
Public txtBoxStyleCollection As New Collection
Public tbStyle2 As TsTextFieldStyle2

Public contAddClaim As ClsContAddClaim

Private Sub cmdCancel_Click()
    
    contAddClaim.cmdCancel_Click Me

End Sub

Private Sub cmdReset_Click()
    
    contAddClaim.resetAddClaim Me
    
End Sub

Private Sub cmdSave_Click()

    contAddClaim.cmdSave_Click Me
    
End Sub

Private Sub dtCreationDate_Change()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame1_Enter()
    MsgBox "Entered"
End Sub

Private Sub Frame1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    MsgBox "Exited"
End Sub

Private Sub frameDashboard_Click()

End Sub

Private Sub frameMain_Click()

End Sub

Private Sub framePageProcess_Click()

End Sub

Private Sub lblCloseDashboard_Click()

    Dim i As Long
    i = 300
    Do Until i = 0
        DoEvents
        i = i - 1
    Loop
    Me.frameDashboard.Width = 0
    Me.frameMain.Left = Me.frameDashboard.Width
    Me.lblShowDashboard.Visible = True
    Me.frameMain.Width = mainWindowWidth - i
'    Me.Painel_Principal.Width = Me.Width - Me.Painel_Lateral.Width
    
    With lblAppHeader
        .Left = 30
        .Width = mainWindowWidth - 30
    End With
    
    With lblAppHeaderLine
        .Left = 0
        .Width = mainWindowWidth
    End With
    
    With multiPageApp
        .Left = Me.frameMain.Left
        .Width = mainWindowWidth
    End With
    With framePageAbout
        .Left = Me.frameMain.Left
        .Width = mainWindowWidth
    End With
    
    With framePageProcess
        .Left = Me.frameMain.Left
        .Width = mainWindowWidth
    End With
    
    With framePageHome
        .Left = Me.frameMain.Left
        .Width = mainWindowWidth
    End With
    


End Sub

Private Sub lblAbout_Click()
    Me.multiPageApp.Value = 2
End Sub

Private Sub lblHome_Click()

    Me.multiPageApp.Value = 0
End Sub

Private Sub lblMenuItemAddClaim_Click()
    Me.multiPageApp.Value = 1
End Sub

Private Sub lblShowDashboard_Click()
    Dim i As Long
    For i = 0 To 300
        DoEvents
    Next
    Me.frameDashboard.Width = 170
    Me.frameMain.Left = Me.frameDashboard.Width
    Me.lblShowDashboard.Visible = False
    Me.lblAppHeader.Left = 10


End Sub

Private Sub MultiPageMain_Change()

End Sub

Private Sub TabStrip1_Change()

End Sub

Private Sub multiPageApp_Change()

End Sub

Private Sub TextBox1_Change()

End Sub


Private Sub txtClaimId_Change()

End Sub

Private Sub txtClaimQuery_Change()

End Sub

Private Sub UserForm_Initialize()
        
    Set contAddClaim = New ClsContAddClaim
    'contAddClaim.setForm Me
    
    Me.Caption = AppUtil.getAppName
    Me.StartUpPosition = 0
    
    Me.Width = Application.UsableWidth * 0.9
    Me.Height = Application.Height * 0.9
    
    mainWindowHeight = Me.Height - 30
    mainWindowWidth = Me.Width - 10
    
    With frameMain
        '.Picture = LoadPicture(ThisWorkbook.Path & "\panelPrinciple.jpg")
        .PictureSizeMode = fmPictureSizeModeStretch
        .Height = mainWindowHeight
        .Width = mainWindowWidth - Me.frameDashboard.Width
        .Left = Me.frameDashboard.Width
        .Top = 0
        .ZOrder (1)
        .BackColor = AppUtil.getThemeColor()
        
    End With
    With frameDashboard
        .Width = 170
        .Left = 0
        .Top = 0
        .Height = mainWindowHeight
        .ZOrder (1)
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.05)
        .BorderColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.05)
        
    End With
            
    With lblHome
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
        .ForeColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    End With
    'With lblProcess
    With lblMenuItemAddClaim
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
        .ForeColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    End With
    With lblAbout
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
        .ForeColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    
    End With
    With lblAppHeader
        .Width = frameMain.Width
        .Left = 10
        .Top = 5
        .Height = 50
        .BackColor = AppUtil.getThemeColor()
        .ForeColor = vbWhite
        .Caption = AppUtil.getAppName()
    End With
    
    With lblAppHeaderLine
        .Width = frameMain.Width
        .Left = 0
        .Top = lblAppHeader.Top + lblAppHeader.Height
        .Height = 1.5
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    End With
    
    With multiPageApp
        .Style = fmTabStyleNone
        .Width = frameMain.Width
        .Left = 0
        .Top = lblCloseDashboard.Height + lblCloseDashboard.Top + 30
        .Height = mainWindowHeight - lblCloseDashboard.Height - lblCloseDashboard.Top
    End With
    
    With framePageHome
        .Width = frameMain.Width
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        .Left = 0
        .Top = 0
        .Height = mainWindowHeight - lblCloseDashboard.Height - lblCloseDashboard.Top
    End With
    With framePageAbout
        .Width = frameMain.Width
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        .Left = 0
        .Top = 0
        .Height = mainWindowHeight - lblCloseDashboard.Height - lblCloseDashboard.Top
    End With
    
    With framePageProcess
        .Width = frameMain.Width
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        .Left = 0
        .Top = 0
        .Height = mainWindowHeight - lblCloseDashboard.Height - lblCloseDashboard.Top
    End With
    multiPageApp.Value = 0
            
 '   txtBoxStyleCollection.Add Styler.getStyledBox(textBoxName, "Enter Name")
'    txtBoxStyleCollection.Add Styler.getStyledBox(textBoxDOB, "Select DOB")
    
    
    
    'Style Page Add Claim
    txtBoxStyleCollection.Add Styler.getStyledBoxULine(Me.txtClaimID, "Claim Id*")
    txtBoxStyleCollection.Add Styler.getStyledBoxULine(Me.txtClaimSite, "Claim Site*")
    txtBoxStyleCollection.Add Styler.getStyledBoxULine(Me.txtProviderName, "Provider Name*")
    txtBoxStyleCollection.Add Styler.getStyledBoxULine(Me.dtCreationDate, "Creation Date*")

    
    Set tbStyle2 = New TsTextFieldStyle2
    tbStyle2.init Me.txtClaimQuery, "Claim Query (Required) "
      
    Set buttonHome = New tsLabelHE
    buttonHome.setMukhLabel lblHome
    buttonHome.setIncDecInFont -2
    
    Set buttonProcessing = New tsLabelHE
    buttonProcessing.setMukhLabel lblMenuItemAddClaim
    buttonProcessing.setIncDecInFont 4
    buttonProcessing.setOnHoverForeColor (400)
    
    
    Set buttonAboutUs = New tsLabelHE
    buttonAboutUs.setMukhLabel lblAbout
    buttonAboutUs.setIncDecInFont 10
    buttonAboutUs.setOnHoverBackColor (1000)
    
    'buttonAboutUs.init lblAbout
    'btnStyleCollection.Add Styler.getStyleButton(btnSelectDate)
            
'    Dim lblF As Variant
'    Set lblF = lblAbout.Font
'    lblF.Size = 24
'    textBoxName.Font = lblF
'    Debug.Print "Hello"
End Sub

Sub setPage(pageNumber As Integer)

    Me.multiPageApp.Value = pageNumber
    
End Sub
