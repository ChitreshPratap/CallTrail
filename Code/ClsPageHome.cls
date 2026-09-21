VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsPageHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsPageHome - the landing tab.
'
' Shows who is signed in, whether they're registered, and what the app
' can do.
'
' NOTE ON THE IMAGE: imgHomeBanner displays a STATIC picture set in the
' designer. MSForms Image controls render only the first frame of a GIF -
' they do not animate, and no property changes that. If you ever want
' motion here, the only reliable route in VBA is swapping frame images on
' an Application.OnTime tick (~1 second floor); a WebBrowser control does
' animate GIFs but runs on the deprecated Internet Explorer engine and is
' often absent or policy-blocked on current machines.
' =====================================================================


Implements IPage

Private m_ctx As clsAppContext
Private m_pg As Msforms.Page
Private m_arranged As Boolean

' Feature rows - caption label and its description
Private Const FEATURE_COUNT As Long = 6
Private m_featureLabels(1 To FEATURE_COUNT) As Msforms.Label

Private WithEvents cmdGoLogCall As Msforms.CommandButton
Attribute cmdGoLogCall.VB_VarHelpID = -1
'Attribute cmdGoLogCall.VB_VarHelpID = -1
Private WithEvents cmdRefreshUser As Msforms.CommandButton
Attribute cmdRefreshUser.VB_VarHelpID = -1
'Attribute cmdRefreshUser.VB_VarHelpID = -1

Private lblWelcome As Msforms.Label
Private lblUserName As Msforms.Label
Private lblRole As Msforms.Label
Private lblRegStatus As Msforms.Label
Private lblGuestNotice As Msforms.Label
Private imgHome As Msforms.Image

Private Property Get IPage_PageTitle() As String
    IPage_PageTitle = "Home"
End Property

Private Sub IPage_InitPage(ctx As clsAppContext, pg As Msforms.Page)
    Set m_ctx = ctx
    Set m_pg = pg
    BindControls
End Sub

Private Sub IPage_OnActivate()
    If Not m_arranged Then
        ArrangeUI
        m_arranged = True
    End If

    ShowUserStatus
    m_ctx.SetStatus "Welcome to CallTrail."
End Sub

' =====================================================================
' Binding
' =====================================================================
Private Sub BindControls()
    Dim i As Long

    Set lblWelcome = BindControl(m_pg, "lblHomeWelcome", IPage_PageTitle)
    Set lblUserName = BindControl(m_pg, "lblHomeUserName", IPage_PageTitle)
    Set lblRole = BindControl(m_pg, "lblHomeRole", IPage_PageTitle)
    Set lblRegStatus = BindControl(m_pg, "lblHomeRegStatus", IPage_PageTitle)
    Set lblGuestNotice = BindControl(m_pg, "lblHomeGuestNotice", IPage_PageTitle)
    Set imgHome = BindControl(m_pg, "imgHomeBanner", IPage_PageTitle)

    Set cmdGoLogCall = BindOptional(m_pg, "cmdHomeStart")
    Set cmdRefreshUser = BindOptional(m_pg, "cmdHomeRefreshUser")

    For i = 1 To FEATURE_COUNT
        Set m_featureLabels(i) = BindControl(m_pg, "lblHomeFeature" & i, IPage_PageTitle)
    Next i
End Sub

' =====================================================================
' Layout
' =====================================================================
Private Sub ArrangeUI()
    Dim ui As New clsFormStyler
    Dim i As Long

    ui.Init m_pg, IPage_PageTitle, startLeft:=14, startTop:=10
    ui.LabelWidth = 95
    ui.fieldWidth = 300
    ui.RowHeight = 16
    ui.RowGap = 4
'    ui.UseScale m_ctx.UIScale

    ' --- who's signed in ---
    ui.Header "lblHomeWelcome", "Welcome to CallTrail"
    ui.ValueRow "lblCapHomeUser", "lblHomeUserName", "Signed in as:"
    ui.ValueRow "lblCapHomeRole", "lblHomeRole", "Role:"
    ui.ValueRow "lblCapHomeReg", "lblHomeRegStatus", "Status:"

    ui.Note "lblHomeGuestNotice", "", height:=30

    ui.Gap 6
    ui.BeginButtonRow
    ui.Button "cmdHomeStart", "Start Calling", 90
    ui.Button "cmdHomeRefreshUser", "Re-check Registration", 130
    ui.EndButtonRow

    ' --- what the app does ---
    ui.Gap
    ui.Header "lblHomeFeaturesHeader", "What you can do"
    For i = 1 To FEATURE_COUNT
        ui.Note "lblHomeFeature" & i, FeatureText(i), height:=16, width:=380
    Next i

    ' --- the image, in its own column on the right ---
    ui.NewColumn 420
    ui.FullWidth "imgHomeBanner", 200, 150
End Sub

Private Function FeatureText(ByVal i As Long) As String
    
    Select Case i
        Case 1: FeatureText = "Log Call  -  find a claim, see previous callers' notes, record yours"
        Case 2: FeatureText = "Search & Edit  -  full detail and call history for one claim"
        Case 3: FeatureText = "View Claims  -  browse and filter by status and date"
        Case 4: FeatureText = "Add Claim  -  enter a single claim, or bulk-add from a sheet"
        Case 5: FeatureText = "Admin  -  reopen closed claims, change sites, archive old records"
        Case 6: FeatureText = "Download  -  take a snapshot of the database for reporting"
    End Select
    
End Function

' =====================================================================
' User status - registered, guest, admin
' =====================================================================
Private Sub ShowUserStatus()
    lblUserName.caption = m_ctx.userName

    If m_ctx.IsRegistered Then
        lblRole.caption = m_ctx.Role
        lblRole.foreColor = RGB(0, 110, 0)
        lblRegStatus.caption = "Registered"
        lblRegStatus.foreColor = RGB(0, 110, 0)
        lblGuestNotice.caption = "All features are available to you." & _
            IIf(m_ctx.DefaultLocation <> "", _
                "  New claims will default to site: " & m_ctx.DefaultLocation & ".", "")
        lblGuestNotice.foreColor = RGB(90, 90, 90)
    Else
        lblRole.caption = "Guest"
        lblRole.foreColor = RGB(170, 0, 0)
        lblRegStatus.caption = "NOT REGISTERED"
        lblRegStatus.foreColor = RGB(170, 0, 0)

        ' Name the user explicitly. "You are not registered" leaves people
        ' guessing which account the app actually saw - which matters on
        ' shared machines and with multiple domain accounts.
        lblGuestNotice.caption = "User '" & m_ctx.userName & "' is not registered in the " & _
            "database. Every other tab is locked. Please contact your administrator to " & _
            "register this username, then use Re-check Registration."
        lblGuestNotice.foreColor = RGB(170, 0, 0)
    End If
End Sub

' =====================================================================
' Buttons
' =====================================================================
Private Sub cmdGoLogCall_Click()
    
    If Not m_ctx.IsRegistered Then
        MsgBox "User '" & m_ctx.userName & "' is not registered." & vbCrLf & vbCrLf & _
               "Ask your administrator to add this username to the Users sheet.", _
               vbExclamation, "Not Registered"
        Exit Sub
    End If
    ' The shell listens for this and switches tabs
    m_ctx.SetStatus "GOTO:LogCall"
    
End Sub

' Lets someone who has just been registered pick it up without
' restarting Excel - otherwise the cached lookup keeps saying Guest.
Private Sub cmdRefreshUser_Click()
    
    m_ctx.ResetUserCache
    ShowUserStatus

    If m_ctx.IsRegistered Then
        MsgBox "Registration found. Role: " & m_ctx.Role & vbCrLf & vbCrLf & _
               "Close and reopen the app to unlock the other tabs.", _
               vbInformation, "Registered"
    Else
        MsgBox "User '" & m_ctx.userName & "' is still not registered.", _
               vbExclamation, "Not Registered"
    End If
    
End Sub

