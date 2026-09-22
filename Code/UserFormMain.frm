VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMain 
   Caption         =   "UserForm1"
   ClientHeight    =   10095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17355
   OleObjectBlob   =   "UserFormMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmMain - code-behind for the tabbed shell
'
' Paste ALL of this into the code module of a UserForm named frmMain.
' See Setup Guide section 7 for the control layout.
'
' This module is deliberately THIN. It owns the MultiPage, the status
' bar and the page registry - nothing else. All actual behaviour lives
' in the clsPage* controllers, which is what makes adding a sixth tab a
' matter of writing one class and adding one line to RegisterPages.
' =====================================================================
Option Explicit

Private m_ctx As clsAppContext
Private m_pages As Collection        ' IPage controllers, in tab order
Private m_initialised As Boolean


Public btnStyleCollection As New Collection
Public collection_navigationButton As New Collection
Public buttonHome As tsLabelHE
Public buttonProcessing As tsLabelHE
Public buttonAboutUs As tsLabelHE
Public mainWindowHeight As Long
Public mainWindowWidth As Long
Public txtBoxStyleCollection As New Collection
Public tbStyle2 As TsTextFieldStyle2

'Public contAddClaim As ClsContAddClaim

'Dim var_viewClaimsTab As ClsViewClaims



'Private Sub cmdCancel_Click()
'
'    contAddClaim.cmdCancel_Click Me
'
'End Sub

'Private Sub cmdClose_Click()
'
'    Me.frameFilterViewClaims.Visible = False
'
'End Sub

'Private Sub cmdReset_Click()
'
'    contAddClaim.resetAddClaim Me
'
'End Sub

'Private Sub cmdSave_Click()
'
'    contAddClaim.cmdSave_Click Me
'
'End Sub


Private Sub Frame1_Enter()
    MsgBox "Entered"
End Sub

Private Sub Frame1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    MsgBox "Exited"
End Sub


Private Sub cmdViewClose_Click()
    Me.frameFilterViewClaims.Visible = False
End Sub


Private Sub framePageViewRecords_Click()

End Sub

Private Sub lblBtnApplyFilters_Click()

    Me.frameFilterViewClaims.Visible = True

End Sub

Private Sub lblCloseDashboard_Click()

    Dim i As Long
    i = 300
    Do Until i = 0
        DoEvents
        i = i - 1
    Loop
    Me.frameDashboard.width = 0
    Me.frameMain.Left = Me.frameDashboard.width
    Me.lblShowDashboard.Visible = True
    Me.frameMain.width = mainWindowWidth - i
'    Me.Painel_Principal.Width = Me.Width - Me.Painel_Lateral.Width
    
    With lblAppHeader
        .Left = 30
        .width = mainWindowWidth - 30
    End With
    
    With lblAppHeaderLine
        .Left = 0
        .width = mainWindowWidth
    End With
    
    With multiPageApp
        .Left = Me.frameMain.Left
        .width = mainWindowWidth
    End With
    With framePageAbout
        .Left = Me.frameMain.Left
        .width = mainWindowWidth
    End With
    
'    With framePageProcess
'        .Left = Me.frameMain.Left
'        .width = mainWindowWidth
'    End With
'
    With framePageHome
        .Left = Me.frameMain.Left
        .width = mainWindowWidth
    End With
    


End Sub

Public Sub gotoPageNumber(ByVal pageNumber As String)
    Dim pageId As String
    
    Select Case pageNumber:
    
        Case 1:
            pageId = "pageAddRecord"
            GoToPage pageId
            
    End Select
            
End Sub


Private Sub GoToPage(ByVal pageName As String)
    
    Dim i As Long
    On Error Resume Next
    For i = 0 To multiPageApp.Pages.Count - 1
        If multiPageApp.Pages(i).Name = pageName Then
            If multiPageApp.Pages(i).enabled And multiPageApp.Pages(i).Visible Then
                multiPageApp.value = i
            End If
            Exit For
        End If
    Next i
    On Error GoTo 0
End Sub

'Private Sub GoToPage(pgNumber As Integer)
'    Me.multiPageApp.value = pgNumber
'End Sub
'

Private Sub lblAbout_Click()
    Me.multiPageApp.value = 2
End Sub

Private Sub lblMenuItemAddClaim_Click()
    
    GoToPage "pageAddRecord"
    
End Sub
Private Sub lblMenuItemViewClaims_Click()
    
    GoToPage "pageViewRecords"
    
End Sub
Private Sub lblMenuItemLogCall_Click()
    GoToPage "pageLogCall"
End Sub

Private Sub lblMenuItemAdmin_Click()
    GoToPage "pageAdmin"
End Sub

Private Sub lblHome_Click()
        
    Me.multiPageApp.value = 0
    
End Sub

Private Sub lblMenuItemSearch_Click()

    GoToPage "pageSearch"
    
End Sub

Private Sub lblSearchDatePickerCreated_On_Click()

End Sub

Private Sub lblSearchDatePickerCreatedOn_Click()

End Sub

Private Sub lblShowDashboard_Click()
    
    Dim i As Long
    For i = 0 To 300
        DoEvents
    Next
    Me.frameDashboard.width = 170
    Me.frameMain.Left = Me.frameDashboard.width
    Me.lblShowDashboard.Visible = False
    Me.lblAppHeader.Left = 10


End Sub


Private Sub UserForm_Initialize()
        
    On Error GoTo Fail
    Set m_ctx = New clsAppContext
    m_ctx.SetStatusSink Me
    
    ' Shell FIRST, so the MultiPage is at its final size before any page
    ' arranges itself inside it. Pages that size controls against the
    ' page width would otherwise lay out against the designer's size.
    ' ArrangeShell also calibrates the UI scale and stores it on the
    ' context for the tabs to reuse.
    
    ArrangeShell
    
    RegisterPages
    
    ApplyTabVisibility
    
    m_initialised = True
    
    ' Activate the first visible tab so it populates immediately
    Me.multiPageApp.value = FirstVisiblePageIndex()
    
    
     'MultiPage.Value = FirstVisiblePageIndex()
     
    ActivateCurrentPage
    
'    Me.multiPageApp.value = 2
    
    Exit Sub

    'Set contAddClaim = New ClsContAddClaim
    'contAddClaim.setForm Me
    
            
    
    
    With framePageHome
        .width = frameMain.width
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        .Left = 0
        .Top = 0
        .height = mainWindowHeight - lblCloseDashboard.height - lblCloseDashboard.Top
    End With
    
    With framePageAbout
        .width = frameMain.width
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        .Left = 0
        .Top = 0
        .height = mainWindowHeight - lblCloseDashboard.height - lblCloseDashboard.Top
    End With
    
    With framePageAddRecord
        .width = frameMain.width
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.9)
        .Left = 0
        .Top = 0
        .height = mainWindowHeight - lblCloseDashboard.height - lblCloseDashboard.Top
    End With
    
    
    multiPageApp.value = 0
            
        
'    Set var_viewClaimsTab = New ClsViewClaims
'    Set var_viewClaimsTab.setForm = Me
'    var_viewClaimsTab.initialize_viewClaimsTab
    
    
Fail:
    MsgBox "The application could not start:" & vbCrLf & vbCrLf & Err.Description, _
           vbCritical, "Startup Error"
           
           
End Sub

Sub setPage(pageNumber As Integer)
    
    Me.multiPageApp.value = pageNumber
    
End Sub

'=======================================

' ---------------------------------------------------------------------
' The page registry. To add a tab:
'   1. Add a Page to MultiPage1 in the designer, with its controls
'   2. Write a clsPageXxx class that Implements IPage
'   3. Add ONE AddPage line here, in the same order as the tabs
' Nothing else in the app changes.
' ---------------------------------------------------------------------
Private Sub RegisterPages()
    
    Set m_pages = New Collection
    
    AddPage New ClsPageHome, multiPageApp.Pages("pageHome")
    AddPage New clsPageAddRecord, multiPageApp.Pages("pageAddRecord")
    AddPage New ClsPageViewRecords, multiPageApp.Pages("pageViewRecords")
    AddPage New ClsPageLogCall, multiPageApp.Pages("pageLogCall")
    AddPage New clsPageAdmin, multiPageApp.Pages("pageAdmin")
    AddPage New ClsPageSearch, multiPageApp.Pages("pageSearch")
    'AddPage New clsPageAdmin, MultiPage1.Pages("pgAdmin")
    
End Sub
 
Private Sub AddPage(ctrl As IPage, pg As MSForms.Page)
    
    On Error GoTo Fail
    ctrl.InitPage m_ctx, pg
    pg.caption = ctrl.pageTitle
    m_pages.Add ctrl
    Exit Sub
 
Fail:
    ' Name the failing tab rather than letting a generic control error
    ' bubble up from a form with 150 controls on it.
    Err.Raise Err.Number, "AddPage", _
        "Tab '" & ctrl.pageTitle & "' failed to initialise:" & vbCrLf & Err.Description
End Sub
 
' Tab access rules.
'
' UNREGISTERED user -> Home only. Every other tab is disabled, which
' greys the caption and blocks selection, so the reason is visible
' rather than the tabs just vanishing.
'
' REGISTERED non-admin -> everything except Admin, which is hidden.
'
' As always, this is convenience and not security: anyone who can open
' the VBE can re-enable a tab. The real enforcement is that every write
' in clsClaimRepository re-checks the role server-side.

Private Sub ApplyTabVisibility()
    
    Dim i As Long
    Dim registered As Boolean
 
    On Error Resume Next
    registered = m_ctx.IsRegistered
         
    'Admin will be visible only if Admin is logged in.
    lblMenuItemAdmin.Visible = (registered And m_ctx.IsAdmin)
    'multiPageApp.Pages("pgAdmin").Visible = (registered And m_ctx.IsAdmin)
     
'    For i = 0 To multiPageApp.Pages.Count - 1
'        If multiPageApp.Pages(i).Name = "pageHome" Then
'            multiPageApp.Pages(i).enabled = True
'
'        Else
'            multiPageApp.Pages(i).enabled = registered
'            If Not registered Then
'                ' Say why on the tab itself - a greyed tab with no
'                ' explanation just looks broken.
'                multiPageApp.Pages(i).caption = multiPageApp.Pages(i).caption & " (locked)"
'            End If
'        End If
'    Next i
           
    'Hide the navigation buttons if not registered
    If Not registered Then
        lblMenuItemAddClaim.Visible = False
        lblMenuItemAdmin.Visible = False
        lblMenuItemLogCall.Visible = False
        lblMenuItemSearch.Visible = False
        lblMenuItemViewClaims.Visible = False
    End If
        
    On Error GoTo 0
End Sub

' Home is the landing page. Falls back to the first usable tab if Home
' is somehow missing.

Private Function FirstVisiblePageIndex() As Long
    Dim i As Long
 
    For i = 0 To multiPageApp.Pages.Count - 1
        If multiPageApp.Pages(i).Name = "pgHome" Then
            FirstVisiblePageIndex = i
            Exit Function
        End If
    Next i
 
    For i = 0 To multiPageApp.Pages.Count - 1
        If multiPageApp.Pages(i).Visible And multiPageApp.Pages(i).enabled Then
            FirstVisiblePageIndex = i
            Exit Function
        End If
    Next i
End Function
  
  
' =====================================================================
' Tab switching
' =====================================================================
'Private Sub MultiPage1_Change()
Private Sub multiPageApp_Change()
    
    If Not m_initialised Then Exit Sub
    ActivateCurrentPage

End Sub
 
' Each controller refreshes itself here, on the tab the user actually
' opened. Nothing reloads speculatively - and because the claim cache is
' shared in clsAppContext, switching tabs doesn't re-read the shared file
' unless something invalidated it.
Private Sub ActivateCurrentPage()
    
    Dim idx As Long
    Dim ctrl As IPage
 
    idx = multiPageApp.value
    If idx < 0 Then Exit Sub
    If idx + 1 > m_pages.Count Then Exit Sub
 
    On Error GoTo Fail
    Set ctrl = m_pages(idx + 1)      ' MultiPage 0-based, Collection 1-based
    ctrl.OnActivate
    Exit Sub
 
Fail:
    ShowStatus "Could not load this tab: " & Err.Description
End Sub
 
' =====================================================================
' Status bar - clsAppContext calls this through its status sink, so
' controllers post messages without knowing a form exists.
' =====================================================================
Public Sub ShowStatus(ByVal msg As String)
    
'    On Error Resume Next
'    lblStatusBar.caption = msg
'    DoEvents
'    On Error GoTo 0

    On Error Resume Next
    ' Pages ask the shell to switch tabs by posting a GOTO: message,
    ' rather than holding a reference to the form and driving it
    ' directly. Keeps controllers decoupled from the shell.
    If Left$(msg, 5) = "GOTO:" Then
        GoToPage "pg" & Mid$(msg, 6)
        Exit Sub
    End If
 
    lblStatusBar.caption = msg
    DoEvents
    On Error GoTo 0
    
End Sub
 
' =====================================================================
' Closing
' =====================================================================
Private Sub cmdClose_Click()
    
    Unload Me

End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim ctrl As IPage
    Dim srch As ClsPageSearch
 
    ' The Search tab is the only one holding unsaved edits, so ask before
    ' throwing them away.
    On Error Resume Next
    For Each ctrl In m_pages
        If TypeOf ctrl Is ClsPageSearch Then
            Set srch = ctrl
            If Not srch.ConfirmDiscardIfDirty() Then
                Cancel = True
                Exit Sub
            End If
        End If
    Next ctrl
    On Error GoTo 0

End Sub
 
Private Sub UserForm_Terminate()

    Set m_pages = Nothing
    Set m_ctx = Nothing
    
End Sub
 

' =====================================================================
' THE SHELL DESIGN - size of the window and where the three top-level
' controls sit. Everything else lives inside a tab and is arranged by
' that tab's own controller.
' =====================================================================
Private Sub ArrangeShell()
    
    Dim registered As Boolean
 
    On Error Resume Next
    registered = m_ctx.IsRegistered
    
    
    Me.caption = AppUtil.getAppName
    
    Me.StartUpPosition = 0
    Me.width = Application.UsableWidth * 0.98
    Me.height = Application.height * 0.98
    
    
    mainWindowHeight = Me.height - 30
    mainWindowWidth = Me.width - 10
    
    With frameDashboard
        .width = 170
        .Left = 0
        .Top = 0
        .height = mainWindowHeight
        .ZOrder (1)
        .BackColor = RGB(237, 237, 237)
        .BorderColor = vbWhite
        
    End With
    
    With lblAppName
        .caption = "Menu Item"
        .foreColor = AppUtil.getThemeColor()
    End With
    
    With frameMain
        '.Picture = LoadPicture(ThisWorkbook.Path & "\panelPrinciple.jpg")
        .PictureSizeMode = fmPictureSizeModeStretch
        .height = mainWindowHeight
        .width = mainWindowWidth - Me.frameDashboard.width
        .Left = Me.frameDashboard.width
        .Top = 0
        .ZOrder (1)
        .BackColor = RGB(237, 237, 237)
        
    End With
            
    lblStatusBar.Top = frameMain.height - 20
    lblStatusBar.Left = 9
    lblStatusBar.width = frameMain.width - 2 * lblStatusBar.Left
    lblStatusBar.height = 20
    lblStatusBar.BackColor = RGB(237, 237, 237)
    lblStatusBar.foreColor = AppUtil.getThemeColor()
    
    With lblHome
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
        .foreColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    End With

    With lblMenuItemAddClaim
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
        .foreColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    End With

    With lblMenuItemViewClaims
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
        .foreColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    
    End With

    With lblMenuItemLogCall
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
        .foreColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    
    End With

    With lblMenuItemSearch
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
        .foreColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    
    End With

'    With lblAbout
'        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
'        .ForeColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
'
'    End With
        
    With lblAppHeader
        .width = frameMain.width
        .Left = 10
        .Top = 5
        .height = 50
        .BackColor = RGB(237, 237, 237)
        .foreColor = AppUtil.getThemeColor()
        .caption = AppUtil.getAppName()
    End With
    
    With lblAppHeaderLine
        .width = frameMain.width
        .Left = 0
        .Top = lblAppHeader.Top + lblAppHeader.height
        .height = 1.5
        .BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    End With
    
    With multiPageApp
        .Style = fmTabStyleNone
        .width = frameMain.width
        .Left = 0
        .Top = lblAppHeaderLine.Top + lblAppHeaderLine.height
        .height = mainWindowHeight - .Top - 20
        .BackColor = vbYellow
        
    End With
            
    collection_navigationButton.Add Styler.getStyledButtonNavigationBar(lblHome)
    collection_navigationButton.Add Styler.getStyledButtonNavigationBar(lblMenuItemAddClaim)
'    collection_navigationButton.Add Styler.getStyledButtonNavigationBar(lblAbout)
    collection_navigationButton.Add Styler.getStyledButtonNavigationBar(lblMenuItemViewClaims)
    collection_navigationButton.Add Styler.getStyledButtonNavigationBar(lblMenuItemLogCall)
    collection_navigationButton.Add Styler.getStyledButtonNavigationBar(lblMenuItemAdmin)
    collection_navigationButton.Add Styler.getStyledButtonNavigationBar(lblMenuItemSearch)
           
    registered = m_ctx.IsRegistered
    
    lblUserProfile.Left = frameMain.width - lblUserProfile.width
    lblUserProfile.Top = 0
    lblUserProfile.foreColor = AppUtil.getThemeColor()
        
    If Not registered Then
        lblUserProfile.caption = "    " & m_ctx.userName & " (Not Registered)"
    Else
        lblUserProfile.caption = "    " & m_ctx.userName & IIf(m_ctx.IsAdmin, "  (Admin)", " (User)")
    End If
    
       
        
End Sub
 

