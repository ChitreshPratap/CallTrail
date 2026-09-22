VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
' =====================================================================
' IPage - the contract every tab controller implements.
'
' VBA has no inheritance, but it does have interfaces via Implements.
' This is what makes the tabs pluggable: frmMain holds a Collection of
' IPage and never needs to know which concrete class is behind each tab.
' Adding a sixth tab means writing one new class and registering it -
' no changes to the shell or to the other five.
'
' Lifecycle:
'   InitPage   - called ONCE when the app form opens. Bind controls here.
'   OnActivate - called EVERY time the user switches to this tab.
'                Refresh data here, not in InitPage.
'   PageTitle  - used for the tab caption and error messages.
' =====================================================================
 
Public Sub InitPage(ctx As clsAppContext, pg As MSForms.Page)

End Sub
 
Public Sub OnActivate()

End Sub
 
Public Property Get pageTitle() As String

End Property
 
