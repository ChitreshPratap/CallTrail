VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TsTextFieldStyle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim var_mukhTextBox As MSForms.TextBox
Dim var_lblCaption As MSForms.Label
Dim var_lblLine As MSForms.Label
Public WithEvents txtDynamic As MSForms.TextBox
Attribute txtDynamic.VB_VarHelpID = -1
Dim txtDynamic2 As MSForms.control
Attribute txtDynamic2.VB_VarHelpID = -1

Sub Init(mukhTextBox As control, Optional placeholder As String = "Enter Value")
    
    Set var_mukhTextBox = mukhTextBox
    Set var_lblCaption = mukhTextBox.Parent.Controls.Add("Forms.Label.1", "lblKey" & mukhTextBox.Name)
    Set var_lblLine = mukhTextBox.Parent.Controls.Add("Forms.Label.1", "lblLine" & mukhTextBox.Name)
    var_lblCaption.caption = placeholder
    var_mukhTextBox.BackStyle = fmBackStyleTransparent
    var_mukhTextBox.SpecialEffect = fmSpecialEffectFlat
    
    
    var_lblCaption.Font = var_mukhTextBox.Font
    var_lblCaption.Font.Size = var_mukhTextBox.Font.Size
    var_lblCaption.Font.Name = var_mukhTextBox.Font.Name
    Set txtDynamic = mukhTextBox
    Set txtDynamic2 = mukhTextBox
    

'    Set var_lblCaption = lblCaption
'    Set var_lblLine = lblLine
    
    
    var_lblCaption.Top = var_mukhTextBox.Top + 1.5
    var_lblCaption.height = var_mukhTextBox.height - 1.5
    var_lblCaption.width = var_mukhTextBox.width - 7.5
    var_lblCaption.Left = var_mukhTextBox.Left + 7.5
    
    var_lblLine.Top = var_mukhTextBox.Top + var_mukhTextBox.height - 1
    var_lblLine.Left = var_lblCaption.Left - 2
    var_lblLine.width = var_lblCaption.width + 2
    var_lblLine.height = 1.3
    var_lblLine.BackColor = VarnahUtil.getFadeColor(var_mukhTextBox.ForeColor, -0.3)
    
    var_lblCaption.ForeColor = VarnahUtil.getFadeColor(var_mukhTextBox.ForeColor, 0.55)
    var_lblCaption.ZOrder (1)

End Sub

Sub styleWithText()
                    
    var_lblCaption.Font.Size = var_mukhTextBox.Font.Size - 1
    var_lblCaption.height = var_lblCaption.Font.Size + 6
    var_lblCaption.Font.Bold = True
    var_lblCaption.ForeColor = VarnahUtil.getFadeColor(var_mukhTextBox.ForeColor, -0.3)
    
    var_lblCaption.Top = var_mukhTextBox.Top - var_lblCaption.height
'    var_lblCaption.Font.Size = var_mukhTextBox.Font.Size - 2
    var_lblCaption.Left = var_mukhTextBox.Left + 6.5
    
                    
End Sub

Sub styleWithoutText()
    
    var_lblCaption.Font = var_mukhTextBox.Font
    var_lblCaption.Font.Size = var_mukhTextBox.Font.Size
    var_lblCaption.Font.Name = var_mukhTextBox.Font.Name
    
    var_lblCaption.Top = var_mukhTextBox.Top + 1.5
    var_lblCaption.height = var_mukhTextBox.height - 1.5
    var_lblCaption.width = var_mukhTextBox.width - 7.5
    var_lblCaption.Left = var_mukhTextBox.Left + 7.5
    
    var_lblCaption.ForeColor = VarnahUtil.getFadeColor(var_mukhTextBox.ForeColor, 0.55)
    var_lblCaption.ZOrder (1)
    
    
End Sub

Sub styleOnTextChanged()
    
    If var_mukhTextBox.Text = "" Then
        Me.styleWithoutText
    Else
        Me.styleWithText
    End If

End Sub

Private Sub Class_Initialize()

End Sub

Private Sub txtDynamic_Change()
    Me.styleOnTextChanged
End Sub

Private Sub txtDynamic2_GotFocus()
    MsgBox "Hello"
End Sub

Private Sub txtDynamic2_Enter()
    MsgBox "Hello"
End Sub
