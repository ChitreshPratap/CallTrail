VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxStyle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This style place a placeholder on a text box which will disappear when the textbox has color
' Place a underline at the bottom of the textbox
' No other configuration is required

Dim var_mukhTextBox As Msforms.TextBox
Dim var_lblCaption As Msforms.Label
Dim var_lblLine As Msforms.Label

Public WithEvents txtDynamic As Msforms.TextBox
Attribute txtDynamic.VB_VarHelpID = -1

Sub Init(mukhTextBox As control, Optional placeholder As String = "Enter Value")
    
    Set var_mukhTextBox = mukhTextBox
    Set var_lblCaption = mukhTextBox.Parent.Controls.Add("Forms.Label.1", "lblKey" & mukhTextBox.Name)
    Set var_lblLine = mukhTextBox.Parent.Controls.Add("Forms.Label.1", "lblLine" & mukhTextBox.Name)
    var_lblCaption.caption = placeholder
    var_lblCaption.Font = var_mukhTextBox.Font
    var_lblCaption.Font.Size = var_mukhTextBox.Font.Size
    var_lblCaption.Font.Name = var_mukhTextBox.Font.Name
    Set txtDynamic = mukhTextBox

'    Set var_lblCaption = lblCaption
'    Set var_lblLine = lblLine
    
    var_mukhTextBox.BackStyle = fmBackStyleTransparent
    var_mukhTextBox.SpecialEffect = fmSpecialEffectFlat
    var_lblCaption.Top = var_mukhTextBox.Top + 1.5
    var_lblCaption.height = var_mukhTextBox.height - 1.5
    var_lblCaption.width = var_mukhTextBox.width - 7.5
    var_lblCaption.Left = var_mukhTextBox.Left + 7.5
    var_lblLine.Top = var_mukhTextBox.Top + var_mukhTextBox.height - 1
    var_lblLine.Left = var_lblCaption.Left - 2
    var_lblLine.width = var_lblCaption.width + 2
    var_lblLine.height = 1.3
    var_lblLine.BackColor = VarnahUtil.getFadeColor(var_mukhTextBox.foreColor, -0.3)
    var_lblCaption.foreColor = VarnahUtil.getFadeColor(var_mukhTextBox.foreColor, 0.55)
    var_lblCaption.ZOrder (1)

End Sub

Sub styleOnTextChanged()
    If var_mukhTextBox.Text = "" Then
        var_lblCaption.Visible = True
    Else
        var_lblCaption.Visible = False
    End If
End Sub

Private Sub txtDynamic_Change()
    Me.styleOnTextChanged
End Sub

Public Sub txtDynamic_Enter()
    MsgBox "New txt dynamic"
End Sub

'Private Sub txtDynamic_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'    MsgBox "Insider Focus removes"
'End Sub


