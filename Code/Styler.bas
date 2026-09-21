Attribute VB_Name = "Styler"


Function getStyledBox(txtBox As Msforms.TextBox, Optional placeHolderText As String = "Enter Value") As TextBoxStyle1

    Dim styl As TextBoxStyle1
    Set styl = New TextBoxStyle1
    styl.Init txtBox, placeholder:=placeHolderText
    Set getStyledBox = styl
    
End Function

Function getStyleButton(btn As Msforms.Label) As tsLabelHE
        
    Dim btnStyle As tsLabelHE
    Set btnStyle = New tsLabelHE
    btn.BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.2)
    btn.foreColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.8)
    btnStyle.setMukhLabel btn
    btnStyle.setIncDecInFont 2
    btnStyle.setOnHoverBackColor VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.5)
    btnStyle.setOnHoverForeColor VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
    Set getStyleButton = btnStyle

End Function

Function getStyledBoxULine(txtBox As Msforms.TextBox, Optional placeHolderText As String = "Enter Value") As TsTextFieldStyle2
    
    Dim styledBox As TsTextFieldStyle2
    Set styledBox = New TsTextFieldStyle2
    styledBox.Init txtBox, placeHolderText
    Set getStyledBoxULine = styledBox
    
End Function


Function getStyledButtonNavigationBar(lbl As Msforms.Label) As tsLabelHE
    
    Dim btnStyle As tsLabelHE
    Set btnStyle = New tsLabelHE
    lbl.BackColor = RGB(219, 238, 242)
    lbl.foreColor = vbBlack
    btnStyle.setMukhLabel lbl
    btnStyle.setIncDecInFont 4
    btnStyle.setOnHoverBackColor VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.2)
    'btnStyle.setOnHoverForeColor VarnahUtil.getFadeColor(AppUtil.getThemeColor(), -0.4)
    btnStyle.setOnHoverForeColor vbWhite
    Set getStyledButtonNavigationBar = btnStyle
    
End Function

Sub stylePageFrame(pageFrame As Msforms.frame)
    
    pageFrame.BackColor = RGB(255, 255, 255)

End Sub

Sub getStyledHeading(lblHeading As Msforms.Label)
    
    lblHeading.foreColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.2)
        
End Sub

Sub getStyledHeadingULine(lblHeadingULine As Msforms.Label)
    lblHeadingULine.BackColor = VarnahUtil.getFadeColor(AppUtil.getThemeColor(), 0.2)
    
End Sub

