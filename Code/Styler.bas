Attribute VB_Name = "Styler"


Function getStyledBox(txtBox As MSForms.TextBox, Optional placeHolderText As String = "Enter Value") As TextBoxStyle1

    Dim styl As TextBoxStyle1
    Set styl = New TextBoxStyle1
    styl.init txtBox, placeholder:=placeHolderText
    Set getStyledBox = styl
    
End Function

Function getStyleButton(btn As MSForms.Label) As MSForms.Label
    Dim styl As tsLabelHE
    Set styl = New tsLabelHE
    'styl.init btn
    Set getStyleButton = btn
End Function

Function getStyledBoxULine(txtBox As MSForms.TextBox, Optional placeHolderText As String = "Enter Value") As TsTextFieldStyle2
    
    Dim styledBox As TsTextFieldStyle2
    Set styledBox = New TsTextFieldStyle2
    styledBox.init txtBox, placeHolderText
    Set getStyledBoxULine = styledBox
    
End Function
