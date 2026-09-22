Attribute VB_Name = "FormUtil"
Option Explicit

Public Function getControlByName(parentContainer As Object, ByVal ctrlName As String, _
                                ByVal isOptional As Boolean) As Object
    
    On Error GoTo NotFound
    Set getControlByName = parentContainer.Controls(ctrlName)
    Exit Function

NotFound:
    If isOptional Then
        Set getControlByName = Nothing
        Exit Function
    End If
    Err.Raise vbObjectError + 21, "FormUtil_", _
        "Control '" & ctrlName & "' was not found on Container :  '" & parentContainer.Name & vbCrLf & vbCrLf & _
        "Check the name in the form designer - it must match exactly."
End Function


