Option Compare Database
Option Explicit

Public Function LogError(lngErrNumber As Long, ByVal strErrMsg As String, strCallingProc As String, _
        Optional bShowUser As Boolean = True, Optional varParam As Variant)
    Select Case lngErrNumber
    Case 2501&
        'Just ignore this error.
    Case 3314&, 2101&, 2115&            'can't save.
        If bShowUser Then
            strErrMsg = "Record cannot be saved at this time." & vbCrLf & _
                "Complete the entry, or press <Esc> to undo."
            MsgBox strErrMsg, vbExclamation, strCallingProc
        End If
    Case Else
        If bShowUser Then
            MsgBox "Error " & lngErrNumber & ": " & strErrMsg, vbExclamation, strCallingProc
        End If
    End Select
End Function