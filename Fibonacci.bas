Function ARC_FIBONACCI(Optional param As Variant) '
On Error GoTo Err
Dim Number
'''''''''''''''''''Parameter Check''''''''''''''''''''
If IsMissing(param) Or TypeName(param) = "String" Then
    GoTo Err
ElseIf TypeName(param) = "Range" Then
    If IsNumeric(param.Value) Then
        Number = param.Value
    Else
        GoTo Err
    End If
Else
    Number = param
End If
'''''''''''''''''''End of Parameter Check''''''''''''''''''''
If Number <= 1 Then  'Base Case
    ARC_FIBONACCI = Number
    GoTo Ext
Else
    ARC_FIBONACCI = ARC_FIBONACCI(Number - 1) + ARC_FIBONACCI(Number - 2)
    GoTo Ext
End If
Err:
    ARC_FIBONACCI = "Missing argument"
Ext:
End Function
