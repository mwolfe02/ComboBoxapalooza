Option Compare Database
Option Explicit

Sub PopulateLargeTable()
    
End Sub

Function HasEleven(txt As String) As Boolean
    HasEleven = (txt Like "*eleven*")
End Function

Sub Throw(Msg As String)
    'Dummy implementation
    'See: https://nolongerset.com/throwing-errors-in-vba/
    Debug.Print Msg
End Sub