Option Compare Database
Option Explicit

'Add a call to this function to the OnKeyUp and OnMouseUp events for all combo boxes: 'vv
' =EnterCombo([Form].[ActiveControl])   for combo boxes with 1 column only
' =EnterCombo([Form].[ActiveControl],1) for combo boxes with 2 columns (1st hidden) ^^
Function EnterCombo(cmb As ComboBox, Optional col As Integer = 0) 'vv
    On Error GoTo Err_EnterCombo
    
    If Nz(cmb.Column(col)) = "" Then
        cmb.Dropdown
    Else
        'we need to check the current selection length, otherwise the control does not work as
        'expected when the user clicks directly on the triangular dropdown symbol
        If cmb.SelLength = 0 And cmb.SelStart > 0 And Nz(cmb.Column(col)) <> "" Then
            cmb.SelStart = 0
            cmb.SelLength = Len(Nz(cmb.Column(col)))
        End If
    End If
        
Exit_EnterCombo:
    Exit Function
Err_EnterCombo:
    'if there is no record selected an error will be generated
    '(happens if control is in form header/footer and AllowAdditions is set to false
    ' and the form is filtered so there are no records)
    'there is no need to fill up the error log with these errors, so we won't record it
    If Err.Number <> 2185 Then
        'Replace with your own error logging function:
        'LogError Err.Number, Err.Description, "EnterCombo", cmb.Name, False
    End If
    Resume Exit_EnterCombo
End Function '^^