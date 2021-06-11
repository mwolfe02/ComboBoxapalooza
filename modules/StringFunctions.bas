Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : RegExReplace
' Author    : Mike Wolfe <mike@nolongerset.com>
' Date      : 11/4/2010
' Purpose   : Attempts to replace text in the TextToSearch with text and back references
'               from the ReplacePattern for any matches found using SearchPattern.
' Notes     - If no matches are found, TextToSearch is returned unaltered.  To get
'               specific info from a string, use RegExExtract instead.
'>>> RegExReplace("(.*)(\d{3})[\)\s.-](\d{3})[\s.-](\d{4})(.*)", "My phone # is 570.555.1234.", "$1($2)$3-$4$5")
'My phone # is (570)555-1234.
'---------------------------------------------------------------------------------------
'
Function RegExReplace(SearchPattern As String, TextToSearch As String, ReplacePattern As String, _
                      Optional GlobalReplace As Boolean = True, _
                      Optional IgnoreCase As Boolean = False, _
                      Optional MultiLine As Boolean = False) As String
Dim RE As Object

    Set RE = CreateObject("vbscript.regexp")
    With RE
        .MultiLine = MultiLine
        .Global = GlobalReplace
        .IgnoreCase = IgnoreCase
        .Pattern = SearchPattern
    End With
    
    RegExReplace = RE.Replace(TextToSearch, ReplacePattern)
End Function