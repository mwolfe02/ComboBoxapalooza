Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : weComboLookup
' Author    : Mike
' Date      : 2/16/2015 - 5/25/2016 10:09
' Purpose   : This class module adds support for filtering combo boxes on multiple fields.
' Usage     : In declarations section of Form module:
'   Private cbFromAcctIDLookup As New weComboLookup
'   Private cbToAcctIDLookup As New weComboLookup
'
'   Private Sub Form_Open(Cancel As Integer)
'       cbFromAcctIDLookup.Initialize Me.cbFromAcctID
'       cbToAcctIDLookup.Initialize Me.cbToAcctID
'   End Sub
'
'           : In RowSource, add one or more clauses like the following to the WHERE clause:
'   MyField LIKE '**'
'
' Notes     - RowSourceType must equal "Table/Query"
'           - RowSource must have one or more "Like '**'" clauses
'           - If RowSource is the name of a query, the query must have 1+ "Like '**'" clauses
'           - The RowSource is checked each time the control is entered, so the class
'               supports runtime updates made to the RowSource (e.g., "cascading" combos)
'           - The RowSource is returned to its unfiltered state:
'               o after it is updated (via the AfterUpdate event)
'               o when the text is cleared via Escape-key-triggered Undo (via KeyUp event)
'               o when the text is manually cleared by the user (via the OnChange event)
'           - According to JETSHOWPLAN, the "LIKE '**'" clause forces a table scan and
'               does not benefit from the "MyField Like '*'" --> "MyField Is Not Null"
'               optimization that a single asterisk Like clause receives; however, it
'               does appear to perform all other indexing tasks first, so hopefully
'               the performance will not be awful; that said, there is a tradeoff between
'               providing an easy to implement feature and maximum performance; YMMV
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private WithEvents Ctl As Access.ComboBox
Attribute Ctl.VB_VarHelpID = -1
Private mOriginalRowSrc As String
Private mFilteredRowSrc As String
Private mMinTextLength As Integer
Private mFilter As String
Private mSearchFilteredRowSrc As String
Private mFilteringEnabled As Boolean

Private mControlName As String
Private mFormName As String

Private Type RowSrcInfo
    HasFilterPlaceHolder As Boolean
    HasContainsFilterInPlace As Boolean
    SupportsContainsFilter As Boolean
    SqlString As String
End Type

'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : Mike
' Date      : 2/16/2015
' Purpose   : Initializes the object instance to provide enhanced filtering.
' Notes     - The RowSource
'---------------------------------------------------------------------------------------
'
Public Sub Initialize(ComboBoxControl As Access.ComboBox, Optional MinTextLength As Integer = 1)
    Set Ctl = ComboBoxControl
    Ctl.OnEnter = SetEventProc(Ctl.OnEnter)
    Ctl.AfterUpdate = SetEventProc(Ctl.AfterUpdate)
    Ctl.OnChange = SetEventProc(Ctl.OnChange)
    Ctl.OnKeyUp = SetEventProc(Ctl.OnKeyUp)
    mMinTextLength = MinTextLength
    
    'we set the following module level variables to assist with debugging errors
    On Error Resume Next
    mControlName = Ctl.NAME: Debug.Assert Len(mControlName) > 0
    mFormName = Ctl.Parent.NAME: Debug.Assert Len(mFormName) > 0
End Sub

Private Property Get UnfilteredRowSrc() As String
    If Not GetRowSrcInfo(mOriginalRowSrc).HasContainsFilterInPlace Then
        UnfilteredRowSrc = mOriginalRowSrc
    ElseIf Not GetRowSrcInfo(Ctl.RowSource).HasContainsFilterInPlace Then
        UnfilteredRowSrc = Ctl.RowSource
    Else
        UnfilteredRowSrc = FilteredRowSrc("")
    End If
End Property

'>>> RegExReplace("('\*)[^*]*(\*')", "AcctNum Like '**'", "$1Mike$2")
' AcctNum Like '*Mike*'
'>>> RegExReplace("('\*)[^*]*(\*')", "AcctNum Like '*John*'", "$1Mike$2")
' AcctNum Like '*Mike*'
'>>> RegExReplace("(""\*)[^*]*(\*"")", "AcctNum Like ""*John*"" OR ""*Tom*""", "$1Mike$2")
' AcctNum Like "*Mike*" OR "*Mike*"
Private Property Get FilteredRowSrc(FilterTxt As String) As String
    Dim rsi As RowSrcInfo
    rsi = GetRowSrcInfo(Ctl.RowSource)
    Debug.Assert rsi.SupportsContainsFilter
    
    Dim SqlBase As String
    SqlBase = rsi.SqlString  'may not match Ctl.RowSource if GetRowSrcInfo calls itself recursively
    
    Dim CleanTxt As String
    CleanTxt = Replace(FilterTxt, "*", "[*]")
    
    Dim EscapedSingleQuotes As String
    EscapedSingleQuotes = Replace(CleanTxt, "'", "''")
    FilteredRowSrc = RegExReplace("('\*)[^*]*(\*')", SqlBase, "$1" & EscapedSingleQuotes & "$2")
    
    Dim EscapedDoubleQuotes As String
    EscapedDoubleQuotes = Replace(CleanTxt, """", """""")
    FilteredRowSrc = RegExReplace("(""\*)[^*]*(\*"")", FilteredRowSrc, "$1" & EscapedDoubleQuotes & "$2")
End Property

'---------------------------------------------------------------------------------------
' Procedure : GetRowSrcInfo
' Author    : Mike
' Date      : 2/17/2015
' Purpose   : Returns info about a RowSrc's "contains" filter support and its current state.
'---------------------------------------------------------------------------------------
'
Private Function GetRowSrcInfo(RowSrc As String) As RowSrcInfo
    Dim rsi As RowSrcInfo
    rsi.SqlString = RowSrc
    rsi.HasFilterPlaceHolder = (InStr(RowSrc, """**""") > 0) Or _
                               (InStr(RowSrc, "'**'") > 0)

    'We will assume that if there is at least one FilterPlaceHolder, then
    ' for our purposes, we assume that there is no
    ' "contains filter" (e.g., "LIKE '*sometext*'") in place
    If Not rsi.HasFilterPlaceHolder Then
        Dim OpenFilterPos As Long, CloseFilterPos As Long
        OpenFilterPos = InStr(RowSrc, """*")
        If OpenFilterPos > 0 Then
            CloseFilterPos = InStr(RowSrc, "*""")
        End If
        rsi.HasContainsFilterInPlace = (OpenFilterPos > 0 And CloseFilterPos > OpenFilterPos)

        If Not rsi.HasContainsFilterInPlace Then
            OpenFilterPos = InStr(RowSrc, "'*")
            If OpenFilterPos > 0 Then
                CloseFilterPos = InStr(OpenFilterPos, RowSrc, "*'")
            End If
            rsi.HasContainsFilterInPlace = (OpenFilterPos > 0 And CloseFilterPos > OpenFilterPos)
        End If
    End If

    rsi.SupportsContainsFilter = (rsi.HasFilterPlaceHolder Or rsi.HasContainsFilterInPlace)
    If Not rsi.SupportsContainsFilter Then
        'Allow for RowSource's that are the names of query definitions
        On Error Resume Next
        Dim QrySql As String
        QrySql = CurrentDb.QueryDefs(RowSrc).SQL
        On Error GoTo 0
        If Len(QrySql) > 0 Then rsi = GetRowSrcInfo(QrySql)
    End If

    GetRowSrcInfo = rsi
End Function

'---------------------------------------------------------------------------------------
' Procedure : MinTextLength
' Author    : Mike
' Date      : 2/16/2015
' Purpose   : Sets the minimum length for a user-entered string before filtering occurs.
' Notes     - Set to 1 so that filtering begins as soon as the user enters text.
'           - Set to 3 so that filtering does not happen until at least 3 characters are entered.
'           - If performance is slow, the MinTextLength may be increased so that fewer
'               results are returned and filtering does not happen as frequently.
'---------------------------------------------------------------------------------------
'
Public Property Let MinTextLength(Value As Integer)
    mMinTextLength = Value
End Property
Public Property Get MinTextLength() As Integer
    MinTextLength = mMinTextLength
End Property

'---------------------------------------------------------------------------------------
' Procedure : Ctl_AfterUpdate
' Author    : Mike
' Date      : 2/16/2015
' Purpose   : After the user makes a selection, we restore the default Row Source so that
'               all options are available the next time the user edits the control.
'---------------------------------------------------------------------------------------
'
Private Sub Ctl_AfterUpdate()
    If Not mFilteringEnabled Then Exit Sub
    Ctl.RowSource = UnfilteredRowSrc
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Ctl_Change
' Author    : Mike
' Date      : 2/16/2015
' Purpose   : As the user edits the field, the RowSource is dynamically updated to
'               show the matching records.
'---------------------------------------------------------------------------------------
'
Private Sub Ctl_Change()
    If Not mFilteringEnabled Then Exit Sub
    
    Dim SelectionLength As Integer
    SelectionLength = GetSelLength(Ctl)
    
    If SelectionLength > 0 Then
        If Len(Ctl.Text) > 0 Then Ctl.Dropdown
        Exit Sub
    End If
    
    Dim txt As String
    txt = Ctl.Text
    If Len(txt) < Me.MinTextLength Then
        Dim UnfilteredRowSource As String
        UnfilteredRowSource = UnfilteredRowSrc
        If Ctl.RowSource <> UnfilteredRowSource Then Ctl.RowSource = UnfilteredRowSource
        Exit Sub
    End If
    
    Dim SavePos As Integer: SavePos = Ctl.SelStart

    Ctl.RowSource = FilteredRowSrc(Ctl.Text)
    Ctl.Dropdown

    Ctl.SetFocus
    Ctl.SelLength = 0
    Ctl.SelStart = SavePos
End Sub

'Avoids error 2185 "Can't reference prop unless the ctl has the focus"
Private Function GetSelLength(Ctl As Control) As Integer
    On Error Resume Next
    GetSelLength = Ctl.SelLength
End Function

' Developer convenience function to prevent accidentally obliterating custom form/control properties.
Private Function SetEventProc(EventProp As String, Optional OverRidableText As String) As String
    If Len(EventProp) = 0 Or EventProp = "[Event Procedure]" Or EventProp = OverRidableText Then
        SetEventProc = "[Event Procedure]"
    Else
        Throw EventProp & " must be changed to '[Event Procedure]'"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Ctl_Enter
' Author    : Mike
' Date      : 2/17/2015
' Purpose   : The RowSource may have changed due to external processing, so we do another check
'---------------------------------------------------------------------------------------
'
Private Sub Ctl_Enter()
    Debug.Assert Ctl.RowSourceType = "Table/Query"
    mOriginalRowSrc = Ctl.RowSource
    Dim rsi As RowSrcInfo
    rsi = GetRowSrcInfo(Ctl.RowSource)
    If Not rsi.SupportsContainsFilter Then
        mFilteringEnabled = False
        'Alert the developer but not the user
        Debug.Assert False   'Throw "RowSource does not support advanced filtering: {0}", Ctl.RowSource
    Else
        mFilteringEnabled = True
    End If
End Sub

' Handle user clearing a field by pressing the Escape key
Private Sub Ctl_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not mFilteringEnabled Then Exit Sub
    If KeyCode = vbKeyEscape And Len(Ctl.Text) = 0 Then Ctl_Change
End Sub