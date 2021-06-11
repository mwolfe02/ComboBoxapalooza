Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =16
    GridY =16
    Width =8884
    DatasheetFontHeight =11
    ItemSuffix =23
    Right =14040
    Bottom =12240
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x6aca8af815a8e540
    End
    Caption ="Target Size"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =9810
            BackColor =16247774
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =8
            BackTint =20.0
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3240
                    Top =1170
                    Width =2790
                    Height =390
                    FontSize =14
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cbProgressiveFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT US_State FROM US_State WHERE US_State Like '**' ORDER BY US_State; "
                    OnKeyUp ="[Event Procedure]"
                    OnMouseUp ="=EnterCombo([cbProgressiveFilter])"
                    OnChange ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =1170
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1440
                            Top =1170
                            Width =1845
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label18"
                            Caption ="Choose State:"
                            GridlineColor =10921638
                            LayoutCachedLeft =1440
                            LayoutCachedTop =1170
                            LayoutCachedWidth =3285
                            LayoutCachedHeight =1560
                            ForeTint =85.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3240
                    Top =3780
                    Width =2790
                    Height =390
                    FontSize =14
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cbStateLookup"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT US_State FROM US_State WHERE US_State Like '**' ORDER BY US_State; "
                    OnKeyUp ="[Event Procedure]"
                    OnMouseUp ="=EnterCombo([cbProgressiveFilter])"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =3780
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =4170
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1440
                            Top =3780
                            Width =1845
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label22"
                            Caption ="Choose State:"
                            GridlineColor =10921638
                            LayoutCachedLeft =1440
                            LayoutCachedTop =3780
                            LayoutCachedWidth =3285
                            LayoutCachedHeight =4170
                            ForeTint =85.0
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private StateLookup As weComboLookup

Private Sub Form_Load()
    Set StateLookup = New weComboLookup
    StateLookup.Initialize Me.cbExample
End Sub
Private Sub Form_Close()
    Set StateLookup = Nothing
End Sub

Private Sub cbProgressiveFilter_Change()
    Dim FilterTxt As String
    FilterTxt = Me.cbExample.Text
    'FilterTxt = ExtractFilterTxt(Me.cbExample)
    
    SetRowSrc BuildRowSrc(FilterTxt)
End Sub

Private Sub SetRowSrc(SqlString As String)
    'Don't update the RowSource if there is no change
    '   as this will force an unnecessary requery
    If Me.cbExample.RowSource = SqlString Then Exit Sub
    
    Me.cbExample.RowSource = SqlString
End Sub

Private Function BuildRowSrc(UserText As String) As String
    Dim s As String
    s = "SELECT US_State FROM US_State "
    If Len(UserText) > 0 Then s = s & "WHERE US_State Like '*" & UserText & "*' "
    s = s & "ORDER BY US_State;"
    
    BuildRowSrc = s
End Function

Private Sub cbStateLookup_KeyUp(KeyCode As Integer, Shift As Integer)
    EnterCombo Me.cbStateLookup
End Sub


'Once a
Private Function ExtractFilterTxt(Combo As ComboBox) As String

End Function
