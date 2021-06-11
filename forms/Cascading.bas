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
    Width =9540
    DatasheetFontHeight =11
    ItemSuffix =32
    Right =14040
    Bottom =12240
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x6aca8af815a8e540
    End
    Caption ="Cascading Combos"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
            Height =6660
            BackColor =16247774
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =8
            BackTint =20.0
            Begin
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =180
                    Top =2700
                    Width =8460
                    Height =2880
                    BackColor =14602694
                    BorderColor =10921638
                    Name ="Box8"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =2700
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =5580
                    BackThemeColorIndex =-1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =6480
                    Left =2880
                    Top =540
                    Width =1800
                    Height =390
                    FontSize =14
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"360\""
                    Name ="cbPlaceAll"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Distinct Min(Place.PlaceID), Place.PlaceName, Place.CountyName, Place.Sta"
                        "teName\015\012FROM Place\015\012GROUP BY PlaceName, CountyName, StateName\015\012"
                        "ORDER BY Place.PlaceName;\015\012"
                    ColumnWidths ="0;2160;2160;2160"
                    OnKeyUp ="=EnterCombo([Form].[ActiveControl])"
                    OnMouseUp ="=EnterCombo([Form].[ActiveControl])"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =540
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =930
                End
                Begin Label
                    OverlapFlags =85
                    Left =630
                    Top =540
                    Width =2145
                    Height =390
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =2500134
                    Name ="Label3"
                    Caption ="Every Place:"
                    GridlineColor =10921638
                    LayoutCachedLeft =630
                    LayoutCachedTop =540
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =930
                    ForeTint =85.0
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    Left =540
                    Top =2430
                    Width =2625
                    Height =390
                    FontSize =14
                    BackColor =14602694
                    BorderColor =8355711
                    ForeColor =2500134
                    Name ="Label9"
                    Caption ="Cascading Combos"
                    GridlineColor =10921638
                    LayoutCachedLeft =540
                    LayoutCachedTop =2430
                    LayoutCachedWidth =3165
                    LayoutCachedHeight =2820
                    BackThemeColorIndex =-1
                    ForeTint =85.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListWidth =6480
                    Left =2880
                    Top =2970
                    Width =1800
                    Height =390
                    FontSize =14
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"10\";\"200\""
                    Name ="cbState"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Place.StateName FROM Place WHERE (((Place.StateName)<>\"\")); "
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyUp ="=EnterCombo([Form].[ActiveControl])"
                    OnMouseUp ="=EnterCombo([Form].[ActiveControl])"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =2970
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =3360
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =540
                            Top =2970
                            Width =2145
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label21"
                            Caption ="State:"
                            GridlineColor =10921638
                            LayoutCachedLeft =540
                            LayoutCachedTop =2970
                            LayoutCachedWidth =2685
                            LayoutCachedHeight =3360
                            ForeTint =85.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListWidth =6480
                    Left =2880
                    Top =3480
                    Width =1800
                    Height =390
                    FontSize =14
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"10\";\"200\""
                    Name ="cbCounty"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Place.StateName FROM Place WHERE (((Place.StateName)<>\"\")); "
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyUp ="=EnterCombo([Form].[ActiveControl])"
                    OnMouseUp ="=EnterCombo([Form].[ActiveControl])"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =3480
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =3870
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =540
                            Top =3480
                            Width =2145
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label23"
                            Caption ="County:"
                            GridlineColor =10921638
                            LayoutCachedLeft =540
                            LayoutCachedTop =3480
                            LayoutCachedWidth =2685
                            LayoutCachedHeight =3870
                            ForeTint =85.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ListWidth =6480
                    Left =2880
                    Top =3990
                    Width =1800
                    Height =390
                    FontSize =14
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"10\";\"200\""
                    Name ="cbPlace"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Place.StateName FROM Place WHERE (((Place.StateName)<>\"\")); "
                    ColumnWidths ="1440"
                    OnKeyUp ="=EnterCombo([Form].[ActiveControl])"
                    OnMouseUp ="=EnterCombo([Form].[ActiveControl])"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =3990
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =4380
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =540
                            Top =3990
                            Width =2145
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label25"
                            Caption ="Place:"
                            GridlineColor =10921638
                            LayoutCachedLeft =540
                            LayoutCachedTop =3990
                            LayoutCachedWidth =2685
                            LayoutCachedHeight =4380
                            ForeTint =85.0
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =990
                    Width =2160
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text10"
                    ControlSource ="=[cbPlaceAll].[Column](2)"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =990
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =1305
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1980
                            Top =990
                            Width =795
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label27"
                            Caption ="County:"
                            GridlineColor =10921638
                            LayoutCachedLeft =1980
                            LayoutCachedTop =990
                            LayoutCachedWidth =2775
                            LayoutCachedHeight =1305
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2880
                    Top =1440
                    Width =2160
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text31"
                    ControlSource ="=[cbPlaceAll].[Column](3)"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =1440
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =1755
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1980
                            Top =1440
                            Width =735
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label28"
                            Caption ="State:"
                            GridlineColor =10921638
                            LayoutCachedLeft =1980
                            LayoutCachedTop =1440
                            LayoutCachedWidth =2715
                            LayoutCachedHeight =1755
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

Private StateLookup As New weComboLookup
Private CountyLookup As New weComboLookup
Private PlaceLookup As New weComboLookup

'The Place database table came from here: http://download.geonames.org/export/zip/
'see: https://stackoverflow.com/a/10484216/154439

Private Sub cbState_AfterUpdate()
    Me.cbCounty.Value = Null
    Me.cbCounty.RowSource = "SELECT DISTINCT CountyName FROM Place WHERE StateName='" & Me.cbState.Value & "'"
    
    Me.cbPlace.Value = Null
    Me.cbPlace.RowSource = vbNullString
End Sub


Private Sub cbCounty_AfterUpdate()
    Me.cbPlace.Value = Null
    Me.cbPlace.RowSource = "SELECT DISTINCT PlaceName FROM Place " & _
                           "WHERE StateName='" & Me.cbState.Value & "'" & _
                           "  AND CountyName='" & Me.cbCounty.Value & "'"
End Sub

Private Sub Form_Open(Cancel As Integer)
'    StateLookup.Initialize Me.cbState
'    CountyLookup.Initialize Me.cbCounty
'    PlaceLookup.Initialize Me.cbPlace
End Sub
