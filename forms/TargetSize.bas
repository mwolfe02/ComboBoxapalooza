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
    ItemSuffix =21
    Right =14235
    Bottom =12465
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x6aca8af815a8e540
    End
    Caption ="Target Size"
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
            Height =9810
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
                    Top =1890
                    Width =8460
                    Height =2880
                    BackColor =14602694
                    BorderColor =10921638
                    Name ="Box8"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =1890
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =4770
                    BackThemeColorIndex =-1
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =2880
                    Top =990
                    Width =1800
                    Height =390
                    FontSize =14
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="Combo0"
                    RowSourceType ="Value List"
                    RowSource ="Small;Medium;Large;Extra Large"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =990
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =1380
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1080
                            Top =990
                            Width =1845
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label1"
                            Caption ="Attached Label:"
                            GridlineColor =10921638
                            LayoutCachedLeft =1080
                            LayoutCachedTop =990
                            LayoutCachedWidth =2925
                            LayoutCachedHeight =1380
                            ForeTint =85.0
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =270
                    Width =1800
                    Height =390
                    FontSize =14
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="Combo2"
                    RowSourceType ="Value List"
                    RowSource ="Small;Medium;Large;Extra Large"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =270
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =660
                End
                Begin Label
                    OverlapFlags =85
                    Left =630
                    Top =270
                    Width =2145
                    Height =390
                    FontSize =14
                    BorderColor =8355711
                    ForeColor =2500134
                    Name ="Label3"
                    Caption ="Unattached Label:"
                    GridlineColor =10921638
                    LayoutCachedLeft =630
                    LayoutCachedTop =270
                    LayoutCachedWidth =2775
                    LayoutCachedHeight =660
                    ForeTint =85.0
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3150
                    Top =2160
                    Width =1800
                    Height =390
                    FontSize =14
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="cbSingleColumn"
                    RowSourceType ="Value List"
                    RowSource ="Small;Medium;Large;Extra Large"
                    OnKeyUp ="=EnterCombo([Form].[ActiveControl])"
                    OnMouseUp ="=EnterCombo([Form].[ActiveControl])"
                    GridlineColor =10921638

                    LayoutCachedLeft =3150
                    LayoutCachedTop =2160
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =2550
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =360
                            Top =2160
                            Width =2625
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label5"
                            Caption ="Single Column:"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2160
                            LayoutCachedWidth =2985
                            LayoutCachedHeight =2550
                            ForeTint =85.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3150
                    Top =3060
                    Width =1800
                    Height =390
                    FontSize =14
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="cbHiddenBoundColumn"
                    RowSourceType ="Value List"
                    RowSource ="S;Small;M;Medium;L;Large;XL;Extra Large"
                    ColumnWidths ="0;144"
                    OnKeyUp ="=EnterCombo([Form].[ActiveControl])"
                    OnMouseUp ="=EnterCombo([Form].[ActiveControl])"
                    GridlineColor =10921638

                    LayoutCachedLeft =3150
                    LayoutCachedTop =3060
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =3450
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =360
                            Top =3060
                            Width =2730
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label7"
                            Caption ="Hidden Bound Column:"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3060
                            LayoutCachedWidth =3090
                            LayoutCachedHeight =3450
                            ForeTint =85.0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    Left =540
                    Top =1620
                    Width =2625
                    Height =390
                    FontSize =14
                    BackColor =14602694
                    BorderColor =8355711
                    ForeColor =2500134
                    Name ="Label9"
                    Caption ="EnterCombo Enabled"
                    GridlineColor =10921638
                    LayoutCachedLeft =540
                    LayoutCachedTop =1620
                    LayoutCachedWidth =3165
                    LayoutCachedHeight =2010
                    BackThemeColorIndex =-1
                    ForeTint =85.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5310
                    Top =2250
                    Width =2160
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text10"
                    ControlSource ="=[cbSingleColumn]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5310
                    LayoutCachedTop =2250
                    LayoutCachedWidth =7470
                    LayoutCachedHeight =2565
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =5310
                    Top =1890
                    Width =1680
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label12"
                    Caption ="Combo box value"
                    GridlineColor =10921638
                    LayoutCachedLeft =5310
                    LayoutCachedTop =1890
                    LayoutCachedWidth =6990
                    LayoutCachedHeight =2205
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5310
                    Top =3060
                    Width =2520
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text13"
                    ControlSource ="=[cbHiddenBoundColumn]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5310
                    LayoutCachedTop =3060
                    LayoutCachedWidth =7830
                    LayoutCachedHeight =3375
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3150
                    Top =3870
                    Width =1800
                    Height =390
                    FontSize =14
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="cbHiddenBoundWorking"
                    RowSourceType ="Value List"
                    RowSource ="S;Small;M;Medium;L;Large;XL;Extra Large"
                    ColumnWidths ="0;144"
                    OnKeyUp ="=EnterCombo([Form].[ActiveControl],1)"
                    OnMouseUp ="=EnterCombo([Form].[ActiveControl],1)"
                    GridlineColor =10921638

                    LayoutCachedLeft =3150
                    LayoutCachedTop =3870
                    LayoutCachedWidth =4950
                    LayoutCachedHeight =4260
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =360
                            Top =3870
                            Width =2730
                            Height =735
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label15"
                            Caption ="Hidden Bound Column:\015\012  (col Arg set to 1)"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3870
                            LayoutCachedWidth =3090
                            LayoutCachedHeight =4605
                            ForeTint =85.0
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5310
                    Top =3870
                    Width =2520
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text16"
                    ControlSource ="=[cbHiddenBoundWorking]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5310
                    LayoutCachedTop =3870
                    LayoutCachedWidth =7830
                    LayoutCachedHeight =4185
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3150
                    Top =5130
                    Width =2790
                    Height =390
                    FontSize =14
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cbExample"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT US_State FROM US_State ORDER BY US_State; "
                    OnKeyUp ="=EnterCombo([cbExample])"
                    OnMouseUp ="=EnterCombo([cbExample])"
                    GridlineColor =10921638

                    LayoutCachedLeft =3150
                    LayoutCachedTop =5130
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =5520
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1350
                            Top =5130
                            Width =1845
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label18"
                            Caption ="Choose State:"
                            GridlineColor =10921638
                            LayoutCachedLeft =1350
                            LayoutCachedTop =5130
                            LayoutCachedWidth =3195
                            LayoutCachedHeight =5520
                            ForeTint =85.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3150
                    Top =6120
                    Width =2790
                    Height =390
                    FontSize =14
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="Combo19"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT US_State FROM US_State ORDER BY US_State; "
                    OnKeyUp ="=EnterCombo([cbExample])"
                    OnMouseUp ="=EnterCombo([cbExample])"
                    GridlineColor =10921638

                    LayoutCachedLeft =3150
                    LayoutCachedTop =6120
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =6510
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1350
                            Top =6120
                            Width =1845
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label20"
                            Caption ="Choose State:"
                            GridlineColor =10921638
                            LayoutCachedLeft =1350
                            LayoutCachedTop =6120
                            LayoutCachedWidth =3195
                            LayoutCachedHeight =6510
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

Private Sub cbExample_KeyUp(KeyCode As Integer, Shift As Integer)
    EnterCombo Me.cbExample
End Sub

Private Sub cbExample_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EnterCombo Me.cbExample
End Sub
