Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7320
    DatasheetFontHeight =11
    ItemSuffix =4
    Right =14700
    Bottom =17640
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x809cf55f5ea8e540
    End
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
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
            Height =5940
            BackColor =16247774
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =8
            BackTint =20.0
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =240
                    Height =315
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="Combo0"
                    RowSourceType ="Table/Query"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =240
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =555
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =240
                            Width =1875
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label1"
                            Caption ="Unassociated label:"
                            GridlineColor =10921638
                            LayoutCachedLeft =660
                            LayoutCachedTop =240
                            LayoutCachedWidth =2535
                            LayoutCachedHeight =555
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =720
                    Top =2760
                    Width =1800
                    Height =390
                    FontSize =14
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="cbSingleColumn"
                    RowSourceType ="Value List"
                    RowSource ="Extra Small;Small;Medium;Large;Extra Large"
                    GridlineColor =10921638

                    LayoutCachedLeft =720
                    LayoutCachedTop =2760
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =3150
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =720
                            Top =2280
                            Width =1050
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label5"
                            Caption ="Original:"
                            GridlineColor =10921638
                            LayoutCachedLeft =720
                            LayoutCachedTop =2280
                            LayoutCachedWidth =1770
                            LayoutCachedHeight =2670
                            ForeTint =85.0
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3120
                    Top =2760
                    Width =1800
                    Height =390
                    FontSize =14
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =3484194
                    Name ="Combo2"
                    RowSourceType ="Value List"
                    RowSource ="Extra Small;Small;Medium;Large;Extra Large"
                    OnKeyUp ="=EnterCombo([Form].[ActiveControl])"
                    OnMouseUp ="=EnterCombo([Form].[ActiveControl])"
                    GridlineColor =10921638

                    LayoutCachedLeft =3120
                    LayoutCachedTop =2760
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =3150
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3120
                            Top =2280
                            Width =1275
                            Height =390
                            FontSize =14
                            BorderColor =8355711
                            ForeColor =2500134
                            Name ="Label3"
                            Caption ="Enhanced:"
                            GridlineColor =10921638
                            LayoutCachedLeft =3120
                            LayoutCachedTop =2280
                            LayoutCachedWidth =4395
                            LayoutCachedHeight =2670
                            ForeTint =85.0
                        End
                    End
                End
            End
        End
    End
End
