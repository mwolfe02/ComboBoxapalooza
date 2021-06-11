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
    Width =8884
    DatasheetFontHeight =11
    ItemSuffix =7
    Right =20040
    Bottom =29010
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x65f0803196a8e540
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
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
            Height =7560
            BackColor =15854048
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =10080
                    Left =1920
                    Top =2040
                    Width =1200
                    Height =315
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cbOriginal"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Airport.Rank, Airport.IATA_Code, Airport.Airport, Airport.City, Airport.S"
                        "tate FROM Airport WHERE (((Airport.IATA_Code) Like '**')) OR (((Airport.Airport)"
                        " Like '**')) OR (((Airport.City) Like '**')) OR (((Airport.State) Like '**')) OR"
                        "DER BY Airport.IATA_Code; "
                    ColumnWidths ="0;1008;5040;3312;576"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =2040
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =2355
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =2040
                            Width =840
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label1"
                            Caption ="Airport:"
                            GridlineColor =10921638
                            LayoutCachedLeft =660
                            LayoutCachedTop =2040
                            LayoutCachedWidth =1500
                            LayoutCachedHeight =2355
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =660
                    Top =1380
                    Width =810
                    Height =315
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label3"
                    Caption ="Original"
                    GridlineColor =10921638
                    LayoutCachedLeft =660
                    LayoutCachedTop =1380
                    LayoutCachedWidth =1470
                    LayoutCachedHeight =1695
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListWidth =10080
                    Left =1920
                    Top =3600
                    Width =1200
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =3484194
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"510\""
                    Name ="cbAirport"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Airport.Rank, Airport.IATA_Code, Airport.Airport, Airport.City, Airport.S"
                        "tate FROM Airport WHERE (((Airport.IATA_Code) Like '**')) OR (((Airport.Airport)"
                        " Like '**')) OR (((Airport.City) Like '**')) OR (((Airport.State) Like '**')) OR"
                        "DER BY Airport.IATA_Code; "
                    ColumnWidths ="0;1008;5040;3312;576"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =3600
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =3915
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =660
                            Top =3600
                            Width =840
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label5"
                            Caption ="Airport:"
                            GridlineColor =10921638
                            LayoutCachedLeft =660
                            LayoutCachedTop =3600
                            LayoutCachedWidth =1500
                            LayoutCachedHeight =3915
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =660
                    Top =2940
                    Width =975
                    Height =315
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label6"
                    Caption ="Enhanced"
                    GridlineColor =10921638
                    LayoutCachedLeft =660
                    LayoutCachedTop =2940
                    LayoutCachedWidth =1635
                    LayoutCachedHeight =3255
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

Dim AirportLookup As New weComboLookup

Private Sub Form_Open(Cancel As Integer)
    AirportLookup.Initialize Me.cbAirport
End Sub
