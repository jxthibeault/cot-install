Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =12
    Left =-105
    Top =1215
    RecSrcDt = Begin
        0xa30bc3a59febe540
    End
    RecordSource ="SELECT [qryEquipTypeAccessoriesAndStartup].[strDescription], [qryEquipTypeAccess"
        "oriesAndStartup].[intOptionFor], [qryEquipTypeAccessoriesAndStartup].[ysnInStock"
        "], [qryEquipTypeAccessoriesAndStartup].[ysnReadyForInstall], [qryEquipTypeAccess"
        "oriesAndStartup].[lngID], [qryEquipTypeAccessoriesAndStartup].[intInstall] FROM "
        "qryEquipTypeAccessoriesAndStartup; "
    Caption ="subrptAccessoryEquipment"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    FitToPage =1
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
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
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
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =495
            Name ="secReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =75.0
            Begin
                Begin Label
                    OverlapFlags =12
                    TextAlign =2
                    Left =120
                    Top =240
                    Width =4320
                    Height =240
                    BorderColor =8355711
                    Name ="lblDescription"
                    Caption ="Item Description"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =240
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =480
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =12
                    TextAlign =2
                    Left =4740
                    Top =240
                    Width =2220
                    Height =255
                    BorderColor =8355711
                    Name ="lblForLine"
                    Caption ="For Line"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4740
                    LayoutCachedTop =240
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =495
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =4
                    TextAlign =2
                    Left =3578
                    Width =4185
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    Name ="lblReportHeader"
                    Caption ="Options, Accessories, Startup Supplies"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3578
                    LayoutCachedWidth =7763
                    LayoutCachedHeight =315
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =4
                    TextAlign =2
                    Left =10020
                    Top =240
                    Width =1380
                    Height =255
                    BorderColor =8355711
                    Name ="lblLoaded"
                    Caption ="Loaded"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10020
                    LayoutCachedTop =240
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =495
                    ForeTint =100.0
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="secPageHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="secPageDetail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Width =4320
                    Height =299
                    ColumnWidth =2385
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDescription"
                    ControlSource ="strDescription"
                    StatusBarText ="Description of equipment"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =299
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4740
                    Width =2220
                    Height =299
                    ColumnWidth =2145
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtForLine"
                    ControlSource ="intOptionFor"
                    StatusBarText ="Master equipment number"
                    GridlineColor =10921638

                    LayoutCachedLeft =4740
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =299
                End
                Begin Rectangle
                    Left =10620
                    Top =60
                    Width =180
                    Height =180
                    Name ="shpLoadedBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =10620
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =240
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="secPageFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="secReportFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
