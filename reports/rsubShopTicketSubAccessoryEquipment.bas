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
    ItemSuffix =16
    Left =135
    Top =6240
    OrderBy ="intOptionFor"
    RecSrcDt = Begin
        0x70bfb80b7aece540
    End
    RecordSource ="SELECT qryEquipTypeAccessoriesAndStartup.strDescription, qryEquipTypeAccessories"
        "AndStartup.intOptionFor, qryEquipTypeAccessoriesAndStartup.ysnInStock, qryEquipT"
        "ypeAccessoriesAndStartup.ysnReadyForInstall, qryEquipTypeAccessoriesAndStartup.l"
        "ngID, qryEquipTypeAccessoriesAndStartup.intInstall, qryEquipTypeAccessoriesAndSt"
        "artup.ysnInStock, qryEquipTypeAccessoriesAndStartup.ysnReadyForInstall FROM qryE"
        "quipTypeAccessoriesAndStartup; "
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
                    TextAlign =2
                    Left =8820
                    Top =240
                    Width =1380
                    Height =255
                    BorderColor =8355711
                    Name ="lblInStock"
                    Caption ="In Stock"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8820
                    LayoutCachedTop =240
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =495
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =10140
                    Top =240
                    Width =1380
                    Height =255
                    BorderColor =8355711
                    Name ="lblReady"
                    Caption ="Ready"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10140
                    LayoutCachedTop =240
                    LayoutCachedWidth =11520
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
                Begin CheckBox
                    Left =9420
                    Top =60
                    TabIndex =2
                    BorderColor =10921638
                    Name ="chkInStock"
                    ControlSource ="Expr1002"
                    StatusBarText ="Equipment in stock"
                    GridlineColor =10921638

                    LayoutCachedLeft =9420
                    LayoutCachedTop =60
                    LayoutCachedWidth =9680
                    LayoutCachedHeight =300
                End
                Begin CheckBox
                    Left =10740
                    Top =60
                    TabIndex =3
                    BorderColor =10921638
                    Name ="chkReady"
                    ControlSource ="Expr1003"
                    StatusBarText ="Equipment ready for install"
                    GridlineColor =10921638

                    LayoutCachedLeft =10740
                    LayoutCachedTop =60
                    LayoutCachedWidth =11000
                    LayoutCachedHeight =300
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
