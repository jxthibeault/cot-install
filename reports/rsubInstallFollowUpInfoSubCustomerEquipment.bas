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
    ItemSuffix =25
    Top =900
    RecSrcDt = Begin
        0xef5bd55935eee540
    End
    RecordSource ="qryEquipTypeCustomerEquipment"
    Caption ="subrptCustomerEquipment"
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
            Height =540
            Name ="secReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =75.0
            Begin
                Begin Label
                    TextAlign =2
                    Left =2100
                    Top =240
                    Width =3120
                    Height =255
                    BorderColor =8355711
                    Name ="lblDescription"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2100
                    LayoutCachedTop =240
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =495
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =5280
                    Top =240
                    Width =2340
                    Height =255
                    BorderColor =8355711
                    Name ="lblSerialNumber"
                    Caption ="Serial Number"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5280
                    LayoutCachedTop =240
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =495
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =60
                    Top =240
                    Width =1980
                    Height =255
                    BorderColor =8355711
                    Name ="lblEQID"
                    Caption ="EQ Number"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =240
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =495
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =4620
                    Width =2235
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    Name ="lblReportHeader"
                    Caption ="Installed Equipment"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4620
                    LayoutCachedWidth =6855
                    LayoutCachedHeight =315
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =7680
                    Top =240
                    Width =3720
                    Height =255
                    BorderColor =8355711
                    Name ="lblLocation"
                    Caption ="Installed Location"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7680
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
                    Left =2100
                    Width =3120
                    Height =300
                    ColumnWidth =2385
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDescription"
                    ControlSource ="strDescription"
                    StatusBarText ="Description of equipment"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5280
                    Width =2340
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSerialNumber"
                    ControlSource ="strSerialNumber"
                    StatusBarText ="Equipment serial number"
                    GridlineColor =10921638

                    LayoutCachedLeft =5280
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =1980
                    Height =300
                    FontWeight =700
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtEQIDPrimary"
                    ControlSource ="strEQID"
                    StatusBarText ="Equipment asset number"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7680
                    Width =3720
                    Height =300
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtLocation"
                    ControlSource ="strLocation"
                    StatusBarText ="Equipment serial number"
                    GridlineColor =10921638

                    LayoutCachedLeft =7680
                    LayoutCachedWidth =11400
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
