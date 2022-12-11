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
    ItemSuffix =23
    Left =405
    Top =870
    RecSrcDt = Begin
        0xd277cb5d9febe540
    End
    RecordSource ="SELECT [qryEquipTypeCustomerEquipment].[lngID], [qryEquipTypeCustomerEquipment]."
        "[strDescription], [qryEquipTypeCustomerEquipment].[strSerialNumber], [qryEquipTy"
        "peCustomerEquipment].[strEQID], [qryEquipTypeCustomerEquipment].[intMeterMono], "
        "[qryEquipTypeCustomerEquipment].[intMeterColor], [qryEquipTypeCustomerEquipment]"
        ".[ysnInStock], [qryEquipTypeCustomerEquipment].[ysnReadyForInstall], [qryEquipTy"
        "peCustomerEquipment].[intInstall] FROM qryEquipTypeCustomerEquipment; "
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
                    Left =840
                    Top =240
                    Width =3480
                    Height =255
                    BorderColor =8355711
                    Name ="lblDescription"
                    Caption ="Equipment Description"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =840
                    LayoutCachedTop =240
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =495
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =12
                    TextAlign =2
                    Left =6300
                    Top =240
                    Width =2760
                    Height =255
                    BorderColor =8355711
                    Name ="lblSerialNumber"
                    Caption ="Serial Number"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6300
                    LayoutCachedTop =240
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =495
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =12
                    TextAlign =2
                    Left =4320
                    Top =240
                    Width =1980
                    Height =255
                    BorderColor =8355711
                    Name ="lblEQID"
                    Caption ="EQ Number"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4320
                    LayoutCachedTop =240
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =495
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =12
                    TextAlign =2
                    Left =60
                    Width =11340
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    Name ="lblReportHeader"
                    Caption ="Primary Equipment"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedWidth =11400
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
                Begin Label
                    OverlapFlags =4
                    TextAlign =2
                    Left =60
                    Top =240
                    Width =780
                    Height =240
                    BorderColor =8355711
                    Name ="lblLine"
                    Caption ="Line"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =240
                    LayoutCachedWidth =840
                    LayoutCachedHeight =480
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
                    Left =840
                    Width =3480
                    Height =300
                    ColumnWidth =2385
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDescription"
                    ControlSource ="strDescription"
                    StatusBarText ="Description of equipment"
                    GridlineColor =10921638

                    LayoutCachedLeft =840
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6300
                    Width =2760
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSerialNumber"
                    ControlSource ="strSerialNumber"
                    StatusBarText ="Equipment serial number"
                    GridlineColor =10921638

                    LayoutCachedLeft =6300
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4320
                    Width =1980
                    Height =300
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtEQIDPrimary"
                    ControlSource ="strEQID"
                    StatusBarText ="Equipment asset number"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =300
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
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =780
                    Height =300
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtLine"
                    ControlSource ="lngID"
                    StatusBarText ="Primary key - equipment ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedWidth =840
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
