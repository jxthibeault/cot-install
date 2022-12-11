Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11399
    DatasheetFontHeight =11
    ItemSuffix =33
    Left =2055
    Top =2415
    OrderBy ="[strEQID]"
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
            Height =435
            BackColor =1841342
            Name ="secReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =4
                    TextAlign =1
                    Left =120
                    Top =60
                    Width =6630
                    Height =315
                    FontSize =14
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReportHeader"
                    Caption ="Onsite Datasheet"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =6750
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =-1
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
            Height =960
            Name ="secPageDetail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Top =120
                    Width =5400
                    Height =300
                    ColumnWidth =2385
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDescription"
                    ControlSource ="strDescription"
                    StatusBarText ="Description of equipment"
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =120
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =1080
                    Height =300
                    FontWeight =600
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="strEQID"
                    ControlSource ="strEQID"
                    StatusBarText ="Equipment asset number"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =420
                End
                Begin Rectangle
                    Left =1320
                    Top =480
                    Width =4800
                    Height =360
                    BorderColor =-2147483617
                    Name ="shpLocationBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =1320
                    LayoutCachedTop =480
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =840
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    Left =7920
                    Top =480
                    Width =3420
                    Height =360
                    BorderColor =-2147483617
                    Name ="shpIPBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =7920
                    LayoutCachedTop =480
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =840
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    OverlapFlags =12
                    Left =6120
                    Top =480
                    Width =1800
                    Height =360
                    BackColor =12566463
                    BorderColor =-2147483617
                    Name ="shpIPLabelBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =6120
                    LayoutCachedTop =480
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =840
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    OverlapFlags =12
                    Left =60
                    Top =480
                    Width =1260
                    Height =360
                    BackColor =12566463
                    BorderColor =-2147483617
                    Name ="shpLocationLabelBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =840
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =6120
                    Top =540
                    Width =1800
                    Height =240
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblIPAddress"
                    Caption ="IP Address"
                    GridlineColor =10921638
                    LayoutCachedLeft =6120
                    LayoutCachedTop =540
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =60
                    Top =540
                    Width =1260
                    Height =300
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblLocation"
                    Caption ="Location"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =540
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =840
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="secPageFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =2280
            Name ="secReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    OverlapFlags =12
                    Left =60
                    Top =120
                    Width =11280
                    Height =360
                    BackColor =12566463
                    BorderColor =-2147483617
                    Name ="shpNotesLabelBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =480
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =120
                    Top =180
                    Width =11220
                    Height =300
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblNotes"
                    Caption ="Notes for Follow-Up / Additional Work Needed"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =180
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    Left =60
                    Top =480
                    Width =11280
                    Height =1620
                    BorderColor =-2147483617
                    Name ="shpNotesBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =2100
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
            End
        End
    End
End
