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
    Width =11399
    DatasheetFontHeight =11
    ItemSuffix =49
    Top =600
    RecSrcDt = Begin
        0x5aae91bf5aece540
    End
    RecordSource ="tblInstalls"
    Caption ="subrptCustomerEquipment"
    OnOpen ="[Event Procedure]"
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
            Height =450
            BackColor =1841342
            Name ="secReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    TextAlign =1
                    Left =120
                    Top =60
                    Width =3090
                    Height =390
                    FontSize =14
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReportHeader"
                    Caption ="Installation Acceptance"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =3210
                    LayoutCachedHeight =450
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
            Height =7560
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
                    Top =360
                    Width =11160
                    Height =315
                    FontWeight =600
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtCustomer"
                    ControlSource ="strCustomer"
                    StatusBarText ="Customer name as it appears on legal documents"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =360
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =675
                End
                Begin Label
                    Left =120
                    Top =120
                    Width =11160
                    Height =3120
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblCDAText"
                    Caption ="By signing below, the undersigned representative of:\015\012\015\012confirms tha"
                        "t installation of all equipment, in consideration of any exceptions described be"
                        "low (see \"Exceptions\"), has been satisfactorily delivered, installed configure"
                        "d and tested by Connected Office Technologies.\015\012\015\012The undersigned ac"
                        "knowledges that although all equipment has been tested by a trained technician, "
                        "Connected Office Technologies strongly suggests that end users test all critical"
                        " features of devices prior to technician departure.\015\012\015\012The undersign"
                        "ed acknowledges that they have been offered training relevant to the installed e"
                        "quipment and is able to use the installed equipment in such a way that satisfies"
                        " routine day-to-day business needs; and that Connected Office Technologies has s"
                        "atisfactorily completed installation of the equipment."
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =3240
                End
                Begin Rectangle
                    Left =120
                    Top =3780
                    Width =11160
                    Height =3360
                    BorderColor =-2147483617
                    Name ="shpExceptionsBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =3780
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =7140
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    Left =120
                    Top =3420
                    Width =11160
                    Height =360
                    BackColor =12566463
                    BorderColor =-2147483617
                    Name ="shpExceptionsLabelBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =3420
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =3780
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =180
                    Top =3480
                    Width =1260
                    Height =300
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblExceptions"
                    Caption ="Exceptions"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =3480
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =3780
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
            Height =2820
            Name ="secReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    Left =5100
                    Top =120
                    Width =6180
                    Height =576
                    BorderColor =-2147483617
                    Name ="shpNameBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =5100
                    LayoutCachedTop =120
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =696
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    Left =3840
                    Top =120
                    Width =1260
                    Height =576
                    BackColor =12566463
                    BorderColor =-2147483617
                    Name ="shpNameLabelBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =3840
                    LayoutCachedTop =120
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =696
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =3840
                    Top =180
                    Width =1260
                    Height =300
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblName"
                    Caption ="Name"
                    GridlineColor =10921638
                    LayoutCachedLeft =3840
                    LayoutCachedTop =180
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    Left =5100
                    Top =780
                    Width =6180
                    Height =576
                    BorderColor =-2147483617
                    Name ="shpTitleBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =5100
                    LayoutCachedTop =780
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =1356
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    Left =3840
                    Top =780
                    Width =1260
                    Height =576
                    BackColor =12566463
                    BorderColor =-2147483617
                    Name ="shpTitleLabelBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =3840
                    LayoutCachedTop =780
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =1356
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =3840
                    Top =840
                    Width =1260
                    Height =300
                    BackColor =12566463
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblTitle"
                    Caption ="Title"
                    GridlineColor =10921638
                    LayoutCachedLeft =3840
                    LayoutCachedTop =840
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =1140
                    BackShade =75.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Rectangle
                    Left =5100
                    Top =1440
                    Width =6180
                    Height =960
                    BorderColor =-2147483617
                    Name ="shpSignatureBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =5100
                    LayoutCachedTop =1440
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =2400
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Rectangle
                    Left =3840
                    Top =1440
                    Width =1260
                    Height =960
                    BackColor =12566463
                    BorderColor =-2147483617
                    Name ="shpSignatureLabelBox"
                    GridlineColor =10921638
                    LayoutCachedLeft =3840
                    LayoutCachedTop =1440
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =2400
                    BackShade =75.0
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =3840
                    Top =1500
                    Width =1260
                    Height =300
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblSignature"
                    Caption ="Signature"
                    GridlineColor =10921638
                    LayoutCachedLeft =3840
                    LayoutCachedTop =1500
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =1800
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =120
                    Top =960
                    Width =3540
                    Height =540
                    FontSize =14
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =-2147483617
                    Name ="lblSignHere"
                    Caption ="PLEASE SIGN HERE:"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =960
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =1500
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
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
