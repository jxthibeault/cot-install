Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =34
    Left =2160
    Top =5580
    Right =13530
    Bottom =11130
    OrderBy ="[strEQID]"
    RecSrcDt = Begin
        0x0990f03733eee540
    End
    RecordSource ="qryEquipTypeCustomerEquipment"
    Caption ="subCustomerEquipment"
    DatasheetFontName ="Calibri"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
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
            BorderShade =65.0
            ThemeFontIndex =1
            ForeTint =75.0
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
            Height =660
            BackColor =1841342
            Name ="secFormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =1320
                    Top =360
                    Width =2460
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDescription"
                    Caption ="Description"
                    GridlineColor =10921638
                    LayoutCachedLeft =1320
                    LayoutCachedTop =360
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =3840
                    Top =360
                    Width =1860
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblSerialNumber"
                    Caption ="Serial No."
                    GridlineColor =10921638
                    LayoutCachedLeft =3840
                    LayoutCachedTop =360
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =120
                    Top =360
                    Width =1140
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblEQID"
                    Caption ="EQ No."
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =360
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =120
                    Top =120
                    Width =5400
                    Height =315
                    FontSize =12
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFormTitle"
                    Caption ="Onsite Worksheet Information Entry"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =435
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5760
                    Top =360
                    Width =3060
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblLocation"
                    Caption ="On-site Location"
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedTop =360
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8940
                    Top =360
                    Width =2100
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblIpAddress"
                    Caption ="IP Address"
                    GridlineColor =10921638
                    LayoutCachedLeft =8940
                    LayoutCachedTop =360
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =420
            Name ="secFormDetail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =2460
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =1
                    Name ="txtDescription"
                    ControlSource ="strDescription"
                    StatusBarText ="Description of equipment"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3840
                    Top =60
                    Width =1860
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =2
                    Name ="txtSerialNumber"
                    ControlSource ="strSerialNumber"
                    StatusBarText ="Equipment serial number"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =60
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =1140
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    Name ="txtEQID"
                    ControlSource ="strEQID"
                    StatusBarText ="Equipment asset number"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5760
                    Top =60
                    Width =3060
                    Height =300
                    FontSize =10
                    TabIndex =3
                    Name ="txtLocation"
                    ControlSource ="strLocation"
                    StatusBarText ="Equipment serial number"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =60
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8940
                    Top =60
                    Width =2100
                    Height =300
                    FontSize =10
                    TabIndex =4
                    Name ="txtIpAddress"
                    ControlSource ="strIpAddress"
                    StatusBarText ="Equipment serial number"

                    LayoutCachedLeft =8940
                    LayoutCachedTop =60
                    LayoutCachedWidth =11040
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
            Height =0
            Name ="secFormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
