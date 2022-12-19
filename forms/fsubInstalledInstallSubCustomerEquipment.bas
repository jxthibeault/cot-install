Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =11
    ItemSuffix =30
    Left =3735
    Top =2565
    Right =14445
    Bottom =6975
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
            Height =659
            BackColor =1841342
            Name ="secFormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =660
                    Top =360
                    Width =2460
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDescription"
                    Caption ="Description"
                    GridlineColor =10921638
                    LayoutCachedLeft =660
                    LayoutCachedTop =360
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =120
                    Top =360
                    Width =480
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblLine"
                    Caption ="Line"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =360
                    LayoutCachedWidth =600
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3180
                    Top =360
                    Width =1860
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblSerialNumber"
                    Caption ="Serial No."
                    GridlineColor =10921638
                    LayoutCachedLeft =3180
                    LayoutCachedTop =360
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5100
                    Top =360
                    Width =1140
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblEQID"
                    Caption ="EQ No."
                    GridlineColor =10921638
                    LayoutCachedLeft =5100
                    LayoutCachedTop =360
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =6300
                    Top =120
                    Width =1620
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblMetersHeader"
                    Caption ="Meters"
                    GridlineColor =10921638
                    LayoutCachedLeft =6300
                    LayoutCachedTop =120
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =435
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =6300
                    Top =360
                    Width =780
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblMeterMono"
                    Caption ="Mono"
                    GridlineColor =10921638
                    LayoutCachedLeft =6300
                    LayoutCachedTop =360
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7140
                    Top =360
                    Width =780
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblMeterColor"
                    Caption ="Color"
                    GridlineColor =10921638
                    LayoutCachedLeft =7140
                    LayoutCachedTop =360
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =120
                    Top =120
                    Width =2175
                    Height =315
                    FontSize =12
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFormTitle"
                    Caption ="Primary Equipment"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =435
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
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =480
                    Height =315
                    ColumnWidth =1440
                    FontSize =10
                    FontWeight =100
                    Name ="txtRecordId"
                    ControlSource ="lngID"
                    StatusBarText ="Primary key - equipment ID"

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =600
                    LayoutCachedHeight =375
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =660
                    Top =60
                    Width =2460
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =1
                    Name ="txtDescription"
                    ControlSource ="strDescription"
                    StatusBarText ="Description of equipment"

                    LayoutCachedLeft =660
                    LayoutCachedTop =60
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3180
                    Top =60
                    Width =1860
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =2
                    Name ="txtSerialNumber"
                    ControlSource ="strSerialNumber"
                    StatusBarText ="Equipment serial number"

                    LayoutCachedLeft =3180
                    LayoutCachedTop =60
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5100
                    Top =60
                    Width =1140
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =3
                    Name ="txtEQID"
                    ControlSource ="strEQID"
                    StatusBarText ="Equipment asset number"

                    LayoutCachedLeft =5100
                    LayoutCachedTop =60
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =360
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6300
                    Top =60
                    Width =780
                    Height =299
                    ColumnWidth =1050
                    FontSize =10
                    TabIndex =4
                    Name ="txtMeterMono"
                    ControlSource ="intMeterMono"
                    StatusBarText ="Starting mono meter for used equipment"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =60
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =359
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7140
                    Top =60
                    Width =780
                    Height =299
                    ColumnWidth =1050
                    FontSize =10
                    TabIndex =5
                    Name ="txtMeterColor"
                    ControlSource ="intMeterColor"
                    StatusBarText ="Starting color meter for used equipment"

                    LayoutCachedLeft =7140
                    LayoutCachedTop =60
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =359
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7980
                    Top =180
                    Width =2700
                    FontSize =10
                    TabIndex =7
                    Name ="txtIpAddress"
                    ControlSource ="strIpAddress"
                    StatusBarText ="Description of equipment"
                    TopPadding =0
                    HorizontalAnchor =2

                    LayoutCachedLeft =7980
                    LayoutCachedTop =180
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =420
                    BackThemeColorIndex =1
                    BorderThemeColorIndex =1
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7980
                    Width =2700
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="txtLocation"
                    ControlSource ="strLocation"
                    StatusBarText ="Description of equipment"
                    TopPadding =0
                    HorizontalAnchor =2

                    LayoutCachedLeft =7980
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =300
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
