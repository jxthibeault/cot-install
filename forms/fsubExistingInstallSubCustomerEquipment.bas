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
    ItemSuffix =27
    Right =18495
    Bottom =11985
    RecSrcDt = Begin
        0x60d7bad59cebe540
    End
    RecordSource ="SELECT [qryEquipTypeCustomerEquipment].[lngID], [qryEquipTypeCustomerEquipment]."
        "[strDescription], [qryEquipTypeCustomerEquipment].[strSerialNumber], [qryEquipTy"
        "peCustomerEquipment].[strEQID], [qryEquipTypeCustomerEquipment].[intMeterMono], "
        "[qryEquipTypeCustomerEquipment].[intMeterColor], [qryEquipTypeCustomerEquipment]"
        ".[ysnInStock], [qryEquipTypeCustomerEquipment].[ysnReadyForInstall], [qryEquipTy"
        "peCustomerEquipment].[intInstall] FROM qryEquipTypeCustomerEquipment; "
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
            Height =659
            BackColor =1841342
            Name ="secFormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =900
                    Top =360
                    Width =2700
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDescription"
                    Caption ="Description"
                    GridlineColor =10921638
                    LayoutCachedLeft =900
                    LayoutCachedTop =360
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =240
                    Top =360
                    Width =540
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblLine"
                    Caption ="Line"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =360
                    LayoutCachedWidth =780
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3720
                    Top =360
                    Width =1860
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblSerialNumber"
                    Caption ="Serial No."
                    GridlineColor =10921638
                    LayoutCachedLeft =3720
                    LayoutCachedTop =360
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5700
                    Top =360
                    Width =1260
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblEQID"
                    Caption ="EQ No."
                    GridlineColor =10921638
                    LayoutCachedLeft =5700
                    LayoutCachedTop =360
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =7080
                    Top =120
                    Width =1920
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblMetersHeader"
                    Caption ="Meters"
                    GridlineColor =10921638
                    LayoutCachedLeft =7080
                    LayoutCachedTop =120
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =435
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7080
                    Top =360
                    Width =900
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblMeterMono"
                    Caption ="Mono"
                    GridlineColor =10921638
                    LayoutCachedLeft =7080
                    LayoutCachedTop =360
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =8100
                    Top =360
                    Width =900
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblMeterColor"
                    Caption ="Color"
                    GridlineColor =10921638
                    LayoutCachedLeft =8100
                    LayoutCachedTop =360
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =9000
                    Top =360
                    Width =795
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblInStock"
                    Caption ="In Stock"
                    GridlineColor =10921638
                    LayoutCachedLeft =9000
                    LayoutCachedTop =360
                    LayoutCachedWidth =9795
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =9900
                    Top =360
                    Width =645
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReady"
                    Caption ="Ready"
                    GridlineColor =10921638
                    LayoutCachedLeft =9900
                    LayoutCachedTop =360
                    LayoutCachedWidth =10545
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =240
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
                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =2415
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
                    Left =240
                    Top =60
                    Width =540
                    Height =315
                    ColumnWidth =1440
                    FontSize =10
                    FontWeight =100
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtRecordId"
                    ControlSource ="lngID"
                    StatusBarText ="Primary key - equipment ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =780
                    LayoutCachedHeight =375
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
                    Left =900
                    Top =60
                    Width =2700
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDescription"
                    ControlSource ="strDescription"
                    StatusBarText ="Description of equipment"
                    GridlineColor =10921638

                    LayoutCachedLeft =900
                    LayoutCachedTop =60
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3720
                    Top =60
                    Width =1860
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSerialNumber"
                    ControlSource ="strSerialNumber"
                    StatusBarText ="Equipment serial number"
                    GridlineColor =10921638

                    LayoutCachedLeft =3720
                    LayoutCachedTop =60
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =1260
                    Height =300
                    ColumnWidth =3000
                    FontSize =10
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtEQID"
                    ControlSource ="strEQID"
                    StatusBarText ="Equipment asset number"
                    GridlineColor =10921638

                    LayoutCachedLeft =5700
                    LayoutCachedTop =60
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Top =60
                    Width =900
                    Height =299
                    ColumnWidth =1050
                    FontSize =10
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtMeterMono"
                    ControlSource ="intMeterMono"
                    StatusBarText ="Starting mono meter for used equipment"
                    GridlineColor =10921638

                    LayoutCachedLeft =7080
                    LayoutCachedTop =60
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =359
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8100
                    Top =60
                    Width =900
                    Height =299
                    ColumnWidth =1050
                    FontSize =10
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtMeterColor"
                    ControlSource ="intMeterColor"
                    StatusBarText ="Starting color meter for used equipment"
                    GridlineColor =10921638

                    LayoutCachedLeft =8100
                    LayoutCachedTop =60
                    LayoutCachedWidth =9000
                    LayoutCachedHeight =359
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9300
                    Top =120
                    TabIndex =6
                    BorderColor =10921638
                    Name ="chkInStock"
                    ControlSource ="ysnInStock"
                    StatusBarText ="Equipment in stock"
                    GridlineColor =10921638

                    LayoutCachedLeft =9300
                    LayoutCachedTop =120
                    LayoutCachedWidth =9560
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10080
                    Top =120
                    TabIndex =7
                    BorderColor =10921638
                    Name ="chkReadyForInstall"
                    ControlSource ="ysnReadyForInstall"
                    StatusBarText ="Equipment ready for install"
                    GridlineColor =10921638

                    LayoutCachedLeft =10080
                    LayoutCachedTop =120
                    LayoutCachedWidth =10340
                    LayoutCachedHeight =360
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
