Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    RecordLocks =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11339
    DatasheetFontHeight =11
    ItemSuffix =29
    Left =825
    Top =375
    Right =13680
    Bottom =6480
    RecSrcDt = Begin
        0xdf0e36618eeee540
    End
    RecordSource ="SELECT qryEquipTypeCustomerEquipment.lngID, qryEquipTypeCustomerEquipment.strDes"
        "cription, qryEquipTypeCustomerEquipment.strSerialNumber, qryEquipTypeCustomerEqu"
        "ipment.strEQID, qryEquipTypeCustomerEquipment.intMeterMono, qryEquipTypeCustomer"
        "Equipment.intMeterColor, qryEquipTypeCustomerEquipment.ysnInStock, qryEquipTypeC"
        "ustomerEquipment.ysnReadyForInstall, qryEquipTypeCustomerEquipment.intInstall, q"
        "ryEquipTypeCustomerEquipment.strEquipmentType, qryEquipTypeCustomerEquipment.str"
        "Location FROM qryEquipTypeCustomerEquipment; "
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
                    Name ="lblLineNumber"
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
                    HorizontalAnchor =1
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
                    HorizontalAnchor =1
                    LayoutCachedLeft =5700
                    LayoutCachedTop =360
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =9420
                    Top =360
                    Width =795
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblInStock"
                    Caption ="In Stock"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =9420
                    LayoutCachedTop =360
                    LayoutCachedWidth =10215
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10320
                    Top =360
                    Width =645
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReady"
                    Caption ="Ready"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =10320
                    LayoutCachedTop =360
                    LayoutCachedWidth =10965
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7080
                    Top =360
                    Width =2160
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblAssignedLocation"
                    Caption ="Onsite Location"
                    GridlineColor =10921638
                    LayoutCachedLeft =7080
                    LayoutCachedTop =360
                    LayoutCachedWidth =9240
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
                    Name ="txtLineNumber"
                    ControlSource ="lngID"
                    StatusBarText ="Primary key - equipment ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =780
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =93
                    TextAlign =2
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
                    OverlapFlags =93
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
                    HorizontalAnchor =1

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
                    HorizontalAnchor =1

                    LayoutCachedLeft =5700
                    LayoutCachedTop =60
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9720
                    Top =120
                    TabIndex =4
                    BorderColor =10921638
                    Name ="chkInStock"
                    ControlSource ="ysnInStock"
                    StatusBarText ="Equipment in stock"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =9720
                    LayoutCachedTop =120
                    LayoutCachedWidth =9980
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10500
                    Top =120
                    TabIndex =5
                    BorderColor =10921638
                    Name ="chkReadyForInstall"
                    ControlSource ="ysnReadyForInstall"
                    StatusBarText ="Equipment ready for install"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =10500
                    LayoutCachedTop =120
                    LayoutCachedWidth =10760
                    LayoutCachedHeight =360
                End
                Begin ListBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =2820
                    Top =60
                    Height =300
                    ColumnWidth =2955
                    TabIndex =6
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="txtInstall"
                    ControlSource ="intInstall"
                    RowSourceType ="Table/Query"
                    RowSource ="qryOpenInstalls"
                    ColumnWidths ="0;5760;0"
                    StatusBarText ="Associated installation"
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =2820
                    LayoutCachedTop =60
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =360
                End
                Begin ListBox
                    Visible = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =2640
                    Top =60
                    Height =300
                    ColumnWidth =2220
                    TabIndex =7
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="txtEquipmentType"
                    ControlSource ="strEquipmentType"
                    RowSourceType ="Value List"
                    RowSource ="\"Customer Equipment\";\"Accessory\";\"Startup Supplies\";\"Customer Spare Suppl"
                        "ies\";\"Technician Equipment\""
                    StatusBarText ="Classification of equipment use case"
                    DefaultValue ="\"Customer Equipment\""
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =60
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Top =60
                    Width =2160
                    Height =300
                    FontSize =10
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtLocation"
                    ControlSource ="strLocation"
                    StatusBarText ="Equipment asset number"
                    GridlineColor =10921638

                    LayoutCachedLeft =7080
                    LayoutCachedTop =60
                    LayoutCachedWidth =9240
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
