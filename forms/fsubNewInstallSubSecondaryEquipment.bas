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
    Width =10800
    DatasheetFontHeight =11
    ItemSuffix =33
    Left =2880
    Top =5520
    Right =13590
    Bottom =12270
    RecSrcDt = Begin
        0x3547a5ba79ece540
    End
    RecordSource ="SELECT tblInstallEquipment.strEquipmentType, tblInstallEquipment.intQuantity, tb"
        "lInstallEquipment.strDescription, tblInstallEquipment.intOptionFor, tblInstallEq"
        "uipment.intInstall, tblInstallEquipment.ysnReadyForInstall, tblInstallEquipment."
        "ysnInStock FROM tblInstallEquipment WHERE (((tblInstallEquipment.strEquipmentTyp"
        "e)<>\"Customer Equipment\")); "
    Caption ="subAccessoryEquipment"
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
        Begin ComboBox
            AddColon = NotDefault
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
            ForeThemeColorIndex =2
            ForeShade =50.0
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
                    TextAlign =1
                    Left =240
                    Top =120
                    Width =2445
                    Height =345
                    FontSize =12
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFormTitle"
                    Caption ="Secondary Equipment"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =2685
                    LayoutCachedHeight =465
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =240
                    Top =360
                    Width =360
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblQuantity"
                    Caption ="Qty"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =360
                    LayoutCachedWidth =600
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =780
                    Top =360
                    Width =1680
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDescription"
                    Caption ="Description"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =780
                    LayoutCachedTop =360
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4020
                    Top =360
                    Width =1200
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblForLine"
                    Caption ="For Line"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =4020
                    LayoutCachedTop =360
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5340
                    Top =360
                    Width =2700
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblItemType"
                    Caption ="Item Type"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =5340
                    LayoutCachedTop =360
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8400
                    Top =360
                    Width =900
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblInStock"
                    Caption ="In Stock"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =8400
                    LayoutCachedTop =360
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =9420
                    Top =360
                    Width =600
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReady"
                    Caption ="Ready"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =9420
                    LayoutCachedTop =360
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =420
            Name ="secFormDetail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =7080
                    Top =60
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtRecordId"
                    ControlSource ="lngID"
                    StatusBarText ="Primary key - equipment ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =7080
                    LayoutCachedTop =60
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =360
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtQuantity"
                    ControlSource ="intQuantity"
                    StatusBarText ="Equipment quantity"
                    DefaultValue ="1"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =600
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =780
                    Top =60
                    Width =3060
                    Height =315
                    ColumnWidth =2385
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDescription"
                    ControlSource ="strDescription"
                    StatusBarText ="Description of equipment"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =780
                    LayoutCachedTop =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4020
                    Top =60
                    Width =1140
                    Height =315
                    ColumnWidth =2145
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtForLine"
                    ControlSource ="intOptionFor"
                    StatusBarText ="Master equipment number"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =4020
                    LayoutCachedTop =60
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =375
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =8640
                    Top =120
                    TabIndex =5
                    BorderColor =10921638
                    Name ="ysnInStock"
                    ControlSource ="ysnInStock"
                    StatusBarText ="Equipment in stock"
                    DefaultValue ="False"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =8640
                    LayoutCachedTop =120
                    LayoutCachedWidth =8900
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9600
                    Top =120
                    TabIndex =6
                    BorderColor =10921638
                    Name ="ysnReadyForInstall"
                    ControlSource ="ysnReadyForInstall"
                    StatusBarText ="Equipment ready for install"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =9600
                    LayoutCachedTop =120
                    LayoutCachedWidth =9860
                    LayoutCachedHeight =360
                End
                Begin ListBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =7260
                    Top =60
                    Height =300
                    ColumnWidth =2955
                    TabIndex =7
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

                    LayoutCachedLeft =7260
                    LayoutCachedTop =60
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =360
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =5340
                    Top =60
                    Width =2880
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    Name ="cboEquipmentType"
                    ControlSource ="strEquipmentType"
                    RowSourceType ="Value List"
                    RowSource ="\"Accessory\";\"Startup Supplies\";\"Customer Spare Supplies\";\"Technician Equi"
                        "pment\""
                    StatusBarText ="Classification of equipment use case"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    AllowValueListEdits =0

                    LayoutCachedLeft =5340
                    LayoutCachedTop =60
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                End
            End
        End
        Begin FormFooter
            Visible = NotDefault
            Height =0
            Name ="secFormFooter"
            AutoHeight =1
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
