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
    Width =6120
    DatasheetFontHeight =11
    ItemSuffix =11
    Right =15135
    Bottom =11730
    OrderBy ="intOptionFor"
    RecSrcDt = Begin
        0xd8784dae9debe540
    End
    RecordSource ="SELECT [qryEquipTypeAccessoriesAndStartup].strDescription, [qryEquipTypeAccessor"
        "iesAndStartup].intOptionFor, [qryEquipTypeAccessoriesAndStartup].ysnInStock, [qr"
        "yEquipTypeAccessoriesAndStartup].ysnReadyForInstall, [qryEquipTypeAccessoriesAnd"
        "Startup].lngID, [qryEquipTypeAccessoriesAndStartup].intInstall FROM qryEquipType"
        "AccessoriesAndStartup; "
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
                    Width =2175
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFormHeader"
                    Caption ="Accessories"
                    FontName ="Verdana"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =2415
                    LayoutCachedHeight =435
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =240
                    Top =360
                    Width =2820
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDescription"
                    Caption ="Description"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =360
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3240
                    Top =360
                    Width =870
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblForLine"
                    Caption ="For Line:"
                    GridlineColor =10921638
                    LayoutCachedLeft =3240
                    LayoutCachedTop =360
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4380
                    Top =360
                    Width =795
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblInStock"
                    Caption ="In Stock"
                    GridlineColor =10921638
                    LayoutCachedLeft =4380
                    LayoutCachedTop =360
                    LayoutCachedWidth =5175
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5280
                    Top =360
                    Width =645
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReady"
                    Caption ="Ready"
                    GridlineColor =10921638
                    LayoutCachedLeft =5280
                    LayoutCachedTop =360
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =375
            Name ="secFormDetail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =240
                    Top =60
                    Width =2820
                    Height =299
                    ColumnWidth =3000
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDescription"
                    ControlSource ="strDescription"
                    StatusBarText ="Description of equipment"
                    GridlineColor =10921638

                    LayoutCachedLeft =240
                    LayoutCachedTop =60
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =359
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4680
                    Top =120
                    TabIndex =2
                    BorderColor =10921638
                    Name ="chkInStock"
                    ControlSource ="ysnInStock"
                    StatusBarText ="Equipment in stock"
                    GridlineColor =10921638

                    LayoutCachedLeft =4680
                    LayoutCachedTop =120
                    LayoutCachedWidth =4940
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5520
                    Top =120
                    TabIndex =3
                    BorderColor =10921638
                    Name ="chkReadyForInstall"
                    ControlSource ="ysnReadyForInstall"
                    StatusBarText ="Equipment ready for install"
                    GridlineColor =10921638

                    LayoutCachedLeft =5520
                    LayoutCachedTop =120
                    LayoutCachedWidth =5780
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3240
                    Top =60
                    Width =840
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtForLine"
                    ControlSource ="intOptionFor"
                    StatusBarText ="Master equipment number"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =60
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =375
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
