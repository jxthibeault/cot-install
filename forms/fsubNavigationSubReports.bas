Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =2
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7920
    DatasheetFontHeight =11
    ItemSuffix =3
    Right =17220
    Bottom =11730
    RecSrcDt = Begin
        0x61fd9ffb51ece540
    End
    RecordSource ="zstlkpReportTypes"
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
            Height =562
            BackColor =1315470
            Name ="secFormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =7320
                    Height =480
                    FontSize =16
                    FontWeight =500
                    ForeColor =16777215
                    Name ="lblFormTitle"
                    Caption ="System Reports"
                    FontName ="Verdana"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =540
                    LayoutGroup =1
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
            End
        End
        Begin Section
            Height =1320
            Name ="secFormDetail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =180
                    Top =780
                    Width =5460
                    Height =300
                    BorderColor =10921638
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cboReportName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT zstlkpReportTypes.strReportTitle FROM zstlkpReportTypes ORDER BY zstlkpRe"
                        "portTypes.strReportTitle; "
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =180
                    LayoutCachedTop =780
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =1080
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =180
                    Top =360
                    Width =7680
                    Height =315
                    FontSize =12
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblInstructions"
                    Caption ="Select the report you wish to compile, then click \"Generate Report\""
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =360
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =675
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5940
                    Top =780
                    Width =1740
                    Height =300
                    TabIndex =1
                    ForeColor =4210752
                    Name ="cmdGenerateReport"
                    Caption ="Generate Report"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =5940
                    LayoutCachedTop =780
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =1080
                    Gradient =0
                    BackColor =15983578
                    BackTint =20.0
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
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

Private Sub cmdGenerateReport_Click()

    Dim varRequestedReportObjectName As Variant
    
    varRequestedReportObjectName = DLookup("[strReportObjectName]", "zstlkpReportTypes", "[strReportTitle] = '" & cboReportName.Value & "'")
    DoCmd.OpenReport varRequestedReportObjectName, acViewPreview, , , , acWindowNormal
    
End Sub
