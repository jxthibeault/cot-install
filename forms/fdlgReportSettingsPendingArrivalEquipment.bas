Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7140
    DatasheetFontHeight =11
    ItemSuffix =23
    Left =8100
    Top =5355
    Right =18795
    Bottom =11100
    RecSrcDt = Begin
        0x40d4cba4cceee540
    End
    Caption ="Report Settings"
    DatasheetFontName ="Calibri"
    Moveable =0
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
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
            BackColor =1315470
            Name ="secFormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =120
                    Width =7020
                    Height =435
                    FontSize =16
                    FontWeight =500
                    ForeColor =16777215
                    Name ="lblFormTitle"
                    Caption ="Report Settings"
                    FontName ="Verdana"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =555
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
            CanGrow = NotDefault
            Height =2280
            Name ="secFormDetail"
            AlternateBackThemeColorIndex =1
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =1140
                    Top =420
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblEquipmentType"
                    Caption ="Equipment Type"
                    GridlineColor =10921638
                    LayoutCachedLeft =1140
                    LayoutCachedTop =420
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =780
                    ForeTint =100.0
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =1440
                    Top =1560
                    Width =1920
                    Height =420
                    ForeColor =4210752
                    Name ="cmdRun"
                    Caption ="Run Report"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =1560
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =1980
                    Gradient =0
                    BackColor =-2147483607
                    BackThemeColorIndex =-1
                    BackTint =100.0
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
                Begin Label
                    OverlapFlags =85
                    Left =1140
                    Top =960
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblSortBy"
                    Caption ="Sort By"
                    GridlineColor =10921638
                    LayoutCachedLeft =1140
                    LayoutCachedTop =960
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =1320
                    ForeTint =100.0
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =3840
                    Top =1560
                    Width =1920
                    Height =420
                    TabIndex =1
                    ForeColor =4210752
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3840
                    LayoutCachedTop =1560
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =1980
                    Gradient =0
                    BackColor =-2147483607
                    BackThemeColorIndex =-1
                    BackTint =100.0
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
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3180
                    Top =420
                    Width =2820
                    Height =330
                    FontSize =12
                    TabIndex =2
                    BorderColor =10921638
                    Name ="cboEquipmentType"
                    RowSourceType ="Value List"
                    RowSource ="\"All Equipment\";\"Customer Equipment\";\"Accessory\";\"Startup Supplies\";\"Cu"
                        "stomer Spare Supplies\";\"Technician Equipment\""
                    DefaultValue ="\"All Equipment\""
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =3180
                    LayoutCachedTop =420
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =750
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3180
                    Top =960
                    Width =2820
                    Height =330
                    FontSize =12
                    TabIndex =3
                    BorderColor =10921638
                    Name ="cboSortBy"
                    RowSourceType ="Value List"
                    RowSource ="\"Equipment Description\";\"Customer\";\"Equipment Type\";\"Date Requested\""
                    DefaultValue ="\"Equipment Description\""
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =3180
                    LayoutCachedTop =960
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =1290
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

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, "fdlgReportSettingsPendingArrivalEquipment"
    Forms(frmNavigation).SetFocus
End Sub

Private Sub cmdRun_Click()
    
    Dim strEquipmentType As String
    Dim strRequestedSort As String
    
    Dim strFilter As String
    Dim strSort As String
    
    strEquipmentType = cboEquipmentType.Value
    strRequestedSort = cboSortBy.Value
    strFilter = ""
    strSort = ""
    
    If Not strEquipmentType = "All Equipment" Then
        strFilter = "[strEquipmentType] = '" & strEquipmentType & "'"
    End If
    
    If strRequestedSort = "Equipment Description" Then
        strSort = "strDescription"
    ElseIf strRequestedSort = "Customer" Then
        strSort = "strCustomer"
    ElseIf strRequestedSort = "Equipment Type" Then
        strSort = "strEquipmentType"
    ElseIf strRequestedSort = "Date Requested" Then
        strSort = "dtmDateReceived"
    End If
    
    DoCmd.OpenReport "rptPendingArrivalEquipment", acViewPreview, , strFilter, acWindowNormal, strSort
    DoCmd.Close acForm, "fdlgReportSettingsPendingArrivalEquipment"
        
End Sub
