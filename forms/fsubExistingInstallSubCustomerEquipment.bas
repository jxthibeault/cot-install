Version =20
VersionRequired =20
Begin Form
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
    Width =11339
    DatasheetFontHeight =11
    ItemSuffix =30
    Right =15735
    Bottom =11730
    OrderBy ="lngID"
    RecSrcDt = Begin
        0x90526c5e33eee540
    End
    RecordSource ="qryEquipTypeCustomerEquipment"
    Caption ="subCustomerEquipment"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
                    Width =2580
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDescription"
                    Caption ="Description"
                    GridlineColor =10921638
                    LayoutCachedLeft =900
                    LayoutCachedTop =360
                    LayoutCachedWidth =3480
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
                    Left =3540
                    Top =360
                    Width =1860
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblSerialNumber"
                    Caption ="Serial No."
                    GridlineColor =10921638
                    LayoutCachedLeft =3540
                    LayoutCachedTop =360
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5460
                    Top =360
                    Width =1140
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblEQID"
                    Caption ="EQ No."
                    GridlineColor =10921638
                    LayoutCachedLeft =5460
                    LayoutCachedTop =360
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8760
                    Top =360
                    Width =795
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblInStock"
                    Caption ="In Stock"
                    GridlineColor =10921638
                    LayoutCachedLeft =8760
                    LayoutCachedTop =360
                    LayoutCachedWidth =9555
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =9600
                    Top =360
                    Width =645
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblReady"
                    Caption ="Ready"
                    GridlineColor =10921638
                    LayoutCachedLeft =9600
                    LayoutCachedTop =360
                    LayoutCachedWidth =10245
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
                    Left =6660
                    Top =360
                    Width =1980
                    Height =240
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblAssignedLocation"
                    Caption ="Onsite Location"
                    GridlineColor =10921638
                    LayoutCachedLeft =6660
                    LayoutCachedTop =360
                    LayoutCachedWidth =8640
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
                    Name ="lblTag"
                    Caption ="Tag"
                    GridlineColor =10921638
                    LayoutCachedLeft =10320
                    LayoutCachedTop =360
                    LayoutCachedWidth =10965
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
                    Width =2580
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
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3540
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

                    LayoutCachedLeft =3540
                    LayoutCachedTop =60
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5460
                    Top =60
                    Width =1140
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

                    LayoutCachedLeft =5460
                    LayoutCachedTop =60
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9060
                    Top =120
                    TabIndex =4
                    BorderColor =10921638
                    Name ="chkInStock"
                    ControlSource ="ysnInStock"
                    StatusBarText ="Equipment in stock"
                    GridlineColor =10921638

                    LayoutCachedLeft =9060
                    LayoutCachedTop =120
                    LayoutCachedWidth =9320
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9840
                    Top =120
                    TabIndex =5
                    BorderColor =10921638
                    Name ="chkReadyForInstall"
                    ControlSource ="ysnReadyForInstall"
                    StatusBarText ="Equipment ready for install"
                    GridlineColor =10921638

                    LayoutCachedLeft =9840
                    LayoutCachedTop =120
                    LayoutCachedWidth =10100
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Enabled = NotDefault
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6660
                    Top =60
                    Width =1980
                    Height =300
                    FontSize =10
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtLocation"
                    ControlSource ="strLocation"
                    StatusBarText ="Equipment asset number"
                    GridlineColor =10921638

                    LayoutCachedLeft =6660
                    LayoutCachedTop =60
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10320
                    Top =60
                    Width =660
                    Height =300
                    TabIndex =7
                    ForeColor =4210752
                    Name ="cmdTag"
                    Caption ="Tag"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =10320
                    LayoutCachedTop =60
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =360
                    Gradient =0
                    BackThemeColorIndex =5
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

Private Sub cmdTag_Click()

    DoCmd.OpenReport "rptReadyForInstallTag", acViewPreview, , "[tblInstallEquipment.lngID]=" & lngID, acWindowNormal

End Sub

Private Sub Form_Load()

    Dim strCurrentUser As String
    Dim strUserLevel As String
    
    ' Look up current user's name from instance variables and set logged-in-as label
    strCurrentUser = Form_fdlgUserControl.GetCurrentUser()
    
    ' Look up current user's permission level
    strUserLevel = Form_fdlgUserControl.GetUserAccountType(strCurrentUser)
    
    ' Permission-based

    If strUserLevel = "Administrator" Or strUserLevel = "Development" Then
        
        txtLocation.Enabled = True
        
    End If
    
End Sub
