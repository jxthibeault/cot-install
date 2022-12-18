Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =2
    RecordLocks =2
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =82
    Left =7020
    Top =2280
    Right =20160
    Bottom =14010
    RecSrcDt = Begin
        0x830b02b29eebe540
    End
    RecordSource ="SELECT tblInstalls.*, tblInstalls.attAttachments FROM tblInstalls; "
    Caption ="Post-Install Info Collection"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =255
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
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
        Begin Subform
            BorderLineStyle =0
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin Attachment
            BackStyle =0
            BorderLineStyle =0
            PictureSizeMode =3
            Width =4800
            Height =3840
            LabelX =-1800
            AddColon =0
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            CanShrink = NotDefault
            Height =839
            BackColor =1841342
            Name ="secFormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =180
                    Top =180
                    Width =5940
                    Height =480
                    FontSize =16
                    FontWeight =600
                    ForeColor =16777215
                    Name ="lblFormTitle"
                    Caption ="Enter Post-Install Information"
                    FontName ="Verdana"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =180
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =660
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
            CanShrink = NotDefault
            Height =9360
            Name ="secFormDetail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3480
                    Top =600
                    Width =3240
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtITContactName"
                    ControlSource ="strITContactName"
                    StatusBarText ="IT contact name"
                    GridlineColor =10921638

                    LayoutCachedLeft =3480
                    LayoutCachedTop =600
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =915
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3480
                    Top =1020
                    Width =3240
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtITContactPhone"
                    ControlSource ="strITContactPhone"
                    StatusBarText ="IT contact phone number"
                    GridlineColor =10921638

                    LayoutCachedLeft =3480
                    LayoutCachedTop =1020
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =1335
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3480
                    Top =1440
                    Width =3240
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtrITContactEmail"
                    ControlSource ="strITContactEmail"
                    StatusBarText ="IT contact email address"
                    GridlineColor =10921638

                    LayoutCachedLeft =3480
                    LayoutCachedTop =1440
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =1755
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3480
                    Top =180
                    Width =3240
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtInstallScheduled"
                    ControlSource ="dtmInstallScheduled"
                    StatusBarText ="Scheduled date of installation"
                    GridlineColor =10921638

                    LayoutCachedLeft =3480
                    LayoutCachedTop =180
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =495
                End
                Begin Label
                    OverlapFlags =85
                    Left =300
                    Top =600
                    Width =2340
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblITSupportContact"
                    Caption ="IT Support Name:"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =600
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =915
                End
                Begin Line
                    OverlapFlags =85
                    Top =3240
                    Width =11520
                    Name ="linHorizontalCenter"
                    GridlineColor =10921638
                    LayoutCachedTop =3240
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =3240
                End
                Begin Label
                    OverlapFlags =85
                    Left =300
                    Top =180
                    Width =3060
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblInstallScheduled"
                    Caption ="Actual Date of Installation:"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =180
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =495
                End
                Begin Subform
                    OverlapFlags =85
                    Left =300
                    Top =3360
                    Width =10979
                    Height =5820
                    TabIndex =4
                    BorderColor =-2147483617
                    Name ="subCustomerEquipment"
                    SourceObject ="Form.fsubPostInstallEquipment"
                    LinkChildFields ="intInstall"
                    LinkMasterFields ="lngID"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =3360
                    LayoutCachedWidth =11279
                    LayoutCachedHeight =9180
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7200
                    Top =180
                    Width =4080
                    Height =480
                    FontWeight =500
                    TabIndex =5
                    ForeColor =4210752
                    Name ="cmdSaveAndComplete"
                    Caption ="Save and Mark Complete"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =7200
                    LayoutCachedTop =180
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =660
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
                    Overlaps =1
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =300
                    Top =2160
                    Width =10980
                    Height =962
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtPostInstallNotes"
                    ControlSource ="memPostInstallNotes"
                    StatusBarText ="Internal install notes"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =2160
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =3122
                End
                Begin Label
                    OverlapFlags =85
                    Left =300
                    Top =1020
                    Width =2340
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblItSupportPhone"
                    Caption ="IT Phone Number:"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1335
                End
                Begin Label
                    OverlapFlags =85
                    Left =300
                    Top =1455
                    Width =2340
                    Height =315
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblItSupportEmail"
                    Caption ="IT Email Address:"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =1455
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1770
                End
                Begin Label
                    OverlapFlags =247
                    Left =300
                    Top =1860
                    Width =4680
                    Height =315
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lblPostInstallNotes"
                    Caption ="Post-Install and Follow-Up Notes:"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =1860
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =2175
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7200
                    Top =840
                    Width =4080
                    Height =480
                    FontWeight =500
                    TabIndex =7
                    ForeColor =4210752
                    Name ="cmdSaveForLaterEntry"
                    Caption ="Save for Later Entry"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =7200
                    LayoutCachedTop =840
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =1320
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
            End
        End
        Begin FormFooter
            Visible = NotDefault
            CanShrink = NotDefault
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

Private bInstallComplete As Boolean

Private Sub cmdSaveAndComplete_Click()
    
    bInstallComplete = True
    Me.Visible = False
    
End Sub

Private Sub cmdSaveForLaterEntry_Click()

    Me.Visible = False

End Sub

Public Function GetInstallComplete() As Boolean

    GetInstallComplete = bInstallComplete

End Function



Private Sub Form_Load()

    Dim strCurrentUser As String
    Dim strUserLevel As String
    
    bInstallComplete = False
    
    ' Look up current user's name from instance variables and set logged-in-as label
    strCurrentUser = Form_fdlgUserControl.GetCurrentUser()
    
    ' Look up current user's permission level
    strUserLevel = Form_fdlgUserControl.GetUserAccountType(strCurrentUser)
    
    ' Permission-based

    If strUserLevel = "Administrator" Or strUserLevel = "Development" Then
        

        
    End If

End Sub
