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
    OrderByOn = NotDefault
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
    ItemSuffix =19
    Left =10710
    Top =3780
    Right =17850
    Bottom =7305
    OrderBy ="strVersion DESC"
    RecSrcDt = Begin
        0xdf4312ec4fece540
    End
    RecordSource ="zstblInstanceVersion"
    Caption ="Change Password"
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
                    Caption ="Change Password"
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
            Height =2880
            Name ="secFormDetail"
            AlternateBackThemeColorIndex =1
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =1260
                    Top =420
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblOldPass"
                    Caption ="Current Password"
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedTop =420
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =780
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3300
                    Top =1500
                    Width =2580
                    Height =315
                    FontSize =12
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtConfirmPass"
                    InputMask ="Password"
                    GridlineColor =10921638

                    LayoutCachedLeft =3300
                    LayoutCachedTop =1500
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1815
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =1440
                    Top =2100
                    Width =1920
                    Height =420
                    TabIndex =3
                    ForeColor =4210752
                    Name ="cmdConfirm"
                    Caption ="Confirm"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =2100
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =2520
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
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3300
                    Top =960
                    Width =2580
                    Height =315
                    FontSize =12
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtNewPass"
                    InputMask ="Password"
                    GridlineColor =10921638

                    LayoutCachedLeft =3300
                    LayoutCachedTop =960
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =1275
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3300
                    Top =420
                    Width =2580
                    Height =315
                    FontSize =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtOldPass"
                    InputMask ="Password"
                    GridlineColor =10921638

                    LayoutCachedLeft =3300
                    LayoutCachedTop =420
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =735
                End
                Begin Label
                    OverlapFlags =85
                    Left =1260
                    Top =960
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblNewPass"
                    Caption ="New Password"
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedTop =960
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =1320
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1260
                    Top =1500
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblConfirmPass"
                    Caption ="Confirm Password"
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedTop =1500
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =1860
                    ForeTint =100.0
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =3840
                    Top =2100
                    Width =1920
                    Height =420
                    TabIndex =4
                    ForeColor =4210752
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3840
                    LayoutCachedTop =2100
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =2520
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
    DoCmd.Close acForm, "fdlgChangePassword"
    Forms(frmNavigation).SetFocus
End Sub

Private Sub cmdConfirm_Click()
    
    Dim strCorrectCurrentPassword As String
    Dim strCurrentUser As String
    Dim changeResult As Boolean
    
    strCurrentUser = Form_fdlgUserControl.GetCurrentUser()
    strCorrectCurrentPassword = Form_fdlgUserControl.GetUserPassword(strCurrentUser)
    
    
    If IsNull(txtNewPass.Value) Or txtNewPass.Value = "" _
        Or IsNull(txtConfirmPass.Value) Or txtConfirmPass.Value = "" _
        Or IsNull(txtOldPass.Value) Or txtOldPass.Value = "" Then
            MsgBox "Please fill in all information to proceed.", vbOKOnly, "Information Missing"
    ElseIf Not strCorrectCurrentPassword = txtOldPass.Value Then
        MsgBox "Current password incorrect; please check and try again.", vbOKOnly, "Change Password Failed"
        txtOldPass.Value = ""
        txtOldPass.SetFocus
    ElseIf Not txtNewPass.Value = txtConfirmPass.Value Then
        MsgBox "New password confirmation does not match. Please try again.", vbOKOnly, "Change Password Failed"
        txtNewPass.Value = ""
        txtConfirmPass.Value = ""
        txtNewPass.SetFocus
    Else
        changeResult = Form_fdlgUserControl.SetCurrentUserPassword(txtNewPass.Value)
        MsgBox "Password changed successfully.", vbOKOnly, "Password Changed"
        DoCmd.Close acForm, "fdlgChangePassword"
        Forms(frmNavigation).SetFocus
    End If
    
End Sub
