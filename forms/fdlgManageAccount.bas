Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
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
    Cycle =1
    GridX =24
    GridY =24
    Width =7140
    DatasheetFontHeight =11
    ItemSuffix =26
    Left =7005
    Top =3630
    Right =14145
    Bottom =8835
    Filter ="[ID]=1"
    RecSrcDt = Begin
        0x121fd9b9d3efe540
    End
    RecordSource ="tblUsers"
    Caption ="Manage Account"
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
                    Caption ="Manage Account"
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
            Height =4560
            Name ="secFormDetail"
            AlternateBackThemeColorIndex =1
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =360
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblUsername"
                    Caption ="Username"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =360
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =720
                    ForeTint =100.0
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3240
                    Top =1320
                    Width =2700
                    Height =315
                    FontSize =12
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtTitle"
                    ControlSource ="strTitle"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =1320
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1635
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =2640
                    Top =3780
                    Width =1920
                    Height =420
                    TabIndex =4
                    ForeColor =4210752
                    Name ="cmdConfirm"
                    Caption ="Confirm"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2640
                    LayoutCachedTop =3780
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =4200
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
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3240
                    Top =840
                    Width =2700
                    Height =315
                    FontSize =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDisplayName"
                    ControlSource ="strDisplayName"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =840
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1155
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3240
                    Top =360
                    Width =2700
                    Height =315
                    FontSize =12
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtUsername"
                    ControlSource ="strUsername"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =360
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =675
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =840
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblDisplayName"
                    Caption ="Display Name"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =840
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =1200
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =1320
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblTitle"
                    Caption ="Job Title"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =1320
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =1680
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =2220
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblAccountType"
                    Caption ="Account Type"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =2220
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =2580
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3240
                    Top =2220
                    Width =2700
                    Height =330
                    FontSize =12
                    TabIndex =2
                    BorderColor =10921638
                    Name ="cboAccountType"
                    ControlSource ="strAccountType"
                    RowSourceType ="Value List"
                    RowSource ="\"Administrator\";\"Technician\""
                    DefaultValue ="\"Technician\""
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =3240
                    LayoutCachedTop =2220
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =2550
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    ForeShade =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =2700
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblResetPassword"
                    Caption ="Reset Password"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =2700
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =3060
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3240
                    Top =2700
                    Width =2700
                    Height =314
                    FontSize =12
                    TabIndex =3
                    Name ="cmdResetPassword"
                    Caption ="Reset to Default"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =2700
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =3014
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackThemeColorIndex =7
                    BackTint =40.0
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
                Begin Label
                    OverlapFlags =85
                    Left =1200
                    Top =3180
                    Width =1920
                    Height =360
                    FontSize =12
                    BorderColor =8355711
                    Name ="lblDeleteUser"
                    Caption ="Delete Account"
                    GridlineColor =10921638
                    LayoutCachedLeft =1200
                    LayoutCachedTop =3180
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =3540
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3240
                    Top =3180
                    Width =2700
                    Height =314
                    FontSize =12
                    TabIndex =6
                    Name ="cmdDeleteAccount"
                    Caption ="Delete Account"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =3180
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =3494
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Gradient =0
                    BackColor =10856415
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

Private Sub cmdConfirm_Click()
    
    MsgBox "Account updated successfully.", vbInformation + vbOKOnly, "Account Updated"
    DoCmd.Close acForm, "fdlgManageAccount"
    Forms(fdlgUserControl).SetFocus
    
End Sub
    
Private Sub cmdDeleteAccount_Click()

    Dim msgResult As Integer

    msgResult = MsgBox("Are you sure you want to delete the user account belonging to " & txtDisplayName.Value & "?", vbCritical + vbYesNo, "Account Deletion")
    
    If msgResult = vbYes Then
        msgResult = MsgBox("You are about to permanently delete a user account! Please confirm deletion of this account.", vbExclamation + vbOKCancel, "Account Deletion")
        If msgResult = vbOK Then
            MsgBox "User account has been deleted. All sessions open under the deleted account have been disconnected.", vbInformation + vbOKOnly, "Account Deletion"
            Form_fdlgUserControl.DeleteAccount txtUsername.Value
            DoCmd.Close acForm, "fdlgManageAccount"
            DoCmd.Close acForm, "fdlgUserControl"
            DoCmd.OpenForm "fdlgUserControl", acNormal, , , acFormReadOnly, acWindowNormal
            Forms(fdlgUserControl).SetFocus
        End If
    End If

End Sub

Private Sub cmdResetPassword_Click()

    Dim msgResult As Integer
    
    msgResult = MsgBox("Reset " & txtDisplayName.Value & "'s password to default?", vbExclamation + vbYesNo, "Reset Password")

    If msgResult = vbYes Then
        changeResult = Form_fdlgUserControl.SetUserPassword(txtUsername.Value, "Thepassword1")
        MsgBox "Password has been reset to the default (Thepassword1).", vbOKOnly, "Password Reset"
    End If

End Sub

Private Sub Form_Load()

    If [ID] = CInt(Form_fdlgUserControl.GetUserID(Form_fdlgUserControl.GetCurrentUser())) Then
        cmdDeleteAccount.Enabled = False
        cmdResetPassword.Enabled = False
        cboAccountType.Enabled = False
    Else
        cmdDeleteAccount.Enabled = True
        cmdResetPassword.Enabled = True
        cboAccountType.Enabled = True
    End If

End Sub
