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
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =2
    ScrollBars =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =7140
    DatasheetFontHeight =11
    ItemSuffix =27
    Left =10710
    Top =3225
    Right =17850
    Bottom =8430
    RecSrcDt = Begin
        0x0841aa0ad6efe540
    End
    Caption ="Manage Account"
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
                    Caption ="Create New Account"
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
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtTitle"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =1320
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1635
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =1500
                    Top =3780
                    Width =1920
                    Height =420
                    TabIndex =4
                    ForeColor =4210752
                    Name ="cmdConfirm"
                    Caption ="Confirm"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedTop =3780
                    LayoutCachedWidth =3420
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
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDisplayName"
                    GridlineColor =10921638

                    LayoutCachedLeft =3240
                    LayoutCachedTop =840
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =1155
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3240
                    Top =360
                    Width =2700
                    Height =315
                    FontSize =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtUsername"
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
                    TabIndex =3
                    BorderColor =10921638
                    Name ="cboAccountType"
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
                    TextAlign =2
                    Left =1373
                    Top =2700
                    Width =4395
                    Height =585
                    FontWeight =600
                    BorderColor =8355711
                    Name ="lblResetPassword"
                    Caption ="Account will be created with the default\015\012password."
                    GridlineColor =10921638
                    LayoutCachedLeft =1373
                    LayoutCachedTop =2700
                    LayoutCachedWidth =5768
                    LayoutCachedHeight =3285
                    ForeTint =100.0
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =3780
                    Top =3780
                    Width =1920
                    Height =420
                    TabIndex =5
                    ForeColor =4210752
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3780
                    LayoutCachedTop =3780
                    LayoutCachedWidth =5700
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

    Dim intMessageResult As Integer
    
    intMessageResult = MsgBox("Cancel new account creation?", vbExclamation + vbYesNo, "User Account Creation")
    If intMessageResult = vbYes Then
        DoCmd.Close acForm, "fdlgCreateAccount"
        Forms(fdlgUserControl).SetFocus
    End If

End Sub

Private Sub cmdConfirm_Click()
    
    If Form_fdlgUserControl.UserExists(txtUsername.Value) Then
        MsgBox "Username already exists, please select a different username.", vbExclamation + vbOKOnly, "User Account Creation"
        txtUsername.SetFocus
    ElseIf InStr(txtUsername.Value, " ") > 0 Then
        MsgBox "Username cannot contain spaces!", vbExclamation + vbOKOnly, "User Account Creation"
        txtUsername.SetFocus
    Else
        Form_fdlgUserControl.CreateUser txtUsername.Value, txtDisplayName.Value, _
                txtTitle.Value, cboAccountType.Value
        MsgBox "Account created successfully!", vbInformation + vbOKOnly, "Account Created"
        
        DoCmd.Close acForm, "fdlgCreateAccount"
        DoCmd.Close acForm, "fdlgUserControl"
        DoCmd.OpenForm "fdlgUserControl", acNormal, , , acFormReadOnly, acWindowNormal
        Forms(fdlgUserControl).SetFocus
    End If
    
End Sub
