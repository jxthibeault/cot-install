Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =16344
    DatasheetFontHeight =11
    ItemSuffix =119
    Left =-420
    Top =225
    Right =16830
    Bottom =7815
    OrderBy ="dtmInstallScheduled, dtmDateReceived, strCustomer"
    RecSrcDt = Begin
        0xf0ff0ba3d9ebe540
    End
    RecordSource ="qryPendingWorkInstalls"
    Caption ="Installs Pending Work"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
    Moveable =0
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
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
            Height =1080
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
                    Top =690
                    Width =4256
                    Height =360
                    FontSize =12
                    FontWeight =500
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCustomer"
                    Caption ="Customer"
                    GroupTable =1
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    LayoutCachedLeft =60
                    LayoutCachedTop =690
                    LayoutCachedWidth =4316
                    LayoutCachedHeight =1050
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4380
                    Top =690
                    Width =2036
                    Height =360
                    FontSize =12
                    FontWeight =500
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblStatus"
                    Caption ="Status"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4380
                    LayoutCachedTop =690
                    LayoutCachedWidth =6416
                    LayoutCachedHeight =1050
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6480
                    Top =690
                    Width =1616
                    Height =360
                    FontSize =12
                    FontWeight =500
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDateReceived"
                    Caption ="Date Received"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6480
                    LayoutCachedTop =690
                    LayoutCachedWidth =8096
                    LayoutCachedHeight =1050
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =120
                    Width =12120
                    Height =480
                    FontSize =16
                    FontWeight =500
                    ForeColor =16777215
                    Name ="lblFormTitle"
                    Caption ="In-Progress Installs - Pending Work"
                    FontName ="Verdana"
                    GroupTable =2
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =12180
                    LayoutCachedHeight =600
                    LayoutGroup =2
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =2
                End
                Begin EmptyCell
                    Left =12300
                    Top =690
                    Width =135
                    Height =360
                    Name ="EmptyCell99"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =12300
                    LayoutCachedTop =690
                    LayoutCachedWidth =12435
                    LayoutCachedHeight =1050
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8160
                    Top =690
                    Width =1976
                    Height =360
                    FontSize =12
                    FontWeight =500
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDeliveryDate"
                    Caption ="Ship Date"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8160
                    LayoutCachedTop =690
                    LayoutCachedWidth =10136
                    LayoutCachedHeight =1050
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10200
                    Top =690
                    Width =2040
                    Height =360
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblInstallScheduled"
                    Caption ="Install Date"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10200
                    LayoutCachedTop =690
                    LayoutCachedWidth =12240
                    LayoutCachedHeight =1050
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =660
            Name ="secFormDetail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =180
                    Width =4256
                    Height =315
                    ColumnWidth =2895
                    FontSize =12
                    FontWeight =500
                    TabIndex =2
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BackColor =-2147483610
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtCustomer"
                    ControlSource ="strCustomer"
                    StatusBarText ="Customer name as it appears on legal documents"
                    GroupTable =1
                    GridlineColor =10921638
                    HorizontalAnchor =2

                    LayoutCachedLeft =60
                    LayoutCachedTop =180
                    LayoutCachedWidth =4316
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    BackThemeColorIndex =-1
                    GroupTable =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6480
                    Top =180
                    Width =1616
                    Height =315
                    FontSize =12
                    TabIndex =4
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDateReceived"
                    ControlSource ="dtmDateReceived"
                    Format ="Short Date"
                    StatusBarText ="Date installation notice received"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =6480
                    LayoutCachedTop =180
                    LayoutCachedWidth =8096
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4380
                    Top =180
                    Width =2036
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtInstallStatus"
                    ControlSource ="strInstallStatus"
                    StatusBarText ="Install status"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =4380
                    LayoutCachedTop =180
                    LayoutCachedWidth =6416
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12300
                    Top =180
                    Width =135
                    Height =315
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtRecordId"
                    ControlSource ="lngID"
                    StatusBarText ="Primary key - install ID"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =12300
                    LayoutCachedTop =180
                    LayoutCachedWidth =12435
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12900
                    Top =180
                    Width =1613
                    Height =300
                    ForeColor =4210752
                    Name ="cmdOpenInstallDetails"
                    Caption ="View Details"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =12900
                    LayoutCachedTop =180
                    LayoutCachedWidth =14513
                    LayoutCachedHeight =480
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
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =14640
                    Top =180
                    Width =1613
                    Height =299
                    TabIndex =1
                    ForeColor =4210752
                    Name ="cmdPrintTicket"
                    Caption ="Print for Shop"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =14640
                    LayoutCachedTop =180
                    LayoutCachedWidth =16253
                    LayoutCachedHeight =479
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
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8160
                    Top =180
                    Width =1976
                    Height =315
                    FontSize =12
                    TabIndex =5
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtDeliveryDate"
                    ControlSource ="dtmDeliveryDate"
                    StatusBarText ="Originating sales rep"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =8160
                    LayoutCachedTop =180
                    LayoutCachedWidth =10136
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10200
                    Top =180
                    Width =2040
                    Height =315
                    FontSize =12
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtInstallScheduled"
                    ControlSource ="dtmInstallScheduled"
                    StatusBarText ="Scheduled date of installation"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =10200
                    LayoutCachedTop =180
                    LayoutCachedWidth =12240
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin FormFooter
            Height =60
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

Private Sub cmdPrintTicket_Click()

    DoCmd.OpenReport "rptShopTicket", acViewPreview, , "[lngID] = " & txtRecordID.Value, acWindowNormal
    
End Sub

Private Sub cmdOpenInstallDetails_Click()

    DoCmd.OpenForm "frmOpenInstall", acNormal, "", "[lngID]=" & txtRecordID, , acNormal

End Sub


Private Sub Form_Load()

    DoCmd.Maximize

End Sub
