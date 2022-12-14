Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14100
    DatasheetFontHeight =11
    ItemSuffix =103
    Top =825
    Right =14265
    Bottom =8415
    OrderBy ="strCustomer"
    RecSrcDt = Begin
        0x2579385d79ece540
    End
    RecordSource ="qryOpenInstalls"
    Caption ="Modify an Order"
    DatasheetFontName ="Calibri"
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
                    Left =10800
                    Top =690
                    Width =1740
                    Height =360
                    FontSize =12
                    FontWeight =500
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblSalesRep"
                    Caption ="Sales Rep"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10800
                    LayoutCachedTop =690
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =1050
                    ColumnStart =4
                    ColumnEnd =4
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
                    Width =1736
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
                    LayoutCachedWidth =8216
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
                    Caption ="Modify an Open Installation Order"
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
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8280
                    Top =690
                    Width =2460
                    Height =360
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblInstallScheduled"
                    Caption ="Install Scheduled"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8280
                    LayoutCachedTop =690
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =1050
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GroupTable =1
                End
                Begin EmptyCell
                    Left =12600
                    Top =690
                    Width =135
                    Height =360
                    Name ="EmptyCell99"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =12600
                    LayoutCachedTop =690
                    LayoutCachedWidth =12735
                    LayoutCachedHeight =1050
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =645
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
                    Left =10800
                    Top =180
                    Width =1740
                    Height =315
                    FontSize =12
                    TabIndex =4
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSalesRep"
                    ControlSource ="strSalesRep"
                    StatusBarText ="Originating sales rep"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =10800
                    LayoutCachedTop =180
                    LayoutCachedWidth =12540
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
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
                    Left =6480
                    Top =180
                    Width =1736
                    Height =315
                    FontSize =12
                    TabIndex =2
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
                    LayoutCachedWidth =8216
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
                    TabIndex =1
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
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8280
                    Top =180
                    Width =2460
                    Height =315
                    FontSize =12
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtInstallScheduled"
                    ControlSource ="dtmInstallScheduled"
                    StatusBarText ="Scheduled date of installation"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =8280
                    LayoutCachedTop =180
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =12600
                    Top =180
                    Width =135
                    Height =315
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtRecordId"
                    ControlSource ="lngID"
                    StatusBarText ="Primary key - install ID"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =12600
                    LayoutCachedTop =180
                    LayoutCachedWidth =12735
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =12600
                    Top =180
                    Height =300
                    TabIndex =5
                    ForeColor =4210752
                    Name ="cmdModifyOrder"
                    Caption ="Modify Order"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =12600
                    LayoutCachedTop =180
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =480
                    Gradient =0
                    BackColor =15592953
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

Private Sub cmdModifyOrder_Click()

        DoCmd.OpenForm "frmEditInstall", acNormal, "", "[lngID]=" & txtRecordID.Value, acFormEdit, acWindowNormal

    
End Sub
