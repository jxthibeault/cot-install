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
    Width =14700
    DatasheetFontHeight =11
    ItemSuffix =115
    Right =15135
    Bottom =11730
    OrderBy ="[tblInstalls].[dtmInstallScheduled] DESC"
    RecSrcDt = Begin
        0xd008711f32eee540
    End
    RecordSource ="qryInstalledInstalls"
    Caption ="Installs Ready to Ship"
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
                    Width =6236
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
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =690
                    LayoutCachedWidth =6296
                    LayoutCachedHeight =1050
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10560
                    Top =690
                    Width =2025
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
                    GridlineColor =10921638
                    LayoutCachedLeft =10560
                    LayoutCachedTop =690
                    LayoutCachedWidth =12585
                    LayoutCachedHeight =1050
                    ColumnStart =3
                    ColumnEnd =3
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6360
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
                    Name ="lblDateReceived"
                    Caption ="Date Received"
                    GridlineColor =10921638
                    LayoutCachedLeft =6360
                    LayoutCachedTop =690
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =1050
                    ColumnStart =1
                    ColumnEnd =1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =120
                    Width =13500
                    Height =480
                    FontSize =16
                    FontWeight =500
                    ForeColor =16777215
                    Name ="lblFormTitle"
                    Caption ="Installs Pending Post-Processing"
                    FontName ="Verdana"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =13560
                    LayoutCachedHeight =600
                    ThemeFontIndex =-1
                    BorderThemeColorIndex =2
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8160
                    Top =690
                    Width =2340
                    Height =360
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblInstallScheduled"
                    Caption ="Installed On"
                    GridlineColor =10921638
                    LayoutCachedLeft =8160
                    LayoutCachedTop =690
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =1050
                    ColumnStart =2
                    ColumnEnd =2
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
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
                    Width =6236
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
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =180
                    LayoutCachedWidth =6296
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10560
                    Top =180
                    Width =2025
                    Height =315
                    FontSize =12
                    TabIndex =3
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSalesRep"
                    ControlSource ="strSalesRep"
                    StatusBarText ="Originating sales rep"
                    GridlineColor =10921638

                    LayoutCachedLeft =10560
                    LayoutCachedTop =180
                    LayoutCachedWidth =12585
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Top =180
                    Width =1740
                    Height =315
                    FontSize =12
                    TabIndex =1
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
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedTop =180
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8160
                    Top =180
                    Width =2340
                    Height =315
                    FontSize =12
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtInstallScheduled"
                    ControlSource ="dtmInstallScheduled"
                    StatusBarText ="Scheduled date of installation"
                    GridlineColor =10921638

                    LayoutCachedLeft =8160
                    LayoutCachedTop =180
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12645
                    Top =180
                    Width =135
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtRecordId"
                    ControlSource ="lngID"
                    StatusBarText ="Primary key - install ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =12645
                    LayoutCachedTop =180
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =495
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =12840
                    Top =180
                    Width =1800
                    Height =300
                    TabIndex =4
                    ForeColor =4210752
                    Name ="cmdOpenInstallDetails"
                    Caption ="View Details"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =12840
                    LayoutCachedTop =180
                    LayoutCachedWidth =14640
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


Private Sub cmdOpenInstallDetails_Click()

    DoCmd.OpenForm "frmInstalledInstall", acNormal, "", "[lngID]=" & txtRecordID, , acNormal

End Sub
