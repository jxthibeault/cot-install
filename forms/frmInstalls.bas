Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10845
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =345
    Top =1605
    Right =10155
    Bottom =9315
    RecSrcDt = Begin
        0x76dfbbf19ff0e540
    End
    Caption ="Installs"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
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
        Begin Section
            CanGrow = NotDefault
            Height =7560
            BackColor =-2147483644
            Name ="secDetail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin ListBox
                    ColumnHeads = NotDefault
                    SpecialEffect =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =9
                    Left =120
                    Top =2460
                    Width =10620
                    Height =4965
                    FontSize =8
                    BackColor =-2147483644
                    ForeColor =-2147483617
                    BorderColor =-2147483638
                    Name ="lstInstallsList"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [tblInstalls].[lngID] AS [No], [tblInstalls].[strInstallStatus] AS Status"
                        ", [tblInstalls].[strCustomer] AS Customer, [tblInstalls].[strAddressCity] AS Cit"
                        "y, [tblInstalls].[strAddressState] AS State, [tblInstalls].[strSalesRep] AS [Sal"
                        "es Rep], [tblInstalls].[dtmDateReceived] AS Received, [tblInstalls].[dtmDelivery"
                        "Date] AS [Ship Date], [tblInstalls].[dtmInstallScheduled] AS [Install Date] FROM"
                        " tblInstalls ORDER BY [strInstallStatus] DESC; "
                    ColumnWidths ="0;1800;5760;1440;720;1440;1080;2160;2160"
                    StatusBarText ="Installs"
                    FontName ="Arial"
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    VerticalAnchor =2
                    AllowValueListEdits =0

                    ShowOnlyRowSourceValues =255
                    LayoutCachedLeft =120
                    LayoutCachedTop =2460
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =7425
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
    End
End
