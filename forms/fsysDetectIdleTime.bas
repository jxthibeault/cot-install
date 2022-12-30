Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    DatasheetFontHeight =11
    Left =1740
    Top =1410
    Right =10605
    Bottom =4350
    TimerInterval =1000
    RecSrcDt = Begin
        0xebeaa2e0d7eee540
    End
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
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
        Begin Section
            Height =7560
            Name ="secDetail"
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

Sub Form_Timer()

    ' IDLEMINUTES determines how much idle time to wait for before
    ' running the IdleTimeDetected subroutine.
    Const IDLEMINUTES = 15

    Static PrevControlName As String
    Static PrevFormName As String
    Static ExpiredTime

    Dim ActiveFormName As String
    Dim ActiveControlName As String
    Dim ExpiredMinutes

    On Error Resume Next

    ' Get the active form and control name.

    ActiveFormName = Screen.ActiveForm.Name
    If Err Then
        ActiveFormName = "No Active Form"
        Err = 0
    End If

    ActiveControlName = Screen.ActiveControl.Name
    If Err Then
        ActiveControlName = "No Active Control"
        Err = 0
    End If

    ' Record the current active names and reset ExpiredTime if:
    '    1. They have not been recorded yet (code is running
    '       for the first time).
    '    2. The previous names are different than the current ones
    '       (the user has done something different during the timer
    '        interval).
    If (PrevControlName = "") Or (PrevFormName = "") _
        Or (ActiveFormName <> PrevFormName) _
        Or (ActiveControlName <> PrevControlName) Then
        PrevControlName = ActiveControlName
        PrevFormName = ActiveFormName
        ExpiredTime = 0
    Else
        ' ...otherwise the user was idle during the time interval, so
        ' increment the total expired time.
        ExpiredTime = ExpiredTime + Me.TimerInterval
    End If

    ' Does the total expired time exceed the IDLEMINUTES?
    ExpiredMinutes = (ExpiredTime / 1000) / 60
    If ExpiredMinutes >= IDLEMINUTES Then
        ' ...if so, then reset the expired time to zero...
        ExpiredTime = 0
        ' ...and call the IdleTimeDetected subroutine.
        IdleTimeDetected ExpiredMinutes
    End If
    
End Sub

Sub IdleTimeDetected(ExpiredMinutes)

    Dim strSQL As String
    Dim intLoopCount As Integer

    ' Disable warnings, as DoCmd.RunSQL asks user for confirmation before executing
    DoCmd.SetWarnings False
    
    ' On closing the form, remove this connection from connections list
    strSQL = "Delete * From [tblConnections] WHERE [strHostname] = '" & GetHostname() & "'"
    DoCmd.RunSQL strSQL
    
    ' Re-enable warnings (in effect, return to default setting)
    DoCmd.SetWarnings True
    
    ' Close all forms except login form
    For intLoopCount = (Forms.Count - 1) To 0 Step -1
        DoCmd.Close acForm, Forms(intLoopCount).Name
        Next intLoopCount
    
    ' Close all reports
    For intLoopCount = (Reports.Count - 1) To 0 Step -1
        DoCmd.Close acReport, Reports(intLoopCount).Name
        Next intLoopCount
    
    ' Re-open login form
    DoCmd.OpenForm "fdlgLogIn", acNormal, , , , acNormal
    
    MsgBox "User has been logged out due to inactivity.", vbOKOnly, "Logged Out"
    
End Sub
