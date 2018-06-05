Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =3
    PictureSizeMode =3
    DatasheetGridlinesBehavior =3
    PictureType =2
    GridX =24
    GridY =24
    Width =6060
    DatasheetFontHeight =11
    ItemSuffix =29
    Left =30315
    Top =3930
    Right =-29161
    Bottom =8175
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x4a75d9380a06e540
    End
    Caption ="Automated Monthly Emails"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    PictureSizeMode =3
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
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            GridlineColor =16777215
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
            BorderColor =16777215
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
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =4260
            BackColor =14602694
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Image
                    PictureType =2
                    Left =3720
                    Top =3540
                    Width =2340
                    Height =720
                    BorderColor =10921638
                    Name ="Image18"
                    Picture ="BestMERM"
                    GridlineColor =10921638

                    LayoutCachedLeft =3720
                    LayoutCachedTop =3540
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =4260
                    TabIndex =1
                End
                Begin Tab
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =120
                    Top =60
                    Width =5880
                    Height =3435
                    Name ="TabCtl19"
                    FontName ="Calibri Light"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =6000
                    LayoutCachedHeight =3495
                    BackColor =14277081
                    BorderColor =11573124
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    ForeColor =4210752
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =195
                            Top =540
                            Width =5730
                            Height =2880
                            BorderColor =10921638
                            Name ="Page20"
                            Caption ="Monthy"
                            GridlineColor =10921638
                            LayoutCachedLeft =195
                            LayoutCachedTop =540
                            LayoutCachedWidth =5925
                            LayoutCachedHeight =3420
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    Left =275
                                    Top =620
                                    Width =5535
                                    Height =300
                                    FontWeight =700
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Label0"
                                    Caption ="Please select the month you would like to send emails for:"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =275
                                    LayoutCachedTop =620
                                    LayoutCachedWidth =5810
                                    LayoutCachedHeight =920
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =4680
                                    Top =2640
                                    Width =840
                                    Height =420
                                    ForeColor =4210752
                                    Name ="December"
                                    Caption ="DEC"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =4680
                                    LayoutCachedTop =2640
                                    LayoutCachedWidth =5520
                                    LayoutCachedHeight =3060
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =3360
                                    Top =2640
                                    Width =840
                                    Height =420
                                    TabIndex =1
                                    ForeColor =4210752
                                    Name ="November"
                                    Caption ="NOV"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =3360
                                    LayoutCachedTop =2640
                                    LayoutCachedWidth =4200
                                    LayoutCachedHeight =3060
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =695
                                    Top =1100
                                    Width =840
                                    Height =420
                                    TabIndex =2
                                    ForeColor =4210752
                                    Name ="January"
                                    Caption ="JAN"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =695
                                    LayoutCachedTop =1100
                                    LayoutCachedWidth =1535
                                    LayoutCachedHeight =1520
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =2040
                                    Top =2640
                                    Width =840
                                    Height =420
                                    TabIndex =3
                                    ForeColor =4210752
                                    Name ="October"
                                    Caption ="OCT"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =2040
                                    LayoutCachedTop =2640
                                    LayoutCachedWidth =2880
                                    LayoutCachedHeight =3060
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =720
                                    Top =1860
                                    Width =840
                                    Height =420
                                    TabIndex =4
                                    ForeColor =4210752
                                    Name ="May"
                                    Caption ="MAY"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =720
                                    LayoutCachedTop =1860
                                    LayoutCachedWidth =1560
                                    LayoutCachedHeight =2280
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =2015
                                    Top =1100
                                    Width =840
                                    Height =420
                                    TabIndex =5
                                    ForeColor =4210752
                                    Name ="February"
                                    Caption ="FEB"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =2015
                                    LayoutCachedTop =1100
                                    LayoutCachedWidth =2855
                                    LayoutCachedHeight =1520
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =3335
                                    Top =1100
                                    Width =840
                                    Height =420
                                    TabIndex =6
                                    ForeColor =4210752
                                    Name ="March"
                                    Caption ="MAR"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =3335
                                    LayoutCachedTop =1100
                                    LayoutCachedWidth =4175
                                    LayoutCachedHeight =1520
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =4655
                                    Top =1100
                                    Width =840
                                    Height =420
                                    TabIndex =7
                                    ForeColor =4210752
                                    Name ="April"
                                    Caption ="APR"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =4655
                                    LayoutCachedTop =1100
                                    LayoutCachedWidth =5495
                                    LayoutCachedHeight =1520
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =2040
                                    Top =1860
                                    Width =840
                                    Height =420
                                    TabIndex =8
                                    ForeColor =4210752
                                    Name ="June"
                                    Caption ="JUN"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =2040
                                    LayoutCachedTop =1860
                                    LayoutCachedWidth =2880
                                    LayoutCachedHeight =2280
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =3360
                                    Top =1860
                                    Width =840
                                    Height =420
                                    TabIndex =9
                                    ForeColor =4210752
                                    Name ="July"
                                    Caption ="JUL"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =3360
                                    LayoutCachedTop =1860
                                    LayoutCachedWidth =4200
                                    LayoutCachedHeight =2280
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =4680
                                    Top =1860
                                    Width =840
                                    Height =420
                                    TabIndex =10
                                    ForeColor =4210752
                                    Name ="August"
                                    Caption ="AUG"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =4680
                                    LayoutCachedTop =1860
                                    LayoutCachedWidth =5520
                                    LayoutCachedHeight =2280
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =720
                                    Top =2640
                                    Width =840
                                    Height =420
                                    TabIndex =11
                                    ForeColor =4210752
                                    Name ="September"
                                    Caption ="SEP"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =720
                                    LayoutCachedTop =2640
                                    LayoutCachedWidth =1560
                                    LayoutCachedHeight =3060
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
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
                        Begin Page
                            OverlapFlags =215
                            Left =195
                            Top =540
                            Width =5730
                            Height =2880
                            BorderColor =10921638
                            Name ="Page21"
                            Caption ="Groups"
                            GridlineColor =10921638
                            LayoutCachedLeft =195
                            LayoutCachedTop =540
                            LayoutCachedWidth =5925
                            LayoutCachedHeight =3420
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    Left =270
                                    Top =615
                                    Width =5610
                                    Height =300
                                    FontWeight =700
                                    BorderColor =8355711
                                    ForeColor =8355711
                                    Name ="Label23"
                                    Caption ="Please select the group of people you would like to contact:"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =270
                                    LayoutCachedTop =615
                                    LayoutCachedWidth =5880
                                    LayoutCachedHeight =915
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =1380
                                    Top =1020
                                    Width =3360
                                    Height =405
                                    ForeColor =4210752
                                    Name ="Command24"
                                    Caption ="Title V Leads"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =1380
                                    LayoutCachedTop =1020
                                    LayoutCachedWidth =4740
                                    LayoutCachedHeight =1425
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =1380
                                    Top =1500
                                    Width =3360
                                    Height =405
                                    TabIndex =1
                                    ForeColor =4210752
                                    Name ="Command25"
                                    Caption ="Environmental"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =1380
                                    LayoutCachedTop =1500
                                    LayoutCachedWidth =4740
                                    LayoutCachedHeight =1905
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =1380
                                    Top =1980
                                    Width =3360
                                    Height =405
                                    TabIndex =2
                                    ForeColor =4210752
                                    Name ="Command26"
                                    Caption ="Facilites"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =1380
                                    LayoutCachedTop =1980
                                    LayoutCachedWidth =4740
                                    LayoutCachedHeight =2385
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =1380
                                    Top =2940
                                    Width =3360
                                    Height =405
                                    TabIndex =3
                                    ForeColor =4210752
                                    Name ="Command27"
                                    Caption ="All"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =1380
                                    LayoutCachedTop =2940
                                    LayoutCachedWidth =4740
                                    LayoutCachedHeight =3345
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =1380
                                    Top =2460
                                    Width =3360
                                    Height =405
                                    ForeColor =4210752
                                    Name ="Command28"
                                    Caption ="Health and Safety"
                                    OnClick ="[Event Procedure]"
                                    GridlineColor =10921638

                                    LayoutCachedLeft =1380
                                    LayoutCachedTop =2460
                                    LayoutCachedWidth =4740
                                    LayoutCachedHeight =2865
                                    BackColor =15123357
                                    BorderColor =15123357
                                    HoverColor =15652797
                                    PressedColor =11957550
                                    HoverForeColor =4210752
                                    PressedForeColor =4210752
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command24_Click()
'Email Title V Leads
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery2")
    qdf.sql = "Select * From [qryGroupEmail] WHERE [Group]='Title V'"
DoCmd.Hourglass True
    intCnt = DCount("[Contact Name]", "[tempQuery2]")
        If intCnt = 0 Then
    MsgBox "There are no Title V Group members.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to create and email message to the Title V leads." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendGroupEmail
DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub

Private Sub Command25_Click()
'Email Env
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery2")
    qdf.sql = "Select * From [qryGroupEmail] WHERE [Group]='Env'"
DoCmd.Hourglass True
    intCnt = DCount("[Contact Name]", "[tempQuery2]")
        If intCnt = 0 Then
    MsgBox "There are no Env Group members.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to create and email message to the Environmental MERM contacts." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendGroupEmail
DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub Command26_Click()
'Email Facilities
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery2")
    qdf.sql = "Select * From [qryGroupEmail] WHERE [Group]='Facilities'"
DoCmd.Hourglass True
    intCnt = DCount("[Contact Name]", "[tempQuery2]")
        If intCnt = 0 Then
    MsgBox "There are no Facilities Group members.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to create and email message to the Facilities MERM contacts." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendGroupEmail
DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub Command27_Click()
'Email All MERM Contacts
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery2")
    qdf.sql = "Select * From [qryGroupEmail]"
DoCmd.Hourglass True
    intCnt = DCount("[Contact Name]", "[tempQuery2]")
        If intCnt = 0 Then
    MsgBox "There are no Group members meeting this criteria.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to create and email message to all the MERM contacts." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendGroupEmail
DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub


Private Sub Command28_Click()
'Email Health & Safety
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery2")
    qdf.sql = "Select * From [qryGroupEmail] WHERE [Group]='H&S'"
DoCmd.Hourglass True
    intCnt = DCount("[Contact Name]", "[tempQuery2]")
        If intCnt = 0 Then
    MsgBox "There are no Health & Safety Group members.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to create and email message to the Health & Safety MERM contacts." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendGroupEmail
DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub


Private Sub January_Click()
Dim inCnt As Integer
Dim recCnt As Integer
Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='January'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of January." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub February_Click()
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='February'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of February." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub March_Click()
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='March'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of March." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub April_Click()
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='April'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of April." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub May_Click()
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='May'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of May." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub June_Click()
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='June'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of June." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub July_Click()
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='July'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of July." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub August_Click()
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='August'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of August." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub September_Click()
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='September'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of September." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub October_Click()
Dim inCnt As Integer
Dim LResponse As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='October'"

DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of October." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub

Private Sub November_Click()
Dim inCnt As Integer
Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
        Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='November'"

DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of November." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
Private Sub December_Click()
Dim inCnt As Integer

Dim Db As DAO.Database
    Set Db = CurrentDb
    Dim qdf As DAO.QueryDef
    
    Set qdf = Db.QueryDefs("tempQuery")
    qdf.sql = "Select * From [qrymonthlyEmail] WHERE [Month]='December'"
DoCmd.Hourglass True
    intCnt = DCount("[Month]", "[tempQuery]")
        If intCnt = 0 Then
    MsgBox "There are no records outstanding for this month.", 64, "Report Status"
DoCmd.Hourglass False
Exit Sub
 End If
        If intCnt > 0 Then
        LResponse = MsgBox("NOTICE:" & vbCrLf & "You are about to send emails for the month of December." & vbCrLf & "Are you sure you want to continue?", vbYesNo, "Send Monthly Emails?")
        If LResponse = vbYes Then
        Call SendMonthlyEmail
   DoCmd.Hourglass False
        Else
         MsgBox "Action canceled by user.", 64, "Canceled Sending Emails"
DoCmd.Hourglass False
        Exit Sub
        End If
        End If
End Sub
