Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7680
    DatasheetFontHeight =11
    ItemSuffix =12
    Left =29490
    Top =1065
    Right =-28636
    Bottom =6315
    DatasheetGridlinesColor =15132391
    OrderBy ="[Last Name]"
    RecSrcDt = Begin
        0x0802eead8a14e540
    End
    RecordSource ="Contacts"
    Caption ="Manage Users"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
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
        Begin FormHeader
            Height =915
            BackColor =5066944
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =660
                    Top =600
                    Width =1140
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="Last Name_Label"
                    Caption ="Last Name"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Last_Name_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =660
                    LayoutCachedTop =600
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =915
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =2580
                    Top =600
                    Width =1080
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="First Name_Label"
                    Caption ="First Name"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="First_Name_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2580
                    LayoutCachedTop =600
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =915
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =4080
                    Top =600
                    Width =1560
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="rcvEmail_Label"
                    Caption ="Receive Emails"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4080
                    LayoutCachedTop =600
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =900
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =5820
                    Top =600
                    Width =1365
                    Height =315
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="Admin_Label"
                    Caption ="Administrator"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5820
                    LayoutCachedTop =600
                    LayoutCachedWidth =7185
                    LayoutCachedHeight =915
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =87
                    Left =60
                    Top =60
                    Width =2955
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="Label10"
                    Caption ="User Permissions"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3015
                    LayoutCachedHeight =600
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Image
                    PictureType =2
                    Left =5160
                    Top =60
                    Width =2280
                    Height =600
                    BorderColor =10921638
                    Name ="Image11"
                    Picture ="MERM Logo"
                    GridlineColor =10921638

                    LayoutCachedLeft =5160
                    LayoutCachedTop =60
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =660
                End
            End
        End
        Begin Section
            Height =360
            Name ="Detail"
            AlternateBackColor =15592953
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Width =660
                    Height =315
                    ColumnWidth =1440
                    BorderColor =10921638
                    ForeColor =12673797
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Condition ="[Form].[Dirty]"
                            Action ="RunCommand"
                            Argument ="97"
                        End
                        Begin
                            Condition ="[MacroError].[Number]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Condition ="..."
                            Action ="StopMacro"
                        End
                        Begin
                            Action ="OnError"
                            Argument ="2"
                        End
                        Begin
                            Action ="OpenForm"
                            Argument ="Contact Details"
                            Argument ="0"
                            Argument =""
                            Argument ="=\"[ID]=\" & Nz([ID],0)"
                            Argument ="-1"
                            Argument ="3"
                        End
                        Begin
                            Condition ="Not IsNull([ID])"
                            Action ="SetTempVar"
                            Argument ="CurrentID"
                            Argument ="[ID]"
                        End
                        Begin
                            Condition ="IsNull([ID])"
                            Action ="SetTempVar"
                            Argument ="CurrentID"
                            Argument ="Nz(DMax(\"[ID]\",[Form].[RecordSource]),0)"
                        End
                        Begin
                            Action ="Requery"
                        End
                        Begin
                            Action ="SearchForRecord"
                            Argument ="-1"
                            Argument =""
                            Argument ="2"
                            Argument ="=\"[ID]=\" & [TempVars]![CurrentID]"
                        End
                        Begin
                            Action ="RemoveTempVar"
                            Argument ="CurrentID"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"ID\" Event=\"OnClick\" xmlns=\"http://schemas.microsoft.com/"
                                "office/accessservices/2009/11/application\"><Statements><Action Name=\"OnError\""
                                "/><ConditionalBlock><If><Conditio"
                        End
                        Begin
                            Comment ="_AXL:n>[Form].[Dirty]</Condition><Statements><Action Name=\"SaveRecord\"/></Stat"
                                "ements></If></ConditionalBlock><ConditionalBlock><If><Condition>[MacroError].[Nu"
                                "mber]&lt;&gt;0</Condition><Statements><Action Name=\"MessageBox\"><Argument Name"
                                "=\"Message\">=[Macr"
                        End
                        Begin
                            Comment ="_AXL:oError].[Description]</Argument></Action><Action Name=\"StopMacro\"/></Stat"
                                "ements></If></ConditionalBlock><Action Name=\"OnError\"><Argument Name=\"Goto\">"
                                "Fail</Argument></Action><Action Name=\"OpenForm\"><Argument Name=\"FormName\">Co"
                                "ntact Details</Argument"
                        End
                        Begin
                            Comment ="_AXL:><Argument Name=\"WhereCondition\">=\"[ID]=\" &amp; Nz([ID],0)</Argument><A"
                                "rgument Name=\"WindowMode\">Dialog</Argument></Action><ConditionalBlock><If><Con"
                                "dition>Not IsNull([ID])</Condition><Statements><Action Name=\"SetTempVar\"><Argu"
                                "ment Name=\"Name\">Curr"
                        End
                        Begin
                            Comment ="_AXL:entID</Argument><Argument Name=\"Expression\">[ID]</Argument></Action></Sta"
                                "tements></If></ConditionalBlock><ConditionalBlock><If><Condition>IsNull([ID])</C"
                                "ondition><Statements><Action Name=\"SetTempVar\"><Argument Name=\"Name\">Current"
                                "ID</Argument><Argum"
                        End
                        Begin
                            Comment ="_AXL:ent Name=\"Expression\">Nz(DMax(\"[ID]\",[Form].[RecordSource]),0)</Argumen"
                                "t></Action></Statements></If></ConditionalBlock><Action Name=\"Requery\"/><Actio"
                                "n Name=\"SearchForRecord\"><Argument Name=\"WhereCondition\">=\"[ID]=\" &amp; [T"
                                "empVars]![CurrentID]</Arg"
                        End
                        Begin
                            Comment ="_AXL:ument></Action><Action Name=\"RemoveTempVar\"><Argument Name=\"Name\">Curre"
                                "ntID</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedWidth =660
                    LayoutCachedHeight =315
                    DisplayAsHyperlink =1
                    ForeThemeColorIndex =10
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    Left =660
                    Width =1800
                    Height =330
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Last Name"
                    ControlSource ="Last Name"
                    EventProcPrefix ="Last_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =660
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    Left =2580
                    Width =2220
                    Height =330
                    ColumnWidth =3000
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="First Name"
                    ControlSource ="First Name"
                    EventProcPrefix ="First_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =330
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4860
                    Top =60
                    TabIndex =3
                    BorderColor =10921638
                    Name ="rcvEmail"
                    ControlSource ="rcvEmail"
                    GridlineColor =10921638

                    LayoutCachedLeft =4860
                    LayoutCachedTop =60
                    LayoutCachedWidth =5120
                    LayoutCachedHeight =300
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6360
                    Top =60
                    TabIndex =4
                    BorderColor =10921638
                    Name ="Admin"
                    ControlSource ="Admin"
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedTop =60
                    LayoutCachedWidth =6620
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
