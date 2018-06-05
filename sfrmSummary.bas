Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12060
    DatasheetFontHeight =11
    ItemSuffix =18
    Left =945
    Top =2805
    Right =11775
    Bottom =9195
    DatasheetGridlinesColor =0
    RecSrcDt = Begin
        0xe45f2d136b09e540
    End
    RecordSource ="SELECT MERM.EventID, MERM.Site, MERM.CommonName, MERM.DocumentName, MERM.Asset, "
        "MERM.AQMD_ID, MERM.[Assigned To], [Contacts Extended].ID FROM MERM INNER JOIN [C"
        "ontacts Extended] ON MERM.[Assigned To] = [Contacts Extended].[Contact Name] ORD"
        "ER BY MERM.[Assigned To] DESC;"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            BorderTint =50.0
            ForeTint =50.0
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
            BorderShade =65.0
            ForeTint =75.0
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =540
            BackColor =13017476
            Name ="FormHeader"
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Width =270
                    Height =315
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label2"
                    Caption ="ID"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedWidth =390
                    LayoutCachedHeight =315
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1260
                    Width =585
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label10"
                    Caption ="Group"
                    GridlineColor =10921638
                    LayoutCachedLeft =1260
                    LayoutCachedWidth =1845
                    LayoutCachedHeight =285
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =420
                    Width =570
                    Height =525
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label9"
                    Caption ="Site\015\012 í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =420
                    LayoutCachedWidth =990
                    LayoutCachedHeight =525
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5040
                    Width =1485
                    Height =525
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label11"
                    Caption ="Document Name\015\012 í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =5040
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =525
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =8700
                    Width =615
                    Height =465
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label12"
                    Caption ="AQMD\015\012 í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =8700
                    LayoutCachedWidth =9315
                    LayoutCachedHeight =465
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =9480
                    Width =675
                    Height =465
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label13"
                    Caption ="Asset #\015\012 í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =9480
                    LayoutCachedWidth =10155
                    LayoutCachedHeight =465
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10320
                    Width =1485
                    Height =465
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label14"
                    Caption ="Assigned Person\015\012 í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =10320
                    LayoutCachedWidth =11805
                    LayoutCachedHeight =465
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2220
                    Width =1485
                    Height =525
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label17"
                    Caption ="Common Name\015\012 í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =2220
                    LayoutCachedWidth =3705
                    LayoutCachedHeight =525
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =360
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Top =60
                    Width =540
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text0"
                    ControlSource ="EventID"
                    GridlineColor =10921638

                    LayoutCachedTop =60
                    LayoutCachedWidth =540
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =480
                    Top =60
                    Width =540
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text3"
                    ControlSource ="Site"
                    GridlineColor =10921638

                    LayoutCachedLeft =480
                    LayoutCachedTop =60
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =900
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text4"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =900
                    LayoutCachedTop =60
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4200
                    Top =60
                    Width =4440
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =12673797
                    Name ="Text5"
                    ControlSource ="DocumentName"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4200
                    LayoutCachedTop =60
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =300
                    DisplayAsHyperlink =1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =8580
                    Top =60
                    Width =900
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text6"
                    ControlSource ="AQMD_ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =8580
                    LayoutCachedTop =60
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =127
                    IMESentenceMode =3
                    Left =9480
                    Top =60
                    Width =840
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text7"
                    ControlSource ="Asset"
                    GridlineColor =10921638

                    LayoutCachedLeft =9480
                    LayoutCachedTop =60
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =119
                    IMESentenceMode =3
                    Left =10320
                    Top =60
                    Width =1740
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text8"
                    ControlSource ="Assigned To"
                    GridlineColor =10921638

                    LayoutCachedLeft =10320
                    LayoutCachedTop =60
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =119
                    IMESentenceMode =3
                    Left =1800
                    Top =60
                    Width =2400
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text15"
                    ControlSource ="CommonName"
                    GridlineColor =10921638

                    LayoutCachedLeft =1800
                    LayoutCachedTop =60
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =300
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackShade =95.0
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Label11_Click()

Me.OrderByOn = True

If Me.OrderBy = "[DocumentName] DESC" Then
    Me.OrderBy = "[DocumentName] ASC"
Else
    Me.OrderBy = "[DocumentName] DESC"
End If

Me.Refresh
End Sub


Private Sub Label12_Click()

Me.OrderByOn = True

If Me.OrderBy = "[AQMD_ID] DESC" Then
    Me.OrderBy = "[AQMD_ID] ASC"
Else
    Me.OrderBy = "[AQMD_ID] DESC"
End If

Me.Refresh

End Sub

Private Sub Label13_Click()
If Me.OrderBy = "[Asset] DESC" Then
    Me.OrderBy = "[Asset] ASC"
Else
    Me.OrderBy = "[Asset] DESC"
End If

Me.Refresh
End Sub

Private Sub Label14_Click()

Me.OrderByOn = True

If Me.OrderBy = "[Assigned To] DESC" Then
    Me.OrderBy = "[Assigned To] ASC"
Else
    Me.OrderBy = "[Assigned To] DESC"
End If

Me.Refresh

End Sub

Private Sub Label17_Click()
Me.OrderByOn = True

If Me.OrderBy = "[CommonName] DESC" Then
    Me.OrderBy = "[CommonName] ASC"
Else
    Me.OrderBy = "[CommonName] DESC"
End If

Me.Refresh
End Sub

Private Sub Label9_Click()

Me.OrderByOn = True

If Me.OrderBy = "[Site] DESC" Then
    Me.OrderBy = "[Site] ASC"
Else
    Me.OrderBy = "[Site] DESC"
End If

Me.Refresh

End Sub

Private Sub Text5_Click()
DoCmd.OpenForm "EquipForm", , , "DocumentName = '" & Me.DocumentName & "'"
    Forms!EquipForm.ID1.SetFocus
    'Forms!EquipForm.ID1.SelStart = 0
    'Forms!EquipForm.ID1.SelLength = 0
End Sub
