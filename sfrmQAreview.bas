Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6600
    DatasheetFontHeight =11
    ItemSuffix =17
    Right =16815
    Bottom =9150
    DatasheetGridlinesColor =0
    RecSrcDt = Begin
        0xa61fe7006b09e540
    End
    RecordSource ="qryQAreview"
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
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeShade =50.0
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            BorderShade =90.0
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
            Height =375
            BackColor =15064278
            Name ="FormHeader"
            AlternateBackShade =95.0
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =60
                    Width =375
                    Height =315
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label9"
                    Caption ="Site"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =495
                    LayoutCachedHeight =375
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =660
                    Top =60
                    Width =1410
                    Height =315
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label11"
                    Caption ="Document Name"
                    GridlineColor =10921638
                    LayoutCachedLeft =660
                    LayoutCachedTop =60
                    LayoutCachedWidth =2070
                    LayoutCachedHeight =375
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =4260
                    Top =60
                    Width =1440
                    Height =285
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label14"
                    Caption ="Assigned Person"
                    GridlineColor =10921638
                    LayoutCachedLeft =4260
                    LayoutCachedTop =60
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =345
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =420
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =420
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text3"
                    ControlSource ="Site"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =480
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =480
                    Top =120
                    Width =3720
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =12673797
                    Name ="Text5"
                    ControlSource ="DocumentName"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =480
                    LayoutCachedTop =120
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =360
                    DisplayAsHyperlink =1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =4200
                    Top =120
                    Width =1680
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text8"
                    ControlSource ="Assigned To"
                    GridlineColor =10921638

                    LayoutCachedLeft =4200
                    LayoutCachedTop =120
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =360
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



Private Sub Text5_Click()
DoCmd.OpenForm "EquipForm", , , "DocumentName = '" & Me.DocumentName & "'"
    Forms!EquipForm.ID1.SetFocus

End Sub
