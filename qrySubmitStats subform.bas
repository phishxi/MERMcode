Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6630
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =1185
    Top =3750
    Right =8100
    Bottom =4980
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x517b24e0d218e540
    End
    RecordSource ="SELECT [qrySubmitStats].[Contact Name], [qrySubmitStats].[DateComplete] FROM [qr"
        "ySubmitStats]"
    Caption ="qrySubmitStats subform"
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
            ForeTint =50.0
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
            Height =480
            BackColor =15064278
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin TextBox
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =120
                    Top =60
                    Width =2160
                    Height =315
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text4"
                    ControlSource ="Contact Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =60
                            Width =585
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label5"
                            Caption ="Text4"
                            GridlineColor =10921638
                            LayoutCachedTop =60
                            LayoutCachedWidth =585
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2400
                    Top =60
                    Width =1320
                    Height =315
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text7"
                    ControlSource ="=DCount(\"SubmitID\",\"tblReview\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2400
                    LayoutCachedTop =60
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =600
                            Top =60
                            Width =585
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label8"
                            Caption ="Text7"
                            GridlineColor =10921638
                            LayoutCachedLeft =600
                            LayoutCachedTop =60
                            LayoutCachedWidth =1185
                            LayoutCachedHeight =375
                        End
                    End
                End
            End
        End
        Begin Section
            Height =1230
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =360
                    Width =3660
                    Height =330
                    ColumnWidth =3000
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Contact Name"
                    ControlSource ="Contact Name"
                    EventProcPrefix ="Contact_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =360
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =690
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Contact Name_Label"
                            Caption ="Contact Name"
                            EventProcPrefix ="Contact_Name_Label"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =690
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =780
                    Width =1620
                    Height =330
                    ColumnWidth =1620
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="DateComplete"
                    ControlSource ="DateComplete"
                    GridlineColor =10921638
                    ShowDatePicker =0

                    LayoutCachedLeft =2910
                    LayoutCachedTop =780
                    LayoutCachedWidth =4530
                    LayoutCachedHeight =1110
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =780
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="DateComplete_Label"
                            Caption ="DateComplete"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1110
                        End
                    End
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
