Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8712
    DatasheetFontHeight =11
    ItemSuffix =6
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x5ae018a7d647e340
    End
    RecordSource ="qryEventDates"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0xf0030000f0030000f0030000f003000000000000082200003b01000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =204
            FontSize =11
            FontName ="Calibri"
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontCharSet =204
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin BreakLevel
            ControlSource ="EventDate"
        End
        Begin BreakLevel
            ControlSource ="EventID"
        End
        Begin PageHeader
            Height =315
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    Width =1152
                    Height =315
                    Name ="Label0"
                    Caption ="Date"
                    LayoutCachedWidth =1152
                    LayoutCachedHeight =315
                End
                Begin Label
                    Left =2232
                    Width =3600
                    Height =315
                    Name ="Label1"
                    Caption ="Description"
                    LayoutCachedLeft =2232
                    LayoutCachedWidth =5832
                    LayoutCachedHeight =315
                End
                Begin Label
                    Left =6768
                    Width =1944
                    Height =315
                    Name ="Label2"
                    Caption ="Recuring every"
                    LayoutCachedLeft =6768
                    LayoutCachedWidth =8712
                    LayoutCachedHeight =315
                End
                Begin Label
                    TextAlign =3
                    Left =5832
                    Width =792
                    Height =315
                    Name ="Label3"
                    Caption ="Instance"
                    LayoutCachedLeft =5832
                    LayoutCachedWidth =6624
                    LayoutCachedHeight =315
                End
                Begin Label
                    TextAlign =2
                    Left =1152
                    Width =864
                    Height =315
                    Name ="Label4"
                    Caption ="Event ID"
                    LayoutCachedLeft =1152
                    LayoutCachedWidth =2016
                    LayoutCachedHeight =315
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Width =1152
                    Height =315
                    Name ="EventDate"
                    ControlSource ="EventDate"

                    LayoutCachedWidth =1152
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2232
                    Width =3600
                    Height =315
                    ColumnWidth =3090
                    TabIndex =1
                    Name ="EventDescrip"
                    ControlSource ="DocumentName"
                    StatusBarText ="Description of what this event is."

                    LayoutCachedLeft =2232
                    LayoutCachedWidth =5832
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6768
                    Width =1944
                    Height =315
                    TabIndex =2
                    Name ="txtFreq"
                    ControlSource ="=IIf([RecurCount]=0,Null,Trim([PeriodFreq] & \" \" & [PeriodType]))"

                    LayoutCachedLeft =6768
                    LayoutCachedWidth =8712
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5832
                    Width =792
                    Height =315
                    TabIndex =3
                    Name ="InstanceID"
                    ControlSource ="InstanceID"
                    StatusBarText ="Unique number. (Used for Cartesian product queries.)"

                    LayoutCachedLeft =5832
                    LayoutCachedWidth =6624
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1152
                    Width =864
                    Height =315
                    ColumnWidth =1140
                    TabIndex =4
                    Name ="EventID"
                    ControlSource ="EventID"
                    StatusBarText ="Unique automatically assigned number for this entry."

                    LayoutCachedLeft =1152
                    LayoutCachedWidth =2016
                    LayoutCachedHeight =315
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
