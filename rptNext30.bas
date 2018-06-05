Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11490
    DatasheetFontHeight =11
    ItemSuffix =24
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x6f3378f9e81be540
    End
    RecordSource ="Next30"
    Caption ="CurrentMonthReport"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000e22c00004a01000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetBackColor12 =-2147483643
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ForeTint =60.0
            GridlineShade =65.0
        End
        Begin Line
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
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
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            BorderThemeColorIndex =3
            BorderShade =90.0
            ThemeFontIndex =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BorderShade =90.0
            ForeShade =50.0
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Assigned To"
        End
        Begin BreakLevel
            ControlSource ="EventDate"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =780
            BackColor =15064278
            Name ="ReportHeader"
            AlternateBackShade =95.0
            BackTint =20.0
            Begin
                Begin Label
                    Left =60
                    Width =4140
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BorderColor =16777215
                    ForeColor =7500402
                    Name ="Label12"
                    Caption ="Missing Records Report "
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =540
                    ForeTint =100.0
                End
                Begin Label
                    Left =7740
                    Top =480
                    Width =3750
                    Height =270
                    FontSize =9
                    ForeColor =8355711
                    Name ="Label20"
                    Caption ="*Shows records with a due date less than 30."
                    LayoutCachedLeft =7740
                    LayoutCachedTop =480
                    LayoutCachedWidth =11490
                    LayoutCachedHeight =750
                    ForeTint =100.0
                End
                Begin CommandButton
                    DisplayWhen =2
                    Left =10740
                    Top =60
                    Width =720
                    Height =300
                    ForeColor =16777215
                    Name ="Command21"
                    Caption ="Close"
                    LeftPadding =105
                    TopPadding =60
                    RightPadding =120
                    BottomPadding =165
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Command21\" Event=\"OnClick\" xmlns=\"http://schemas.microso"
                                "ft.com/office/accessservices/2009/11/application\" xmlns:a=\"http://schemas.micr"
                                "osoft.com/office/accessservices/"
                        End
                        Begin
                            Comment ="_AXL:2009/11/forms\"><Statements><Action Name=\"CloseWindow\"/></Statements></Us"
                                "erInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =10740
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =1
                    ForeShade =100.0
                    UseTheme =1
                    Gradient =2
                    BackColor =12874308
                    BackThemeColorIndex =8
                    BorderColor =12874308
                    BorderThemeColorIndex =8
                    BorderShade =100.0
                    HoverColor =13667945
                    HoverThemeColorIndex =8
                    HoverTint =80.0
                    PressedColor =10574387
                    PressedThemeColorIndex =8
                    PressedShade =80.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    QuickStyle =34
                    QuickStyleMask =-1
                    WebImagePaddingLeft =7
                    WebImagePaddingTop =4
                    WebImagePaddingRight =7
                    WebImagePaddingBottom =10
                    Overlaps =1
                End
                Begin Image
                    BackStyle =0
                    PictureType =2
                    Left =5100
                    Top =180
                    Width =1920
                    Height =600
                    BorderColor =10921638
                    Name ="Image23"
                    Picture ="MERM Logo"
                    GridlineColor =10921638

                    LayoutCachedLeft =5100
                    LayoutCachedTop =180
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =780
                    TabIndex =3
                End
                Begin CommandButton
                    DisplayWhen =2
                    Left =9240
                    Top =60
                    Width =360
                    TabIndex =1
                    Name ="Command24"
                    Caption ="Command24"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000017000000180000000100180000000000c00600001217000012170000 ,
                        0x0000000000000000fffffffffffffffffffffffffefefedadadac6c6c6c7c7c7 ,
                        0xc7c7c7c7c7c7c7c7c7c8c8c8c8c8c8c9c9c9cacacac9c9c9cacacac9c9c9caca ,
                        0xcac9c9c9cac9cae4e4e4ffffff000000fffffffffffffffffffffffffefefede ,
                        0xdedee7e7e7e8e8e8e8e8e8e8e8e8e8e8e8dfdfdee5e5e5dbdbdbcecfcedcdddc ,
                        0xd3d3d3d9d9d9d5d5d4dfdfdfd5d5d5dfdfdfffffff000000ffffffffffffffff ,
                        0xffffffffffffffebebebffffffffffffffffffffffffffffffd2d3d2adafad8b ,
                        0x8c899495939293919c9d9b8e908d9a9c99949693d3d4d2e8e8e8ffffff000000 ,
                        0xffffffffffffffffffffffffffffffececeaf4f4ff8e8effababffffffffffff ,
                        0xffe7e7e67a7c798f908e999a9880827e9b9c9a8183809496937577749d9f9ce4 ,
                        0xe4e4ffffff000000ffffffffffffffffffffffffffffffececeaf5f5ff7a7aff ,
                        0x7d7dfeb0b0fefffffffbfbfb7b7d7adcdcdb9ea09eb2b3b1cbcbca969895c0c1 ,
                        0xbfdedfdedadbd9e7e7e7ffffff000000ffffffffffffffffffffffffffffffe9 ,
                        0xe9e9ffffffdedefda4a4fe6767fee6e6fdfffffedadbdafcfcfcefefefe8e9e8 ,
                        0xffffffdfe0dcfffffdfffffffffeffe4e4e4ffffff000000ffffffffffffffff ,
                        0xffffffffffffffe8e8e8fffffffffffcf8f8fcbcbcfd7b7bfee6e6fcfcfcffff ,
                        0xfffcfffffefffffefffffce7e7ffc5c5fdcccafae5e1ffe6e6e4ffffff000000 ,
                        0xffffffffffffffffffffffffffffffe8e8e8fffffffbfbfbfcfcfbfffffb7676 ,
                        0xfda1a1fdaaaafca5a5fcc0c0fcc0c0fc6969fd6e6efd9e9efcb5b5fcffffffe4 ,
                        0xe4e3ffffff000000ffffffffffffffffffffffffffffffe7e7e7fffffffafafa ,
                        0xfafafafdfdfad1d1fa8282fcfffffaeaeafa6969fd5656fd8181fc9797fc9d9d ,
                        0xfcc6c6fbffffffe3e3e3ffffff000000ffffffffffffffffffffffffffffffe6 ,
                        0xe6e6fffffff8f8f8f8f8f8f8f8f8fefef87a7afcdcdcf9a9a9fb8383fbfdfdf8 ,
                        0xfffff8fdfdf8fdfdf8fdfdf8ffffffe2e2e2ffffff000000ffffffffffffffff ,
                        0xffffffffffffffe5e5e5fffffff7f7f7f7f7f7f7f7f7fcfcf7c9c9f86d6dfc64 ,
                        0x64fcf5f5f7f9f9f7f7f7f7f7f7f7f7f7f7f7f7f7ffffffe2e2e2ffffff000000 ,
                        0xffffffffffffffffffffffffffffffe4e4e4fffffff5f5f5f5f5f5f5f5f5f6f6 ,
                        0xf5fbfbf54747fcababf8fefef5f5f5f5f5f5f5f5f5f5f5f5f5f5f5f5fcfcfce0 ,
                        0xe0e0ffffff000000ffffffffffffffffffffffffffffffe3e3e3fefefef4f4f4 ,
                        0xf4f4f4f4f4f4f5f5f4f6f6f44c4cfccdcdf6f9f9f4f4f4f4f4f4f4f4f4f4f3f3 ,
                        0xf3f1f1f1f7f7f7dededeffffff000000ffffffffffffffffffffffffffffffe2 ,
                        0xe2e2fcfcfcf2f2f2f2f2f2f2f2f2f5f5f2dcdcf45656fad3d3f4f6f6f2f2f2f2 ,
                        0xf2f2f2f1f1f1eeeeeeebebebf0f0f0dbdbdbffffff000000ffffffffffffffff ,
                        0xffffffffffffffe2e2e2fafafaf1f1f1f1f1f1f1f1f1f6f6f1c0c0f48181f8f2 ,
                        0xf2f1f2f2f1f1f1f1efefefebebebe8e8e8e5e5e5eaeaead9d9d9ffffff000000 ,
                        0xffffffffffffffffffffffffffffffe6e6e1fffffaf6f6f1f6f6f1f6f6f1fcfc ,
                        0xf1c1c1f58080f9ddddf3f8f8f1f1f1eee9e9e9e6e6e6e3e3e3dededee2e2e2d4 ,
                        0xd4d4ffffff000000b9b9d56e6ead6f6fab6f6fab6f6fa95e5e9a6b6ba45f5f9d ,
                        0x60609b65659b6363995c5c9b29299e55559b6868969292b1e7e7e5dededed8d8 ,
                        0xd8cfcfcfcececed4d4d4ffffff0000007d7dbf0000aa0000aa0000aa2626b714 ,
                        0x14ae0000a14949b93f3fb30a0aa22727a700009200008f00008b00008047478d ,
                        0xdadad5b7b7b6a5a5a5a5a5a5b5b5b5efefefffffff0000008383ce0404cb0505 ,
                        0xcd0101cc6e6ee37474e31111cfa4a4ee8484e88686e39191e41111bc0000b400 ,
                        0x00ad0000a0494997c4c4bfb6b6b6bbbbbbc1c1c1e7e7e7ffffffffffff000000 ,
                        0x8b8bdb1c1ce51f1fe31717e27676eeb7b7f78f8ff09292ef4f4fe4aeaef1b6b6 ,
                        0xf24040d90101c50000be0000b145459dbcbcb7fbfbfbe9e9e9e5e5e5ffffffff ,
                        0xffffffffff0000009494e63838fa3c3cf63535f57f7ff9babafb8c8cf6a9a9f7 ,
                        0xa5a5f56c6cecafaff44b4be00101cf0000c60000b94040a0bebeb8e5e5e5e2e2 ,
                        0xe2ffffffffffffffffffffffff0000009c9ced4545fa4848f64545f54545f14b ,
                        0x4bec3535e73d3de33636dd2020d63030d32424cc0f0fc20a0aba0505ae4646a3 ,
                        0xc5c5bfe7e7e7ffffffffffffffffffffffffffffff000000efeffcdcdcfadcdc ,
                        0xf9dcdcf9dbdbf8bcbcdaceceeac4c4e1c4c4e0c3c3debdbdd9babad4b7b7cfb3 ,
                        0xb3caacacc5b1b1c0e5e5e4ffffffffffffffffffffffffffffffffffff000000 ,
                        0xfffffffffffffffffffffffffffffff9f9f5f8f8f5f6f6f2f6f6f2f5f5f2f4f4 ,
                        0xf1f3f3eff2f2eff1f1edf1f1edf5f5f3ffffffffffffffffffffffffffffffff ,
                        0xffffffffff000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Picture ="PDF1.bmp"

                    LayoutCachedLeft =9240
                    LayoutCachedTop =60
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    DisplayWhen =2
                    Left =9780
                    Top =60
                    Width =360
                    TabIndex =2
                    Name ="Command23"
                    Caption ="Command24"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dd8444ddddddddddddd8744ddddddddd44444444dddddddd ,
                        0x44444444ddddddddddd8744d726ddddddd8444d6262ddddddddddd6262626262 ,
                        0xdddd26f6f62ffff6dddd62f8f26262f2dddd262f262ffff6dddd626f626262f2 ,
                        0xdddd26f8f62ffff6dddd62f2f26262f2dddd2626262ffff6dddd626262626262 ,
                        0xddddddd8762ddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =9780
                    LayoutCachedTop =60
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    ThemeFontIndex =-1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin PageHeader
            Height =420
            Name ="PageHeaderSection"
            AlternateBackShade =95.0
            Begin
                Begin Label
                    TextAlign =1
                    Left =60
                    Top =60
                    Width =2100
                    Height =315
                    BorderColor =16777215
                    ForeColor =11573124
                    Name ="Assigned To_Label"
                    Caption ="Assigned To"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Assigned_To_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =2100
                    Top =60
                    Width =660
                    Height =315
                    BorderColor =16777215
                    ForeColor =11573124
                    Name ="Site_Label"
                    Caption ="Site"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =3960
                    Top =60
                    Width =4320
                    Height =315
                    BorderColor =16777215
                    ForeColor =11573124
                    Name ="Document Name_Label"
                    Caption ="MERM Document Name"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Document_Name_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3960
                    LayoutCachedTop =60
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =8340
                    Top =60
                    Width =960
                    Height =315
                    BorderColor =16777215
                    ForeColor =11573124
                    Name ="Asset_Label"
                    Caption ="Asset"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8340
                    LayoutCachedTop =60
                    LayoutCachedWidth =9300
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =9240
                    Top =60
                    Width =1080
                    Height =315
                    BorderColor =16777215
                    ForeColor =11573124
                    Name ="AQMD ID_Label"
                    Caption ="AQMD ID"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="AQMD_ID_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =9240
                    LayoutCachedTop =60
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =3
                    Left =10320
                    Top =60
                    Width =1140
                    Height =315
                    BorderColor =16777215
                    ForeColor =11573124
                    Name ="MERM Date_Label"
                    Caption ="Due Date"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MERM_Date_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10320
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =375
                End
                Begin Label
                    TextAlign =1
                    Left =2760
                    Top =60
                    Width =1116
                    Height =324
                    BorderColor =16777215
                    ForeColor =11573124
                    Name ="Label17"
                    Caption ="MERM Type"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =2760
                    LayoutCachedTop =60
                    LayoutCachedWidth =3876
                    LayoutCachedHeight =384
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            BackColor =15921906
            Name ="GroupHeader0"
            AlternateBackColor =15921906
            AlternateBackShade =95.0
            BackShade =95.0
            Begin
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =360
                    Width =2100
                    Height =330
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="Group Of Assigned To"
                    ControlSource ="Assigned To"
                    EventProcPrefix ="Group_Of_Assigned_To"
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Condition ="IsNull([Screen].[ActiveControl])"
                            Action ="Beep"
                        End
                        Begin
                            Condition ="..."
                            Action ="StopMacro"
                        End
                        Begin
                            Action ="OpenForm"
                            Argument ="MERM Subform"
                            Argument ="0"
                            Argument =""
                            Argument ="=\"[ID]=\" & [Screen].[ActiveControl]"
                            Argument ="-1"
                            Argument ="3"
                        End
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="Requery"
                            Argument ="=[Screen].[ActiveControl].[Name]"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"Group Of Assigned To\" Event=\"OnClick\" xmlns=\"http://sche"
                                "mas.microsoft.com/office/accessservices/2009/11/application\"><Statements><Condi"
                                "tionalBlock><If><Condition>IsNu"
                        End
                        Begin
                            Comment ="_AXL:ll([Screen].[ActiveControl])</Condition><Statements><Action Name=\"Beep\"/>"
                                "<Action Name=\"StopMacro\"/></Statements></If></ConditionalBlock><Action Name=\""
                                "OpenForm\"><Argument Name=\"FormName\">MERM Subform</Argument><Argument Name=\"W"
                                "hereCondition\">=\"[ID]="
                        End
                        Begin
                            Comment ="_AXL:\" &amp; [Screen].[ActiveControl]</Argument><Argument Name=\"WindowMode\">D"
                                "ialog</Argument></Action><Action Name=\"OnError\"/><Action Name=\"Requery\"><Arg"
                                "ument Name=\"ControlName\">=[Screen].[ActiveControl].[Name]</Argument></Action><"
                                "/Statements></UserInte"
                        End
                        Begin
                            Comment ="_AXL:rfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =360
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =330
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =330
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2160
                    Width =600
                    Height =330
                    ColumnWidth =750
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="Site"
                    ControlSource ="Site"
                    GridlineColor =10921638

                    LayoutCachedLeft =2160
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =4020
                    Width =4260
                    Height =330
                    ColumnWidth =3450
                    TabIndex =1
                    BorderColor =13553360
                    ForeColor =12673797
                    Name ="DocumentName1"
                    ControlSource ="DocumentName"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =4020
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =330
                    ForeShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8280
                    Width =960
                    Height =330
                    TabIndex =2
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="Asset"
                    ControlSource ="Asset"
                    GridlineColor =10921638

                    LayoutCachedLeft =8280
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =9180
                    Width =1140
                    Height =330
                    TabIndex =3
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="AQMD ID"
                    ControlSource ="AQMD_ID"
                    EventProcPrefix ="AQMD_ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =9180
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =10320
                    Width =1140
                    Height =330
                    TabIndex =4
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="MERM Date"
                    ControlSource ="EventDate"
                    EventProcPrefix ="MERM_Date"
                    GridlineColor =10921638
                    ShowDatePicker =1

                    LayoutCachedLeft =10320
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2760
                    Width =1200
                    Height =315
                    TabIndex =5
                    Name ="CurrentMonth"
                    ControlSource ="Building"

                    LayoutCachedLeft =2760
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =315
                End
            End
        End
        Begin PageFooter
            Height =420
            Name ="PageFooterSection"
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =5040
                    Height =330
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="Text13"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =390
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6420
                    Top =60
                    Width =5040
                    Height =330
                    TabIndex =1
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="Text14"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6420
                    LayoutCachedTop =60
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =390
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
            AutoHeight =1
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

Private Sub Command24_Click()
'export to PDF
Dim StrPathFile As String

StrPathFile = "C:\Users\G87022\Downloads\Report2.pdf"
DoCmd.OutputTo acOutputReport, "rptNext30", acFormatPDF, StrPathFile, True, acExportQualityScreen
End Sub

Private Sub DocumentName1_Click()
DoCmd.OpenForm "EquipForm", , , "DocumentName = '" & Me.DocumentName & "'"
    Forms!EquipForm.ID1.SetFocus
   ' Forms!EquipForm.ID1.SelStart = 0
   ' Forms!EquipForm.ID1.SelLength = 0
End Sub
