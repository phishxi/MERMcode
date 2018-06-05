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
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =40
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x7a02b0b34b0fe540
    End
    RecordSource ="UserQuery"
    Caption ="User Report"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000002d00006801000001000000 ,
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
            ForeShade =50.0
            GridlineShade =65.0
            BackColor =-2147483633
            BorderLineStyle =0
            BorderShade =90.0
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
            Height =1020
            BackColor =13020235
            Name ="ReportHeader"
            AlternateBackShade =95.0
            Begin
                Begin Image
                    BackStyle =0
                    PictureType =2
                    Left =4800
                    Width =2460
                    Height =1020
                    BorderColor =10921638
                    Name ="Image32"
                    Picture ="MERM Logo"
                    GridlineColor =10921638

                    LayoutCachedLeft =4800
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1020
                    TabIndex =6
                End
                Begin Label
                    Left =60
                    Top =60
                    Width =5220
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BorderColor =16777215
                    ForeColor =16777215
                    Name ="Label10"
                    Caption ="User Defined Report"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =600
                    ForeTint =100.0
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1560
                    Top =540
                    Width =1260
                    Height =315
                    ColumnOrder =3
                    BorderColor =16249583
                    ForeColor =5855577
                    Name ="Text22"
                    ControlSource ="=IIf(IsNull([Forms]![MERMHome]![BeginDate]),\" \",[Forms]![MERMHome]![BeginDate]"
                        ")"
                    GridlineColor =10921638

                    LayoutCachedLeft =1560
                    LayoutCachedTop =540
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =855
                    BorderShade =100.0
                    ForeShade =100.0
                End
                Begin Label
                    Left =120
                    Top =540
                    Width =1425
                    Height =300
                    BorderColor =16777215
                    ForeColor =5855577
                    Name ="Label24"
                    Caption ="Report Period:"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =540
                    LayoutCachedWidth =1545
                    LayoutCachedHeight =840
                    ForeTint =100.0
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3180
                    Top =540
                    Width =1260
                    Height =315
                    ColumnOrder =2
                    TabIndex =1
                    BorderColor =16249583
                    ForeColor =5855577
                    Name ="Text25"
                    ControlSource ="=IIf(IsNull([Forms]![MERMHome]![EndDate]),\" \",[Forms]![MERMHome]![EndDate])"
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedTop =540
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =855
                    BorderShade =100.0
                    ForeShade =100.0
                End
                Begin Label
                    Left =2940
                    Top =540
                    Width =150
                    Height =315
                    FontWeight =700
                    BorderColor =16777215
                    Name ="Label26"
                    Caption ="-"
                    GridlineColor =10921638
                    LayoutCachedLeft =2940
                    LayoutCachedTop =540
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =855
                    ForeTint =100.0
                End
                Begin CommandButton
                    DisplayWhen =2
                    Left =10740
                    Top =60
                    Width =660
                    TabIndex =2
                    ForeColor =16777215
                    Name ="Command27"
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
                                "nterfaceMacro For=\"Command27\" Event=\"OnClick\" xmlns=\"http://schemas.microso"
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
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =1
                    ForeShade =100.0
                    UseTheme =1
                    Shape =1
                    Gradient =2
                    BackColor =0
                    BackThemeColorIndex =0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    HoverColor =-2147483617
                    PressedColor =10921638
                    PressedThemeColorIndex =1
                    PressedShade =65.0
                    HoverForeColor =-2147483610
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =36
                    QuickStyleMask =-1
                    Overlaps =1
                End
                Begin CommandButton
                    DisplayWhen =2
                    Left =10200
                    Top =540
                    Width =360
                    TabIndex =3
                    Name ="Command34"
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

                    LayoutCachedLeft =10200
                    LayoutCachedTop =540
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =900
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    DisplayWhen =2
                    Left =10740
                    Top =540
                    Width =360
                    TabIndex =4
                    Name ="Command23"
                    Caption ="Command24"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dd8444ddddddddddddd8744ddddddddd44444444dddddddd ,
                        0x44444444ddddddddddd8744d726ddddddd8444d6262ddddddddddd6262626262 ,
                        0xdddd26f6f62ffff6dddd62f8f26262f2dddd262f262ffff6dddd626f626262f2 ,
                        0xdddd26f8f62ffff6dddd62f2f26262f2dddd2626262ffff6dddd626262626262 ,
                        0xddddddd8762ddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    LayoutCachedLeft =10740
                    LayoutCachedTop =540
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =900
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9090
                    Top =600
                    Width =840
                    Height =315
                    ColumnOrder =0
                    FontWeight =700
                    TabIndex =5
                    ForeColor =15921906
                    Name ="Text35"
                    ControlSource ="=Count(*)"

                    LayoutCachedLeft =9090
                    LayoutCachedTop =600
                    LayoutCachedWidth =9930
                    LayoutCachedHeight =915
                    ForeThemeColorIndex =1
                    ForeShade =95.0
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =7515
                            Top =600
                            Width =1560
                            Height =315
                            FontWeight =700
                            ForeColor =15921906
                            Name ="Label36"
                            Caption ="Items in Report:"
                            LayoutCachedLeft =7515
                            LayoutCachedTop =600
                            LayoutCachedWidth =9075
                            LayoutCachedHeight =915
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =95.0
                        End
                    End
                End
            End
        End
        Begin PageHeader
            Height =495
            Name ="PageHeaderSection"
            AlternateBackShade =95.0
            Begin
                Begin Label
                    TextAlign =1
                    Left =360
                    Top =60
                    Width =1185
                    Height =315
                    BorderColor =16777215
                    Name ="Assigned To_Label"
                    Caption ="Assigned To"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Assigned_To_Label"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =60
                    LayoutCachedWidth =1545
                    LayoutCachedHeight =375
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =10320
                    Top =60
                    Width =1080
                    Height =315
                    BorderColor =16777215
                    Name ="EventDate_Label"
                    Caption ="Due Date"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10320
                    LayoutCachedTop =60
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =375
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =1
                    Left =3420
                    Top =60
                    Width =1635
                    Height =315
                    BorderColor =16777215
                    Name ="DocumentName_Label"
                    Caption ="Document Name"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3420
                    LayoutCachedTop =60
                    LayoutCachedWidth =5055
                    LayoutCachedHeight =375
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =7440
                    Top =60
                    Width =1230
                    Height =315
                    BorderColor =16777215
                    Name ="Asset_Label"
                    Caption ="Asset #"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7440
                    LayoutCachedTop =60
                    LayoutCachedWidth =8670
                    LayoutCachedHeight =375
                    ForeTint =100.0
                End
                Begin Label
                    TextAlign =2
                    Left =8640
                    Top =60
                    Width =990
                    Height =315
                    BorderColor =16777215
                    Name ="AQMD_ID_Label"
                    Caption ="AQMD ID"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8640
                    LayoutCachedTop =60
                    LayoutCachedWidth =9630
                    LayoutCachedHeight =375
                    ForeTint =100.0
                End
                Begin Line
                    Left =60
                    Top =480
                    Width =11400
                    Name ="Line13"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =480
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =495
            Name ="GroupHeader0"
            AlternateBackColor =15921906
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =2640
                    Height =330
                    ColumnWidth =1890
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="Assigned To"
                    ControlSource ="Assigned To"
                    EventProcPrefix ="Assigned_To"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =390
                End
                Begin Line
                    Left =60
                    Top =480
                    Width =11460
                    BorderColor =12566463
                    Name ="Line14"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2040
                    Top =60
                    Width =1200
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    BorderColor =13553360
                    ForeColor =7949855
                    Name ="AccessTotalsMERMType"
                    ControlSource ="=Count([MERMType])"
                    ControlTipText ="MERMType Value Count"
                    GridlineColor =10921638

                    LayoutCachedLeft =2040
                    LayoutCachedTop =60
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =4
                End
                Begin Label
                    Left =3300
                    Top =60
                    Width =885
                    Height =345
                    FontSize =12
                    FontWeight =700
                    ForeColor =7949855
                    Name ="Label39"
                    Caption ="Records"
                    LayoutCachedLeft =3300
                    LayoutCachedTop =60
                    LayoutCachedWidth =4185
                    LayoutCachedHeight =405
                    ForeThemeColorIndex =4
                    ForeTint =100.0
                    ForeShade =50.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =360
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =9870
                    Width =1380
                    Height =330
                    TabIndex =1
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="EventDate"
                    ControlSource ="EventDate"
                    GridlineColor =10921638

                    LayoutCachedLeft =9870
                    LayoutCachedWidth =11250
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =2940
                    Width =4995
                    Height =330
                    ColumnWidth =4416
                    BorderColor =13553360
                    ForeColor =12673797
                    Name ="DocumentName"
                    ControlSource ="DocumentName"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2940
                    LayoutCachedWidth =7935
                    LayoutCachedHeight =330
                    ForeShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =8100
                    Width =885
                    Height =330
                    ColumnWidth =1050
                    TabIndex =2
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="Asset"
                    ControlSource ="Asset"
                    GridlineColor =10921638

                    LayoutCachedLeft =8100
                    LayoutCachedWidth =8985
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Left =9015
                    Width =855
                    Height =330
                    ColumnWidth =870
                    TabIndex =3
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="AQMD_ID"
                    ControlSource ="AQMD_ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =9015
                    LayoutCachedWidth =9870
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =4
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Width =780
                    Height =315
                    TabIndex =4
                    BorderColor =13553360
                    ForeColor =9605778
                    Name ="Text15"
                    ControlSource ="Building"
                    GridlineColor =10921638

                    LayoutCachedLeft =2040
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =315
                    ForeShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =12
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1500
                    Width =435
                    Height =315
                    TabIndex =5
                    BorderColor =13553360
                    ForeColor =9605778
                    Name ="Text19"
                    ControlSource ="Site"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedWidth =1935
                    LayoutCachedHeight =315
                    ForeShade =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Width =1500
                    Height =315
                    TabIndex =6
                    BorderColor =13553360
                    ForeColor =9605778
                    Name ="Text29"
                    ControlSource ="Priority"
                    GridlineColor =10921638

                    LayoutCachedWidth =1500
                    LayoutCachedHeight =315
                    ForeShade =100.0
                End
            End
        End
        Begin PageFooter
            Height =450
            Name ="PageFooterSection"
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =5040
                    Height =330
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="Text11"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =450
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6360
                    Top =120
                    Width =5040
                    Height =330
                    TabIndex =1
                    BorderColor =13553360
                    ForeColor =3484194
                    Name ="Text12"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedTop =120
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =450
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =240
            Name ="ReportFooter"
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
Dim lngColumn As Long
Dim xlx As Object, xlw As Object, xls As Object, xlc As Object
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim strPathFileName As String, strWorksheetName As String
Dim strRecordsetDataSource As String
Dim blnEXCEL As Boolean, blnHeaderRow As Boolean

CurrentDb.Execute "Insert Into tbl_TempUserQuery Select UserQuery.* From UserQuery"

'Dim strSQL_Delete As String
'Dim strSQL_Insert As String
'Dim strTable As String
'strTable = "tbl_TempUserQuery"
'strSQL_Delete = ""
'strSQL_Insert = ""
'
'strSQL_Delete = "Delete * FROM " & strTable
'strSQL_Insert = "INSERT INTO tbl_TempUserQuery " _
'& "SELECT UserQuery.* From UserQuery"
'CurrentDb.Execute strSQL_Delete
'CurrentDb.Execute strSQL_Insert

blnEXCEL = False

' Replace C:\Filename.xls with the actual path and filename
' that will be used to save the new EXCEL file into which you
' will write the data
strPathFileName = "C:\Filename.xls"

' Replace QueryOrTableName with the real name of the table or query
' whose data are to be written into the worksheet
strRecordsetDataSource = "tbl_TempUserQuery"

' Replace True with False if you do not want the first row of
' the worksheet to be a header row (the names of the fields
' from the recordset)
blnHeaderRow = True

' Establish an EXCEL application object
On Error Resume Next
Set xlx = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
      Set xlx = CreateObject("Excel.Application")
      blnEXCEL = True
End If
Err.Clear
On Error GoTo 0

' Change True to False if you do not want the workbook to be
' visible when the code is running
xlx.Visible = True

' Create a new EXCEL workbook
Set xlw = xlx.Workbooks.Add

' Rename the first worksheet in the EXCEL file to be the first 31
' characters of the string in the strRecordsetDataSource variable
Set xls = xlw.Worksheets(1)
xls.name = Trim(Left(strRecordsetDataSource, 31))

' Replace A1 with the cell reference of the first cell into which the
' headers will be written (blnHeaderRow = True), or into which the data
' values will be written (blnHeaderRow = False)
Set xlc = xls.Range("A1")

Set dbs = CurrentDb()

Set rst = dbs.OpenRecordset(strRecordsetDataSource, dbOpenDynaset, dbReadOnly)

If rst.EOF = False And rst.BOF = False Then
      ' Write the header row to worksheet
      If blnHeaderRow = True Then
            For lngColumn = 0 To rst.Fields.Count - 1
                  xlc.Offset(0, lngColumn).Value = rst.Fields(lngColumn).name
            Next lngColumn
            Set xlc = xlc.Offset(1, 0)
      End If

      ' copy the recordset's data to worksheet
      xlc.CopyFromRecordset rst
End If

rst.Close
Set rst = Nothing
dbs.Close
Set dbs = Nothing

' Save and close the EXCEL file, and clean up the EXCEL objects
Set xlc = Nothing
Set xls = Nothing
'xlw.SaveAs strPathFileName
'xlw.Close False
Set xlw = Nothing
'If blnEXCEL = True Then xlx.Quit
Set xlx = Nothing

MsgBox "Query data loaded to temp table."

End Sub



Private Sub Command23_Click()
Dim StrPathFile As String
StrPathFile = CreateObject("WScript.Shell").specialfolders("MyDocuments")
Debug.Print StrPathFile

StrPathFile = StrPathFile & "\" & "User Report.xlsx"
DoCmd.OutputTo acOutputQuery, "UserQuery", acFormatXLSX, StrPathFile, True, acExportQualityScreen
'Alternative
'DoCmd.OutputTo acOutputQuery, "UserQuery", "User Report(*.xlsx)", , True
End Sub

Private Sub Command34_Click()
'export to PDF
Dim StrPathFile As String
StrPathFile = CreateObject("WScript.Shell").specialfolders("MyDocuments")

Debug.Print StrPathFile

StrPathFile = StrPathFile & "\" & "User Report.pdf"
DoCmd.OutputTo acOutputReport, "UserQuery1", acFormatPDF, StrPathFile, True, acExportQualityScreen

End Sub

Private Sub DocumentName_Click()
DoCmd.OpenForm "EquipForm", , , "DocumentName = '" & Me.DocumentName & "'"
    Forms!EquipForm.ID1.SetFocus
End Sub
