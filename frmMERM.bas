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
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12720
    DatasheetFontHeight =11
    ItemSuffix =51
    Left =30555
    Top =3330
    Right =-22006
    Bottom =11040
    DatasheetGridlinesColor =0
    RecSrcDt = Begin
        0x9956ce97921ae540
    End
    Caption ="Keyword Search"
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            GridlineColor =16777215
            BorderShade =65.0
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeTint =75.0
            GridlineShade =65.0
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderTint =60.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderShade =65.0
            GridlineShade =65.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BorderShade =65.0
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
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            BorderShade =65.0
            ForeShade =50.0
            GridlineShade =65.0
        End
        Begin Subform
            BorderLineStyle =0
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =405
            BackColor =12428174
            Name ="FormHeader"
            AlternateBackShade =95.0
            Begin
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =60
                    Width =3315
                    Height =405
                    FontSize =16
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="Label8"
                    Caption ="Browse All Requirements"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedWidth =3375
                    LayoutCachedHeight =405
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7320
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackShade =95.0
            Begin
                Begin Subform
                    OverlapFlags =215
                    Left =60
                    Top =1500
                    Width =12540
                    Height =5520
                    TabIndex =1
                    BorderColor =10921638
                    Name ="subMERMList"
                    SourceObject ="Form.sfrmMERMList"
                    GridlineColor =10921638
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedLeft =60
                    LayoutCachedTop =1500
                    LayoutCachedWidth =12600
                    LayoutCachedHeight =7020
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =60
                            Top =1200
                            Width =1560
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label1"
                            Caption ="Search Results:"
                            GridlineColor =10921638
                            LayoutCachedLeft =60
                            LayoutCachedTop =1200
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =1515
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2580
                    Top =120
                    Width =2640
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtKeywords"
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =120
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =435
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =900
                            Top =120
                            Width =1620
                            Height =315
                            BorderColor =8355711
                            Name ="Label3"
                            Caption ="Keyword Search:"
                            GridlineColor =10921638
                            LayoutCachedLeft =900
                            LayoutCachedTop =120
                            LayoutCachedWidth =2520
                            LayoutCachedHeight =435
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =5400
                    Top =120
                    Width =1320
                    Height =300
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnSearch"
                    Caption ="Search"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5400
                    LayoutCachedTop =120
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =420
                    BackColor =15123357
                    BorderColor =15123357
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6900
                    Top =120
                    Width =300
                    Height =300
                    TabIndex =3
                    ForeColor =4210752
                    Name ="Command7"
                    Caption ="Command7"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xddddddddddddddddddddddddddddddddddddddddddd74ddddd44444ddddd47dd ,
                        0xdd4444ddddddd4dddd444dddddddd4dddd44d4ddddddd4dddd4ddd44dddd47dd ,
                        0xdddddddd44447ddddddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
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
                    ControlTipText ="Undo Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =6900
                    LayoutCachedTop =120
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =420
                    BackColor =15123357
                    BorderColor =15123357
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Image
                    Left =9600
                    Width =2460
                    Height =540
                    BorderColor =10921638
                    Name ="Image7"
                    Picture ="Hi Res Logo NGC.gif"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    ImageData = Begin
                        0x474946383961cc016c00f70000ffffffffffccffff99ffff66ffff33ffff00ff ,
                        0xccffffccccffcc99ffcc66ffcc33ffcc00ff99ffff99ccff9999ff9966ff9933 ,
                        0xff9900ff66ffff66ccff6699ff6666ff6633ff6600ff33ffff33ccff3399ff33 ,
                        0x66ff3333ff3300ff00ffff00ccff0099ff0066ff0033ff0000ccffffccffcccc ,
                        0xff99ccff66ccff33ccff00ccccffcccccccccc99cccc66cccc33cccc00cc99ff ,
                        0xcc99cccc9999cc9966cc9933cc9900cc66ffcc66cccc6699cc6666cc6633cc66 ,
                        0x00cc33ffcc33cccc3399cc3366cc3333cc3300cc00ffcc00cccc0099cc0066cc ,
                        0x0033cc000099ffff99ffcc99ff9999ff6699ff3399ff0099ccff99cccc99cc99 ,
                        0x99cc6699cc3399cc009999ff9999cc9999999999669999339999009966ff9966 ,
                        0xcc9966999966669966339966009933ff9933cc99339999336699333399330099 ,
                        0x00ff9900cc99009999006699003399000066ffff66ffcc66ff9966ff6666ff33 ,
                        0x66ff0066ccff66cccc66cc9966cc6666cc3366cc006699ff6699cc6699996699 ,
                        0x666699336699006666ff6666cc6666996666666666336666006633ff6633cc66 ,
                        0x33996633666633336633006600ff6600cc66009966006666003366000033ffff ,
                        0x33ffcc33ff9933ff6633ff3333ff0033ccff33cccc33cc9933cc6633cc3333cc ,
                        0x003399ff3399cc3399993399663399333399003366ff3366cc33669933666633 ,
                        0x66333366003333ff3333cc3333993333663333333333003300ff3300cc330099 ,
                        0x33006633003333000000ffff00ffcc00ff9900ff6600ff3300ff0000ccff00cc ,
                        0xcc00cc9900cc6600cc3300cc000099ff0099cc00999900996600993300990000 ,
                        0x66ff0066cc0066990066660066330066000033ff0033cc003399003366003333 ,
                        0x0033000000ff0000cc000099000066000033000000ffffff0000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000002c00000000cc016c000008ff0001081c48b0a0 ,
                        0xc18308132a5cc8b0a1c38710234a9c48b1a2c58b18336adcc8b1a3c78f20438a ,
                        0x1c49b2a4c9932853aa5cc9b2a5cb973063ca9c49b3a6cd9b3873eadcc9b3a7cf ,
                        0x9f40830a1d4ab4a8d1a348932a5dcab4a9d3a750a34a9d4ab5aad5ab58b36add ,
                        0xcab5abd7af60c38a1d4bb6acd9b368d3aa5dcbb6addbb770e3ca9d4bb7aeddbb ,
                        0x78f3eaddcbb7afdfbf80030b1e4cb8b0e1c388132b5eccb8b1e3c790234b9e4c ,
                        0xb9b2e5cb3949a858a16273e7859e3daf209190c48a2a55f2a07ea242a166cea1 ,
                        0x61c346d85976eddbb16f03c00d9bf441de9b7de7ce3d3ae28a27a855b3766d5b ,
                        0x85e6dd240c883c8ebafa7283c39bfba61dfd796bee06bc0fff741ebef577d0dd ,
                        0xcd8f4f1f9d2179f1a09baf009d9b7ef39ca4f2ebcf5f256195fdfacd57d00a79 ,
                        0x90a20a32a42098a07efd1d04e08300aa361009105608e1137a40d8a04115f6f7 ,
                        0x8485fba9925f1e021e444215a3e4e7cc82c81c484a1ee70d640088f929286146 ,
                        0x2aa0b81f32082a48ca1f4f0cf4218d068e58e26efaf9484a8c03fdb7e38b007c ,
                        0xe8e27e0a0db92320a455786441240052e37e41fa5761980715f8601e09a90021 ,
                        0x9a370db9a2332d22c8a44024fc81209c078e12e313297eb9e28122ae480a204c ,
                        0x561128297822f3669c6fe6d79a93892e7aa0a48a229ac7138056bae29c00fc47 ,
                        0x292902fd9129a58b223a227666aa98a0882e8a480a99027dff48e9a4a6364aca ,
                        0x96c6fdb1e0a42d9a9a29290de6d168a488c6692a9402416aa0a26c7248eb9f2b ,
                        0x1890878ba572dae9aa8c021be5acc8343be082c382ea5f8f946e58909aa46a7b ,
                        0x501571d28aab4c76aea8a08d074939af2a64aea02bb6aec2896db1e20ec4e27e ,
                        0x7dd628628f2faa69ecbc5f1ecbb00a762eb8a0b7130ebc607f2b20fc65c10aca ,
                        0x2b628927aaeaafbc827a7c2b41491adcaba0a6b6380a452a648860ab062768ec ,
                        0x9200acf0e5cca4140ceeaaaf0a9460a31fd72bb1bc7f98a6e0a1069a4ba79d07 ,
                        0x039a1f69bace7cf3baf9d18c20ac051528f5bd4e37d970af7262b7737e7fe0a4 ,
                        0xf3a17fc6c93500114b1df07fed929b75ddfa356b2faf30eeffc627d07feaf19f ,
                        0xd452b77da8d4402e4bf876040dde6ecfa1d6daeb13e579aaaaa261ae9062c9cb ,
                        0x16d972d42fc7ea6bafa320a723e7654764b9c73cbaca739c6c425debab95771e ,
                        0x759098ee8c4cda07a55878d0d32ada7182bc8fe725d0087b383bd027135460dd ,
                        0xa44b57e6ae2b2f4b319d222b7e6a41961ffa36bc1a2f2ce29e9dcb1b2cc3dadf ,
                        0x0c787e03a5d8a3ab47ae406d8d80f87cf7af4a06e8fee5d70bbd3e32181bd881 ,
                        0x8a279002710e6756ba19d9e6d7b2a001c07dbf3ad2286e26a21244240f0cf398 ,
                        0x723aa302e4f4693e0a3b16e49c473679cd6773dafb9eac2c9633e6b1ae79511a ,
                        0xdae77a453541dd4955cdaad3bfee0640844cab810b4316f7ff7e86303c710d5c ,
                        0x3d22204d74363ac221c85ba2baa1a300202cedc1e90f55e8cc137485a706b6e6 ,
                        0x43bd3a98125520c54d19845d037346d800e084f54df18c37f41169aac8aaad75 ,
                        0xed80a43040eebad8ab4b75c680971391f214a8c49ac1a97fa591dd9f46c4b873 ,
                        0x05ab6da6e2da0f7bd5333dee4c44a13388a8f8a52da8c5714ac1f21ceb448426 ,
                        0x2b35ea6a2dfcd7af02d6353ffde867aae8df047db5bcfc48af533c8b244e3028 ,
                        0x2f87590d5f02c918f2d448c58335f06d7d429888a8b6ac1bc628662e829d4174 ,
                        0xb848331ac44bbf72c61f38631b26728e98de34961299682c67745062adca4323 ,
                        0xe9d62a1ef54757f303a6d8a2969f7725c46b3c2a56e91ca2ff26ea8d70203afb ,
                        0x6490e868acef09338cc878593f5555bedde95086cc2b9b278b54ce2529ab7cab ,
                        0x5a232005550529a9687706f1a8cd02c92c817dee9f3561e2ecaaf9a5d6044f6e ,
                        0xe7ac9e33f4b42e9922c8922c4d9a403a08885fee6e4e600498bace95cf1d4268 ,
                        0x7f250b5d1555d6a0cd00325355a0d0b21ab54612448d67acf1a733bc05c61222 ,
                        0x1221742b92a0ec59a6ea1da8a904ba9ba9d0e44d838de296ce6bd9ccfa63c095 ,
                        0xe5724a361c9581d8aab20532d4ae18c5592bf18a2685b18c95feab14d3e819b4 ,
                        0x15f6728d3179a9c4c4d72b3d84106168ba9c7ebe77ad06c6126e08431820f6f5 ,
                        0xc21f5950939434552323b75805aaecaadaeae7297fe4897dffa50c41fd61d794 ,
                        0x9e489b63f55205186ce78f9e809c7d1d6c454a5c0885663bd46be92a45cf0521 ,
                        0x1011348ad126496e6c0a6e1cd7a88278224833d47b9dc9a4a82404cd07907e15 ,
                        0x2a057778bd492eab41225d512677fab5d955ea4ea3d0978f7a85937ede501510 ,
                        0x7cdcbdb4d741ef267720f12ad2819eb0b6ba19f35e0a8a6abd642a44821cb464 ,
                        0x1a33dc021534c7066a185b6f02120028d4cef18175782d221145e747b693aa13 ,
                        0x224efa2f29b623d5fdb28f8aaff5e52af789a47822b649ffc56da7202948bf92 ,
                        0x2e7799fa2cba909787042b0eb755a319933038ba66290c9503a1722dc32842bb ,
                        0x2e1827ba6d20815c074b88760bb4146d11679d6433eaa239ffa26c2bdf8b0fe2 ,
                        0xa5f7a5ae6b5febd70cd3a8a2fef8f750f9fce482c214d646cd9720b2a51e7091 ,
                        0xb767d6893822aa3ad899811c6755f899a1bbdd6de736f43c32f72fd112ebcc25 ,
                        0xbb5527abad75d1899a62a7659823952528bf642ce7a4b30ce86edd12c9f8b530 ,
                        0xde80d5c668463aa26fbd498d0fc53ba0d55acfaad8cc47ebb8c6802e5244591d ,
                        0xdda634e326cfc9b3208e5d55ff029ae79ee9a74fac9a1d4dc34cb63f7010bd87 ,
                        0x72697d91010813c96c762af6b1a05c348a3f90682221eca2a5c773b605916675 ,
                        0x33d393695e6aeaefb4ac519c0a2ebc7109689c05ef587cbd2e487566322891f1 ,
                        0xaadb2c26f3045bc0f5e95448a37b6331edecf095daecb84dff0373101de8dffb ,
                        0xc2f6cc4b1ed8f5c068cc15a149bb12f3d6f36c66f383fca197ae4af8830f0423 ,
                        0xd8846682f74d507f867dc312190085c79a4fc46899df017191a2ecdbb90d2fd5 ,
                        0x512770d3220dde99b736e3c9152d5d80e625489f4c369f8b0a6a4e6cd65878f4 ,
                        0xa3c160de2c58494ebb99e6f71da8eb34d177d2f9ca0745f8fab1c855c2d995b6 ,
                        0xdb87754ada70c6c27e32269db7c815b74bbad9432e6b4e338adff55c4d3e3e52 ,
                        0x776909d201f14c5e9f3e79e7def5845c16e951d9fc11f718e88c20e9c156957a ,
                        0xd5669ef0bcb6f54ad4b6b2e64642986407ee547b38339482dddcb89322234896 ,
                        0x7398b916fd2c03e10ed0451348c19c707c92a74db6376c56ff1504d5ac871faa ,
                        0x44f8fc5d6839993a82be1da0272fa1e2befa92b04e754b7d35a6813a813d3cda ,
                        0x786036221dab562cd7434422621028843009e77b68831057775f0d824e0ac235 ,
                        0xc264327455662d26374b336efb613e1e017aa6d62bf6861c7ac0306c3252fe72 ,
                        0x6d76a74a74f53e8c9607a9e22fc71281ca547a05a44de7e138fd027d479320e7 ,
                        0xc1441b12500063835544492ee760aed61ae4447714f343b6f333d602133c941f ,
                        0x878624c8132869972cbee5654dd42b0d12762c33275727288726848bb46d5fe2 ,
                        0x3d13c63466e4386d935c4b164d01442d7e6543be9707a7c549c2c711f8f4644b ,
                        0x83848f2256cb028796d7438373850c356f94d4203a388116ffe62d146235f235 ,
                        0x62f7632905616e71255cf3d12597235fde9622a0c82fbdd2765cc824cb852d2d ,
                        0x467f2e11636a681045c840f345021344337945727dd3711b76605a762c97682c ,
                        0x2ef22e1bb52007063769c85b00a07a83761019a42e7de86a65d4331b62399805 ,
                        0x122a304b8f873fd8c226b3c4306f1389168824f7f5422cb620dca77da2388cf3 ,
                        0x243141a35b37435656b8400a55579e13858662639c417735927a15378a38d109 ,
                        0xbff3638cc7622a983371934fb1f72312767c7bf62ef0947b1c9731c2a58f27a5 ,
                        0x423234258ff23a2e63227126440a177ccb93456a6767d9f711631664914649f3 ,
                        0xd17a1b0957c7e72a07c326ac4852275505b766549c65528aff7553c7e86b3608 ,
                        0x56883850c7e2225118255ce80474c339f6b46a8da28a2d91400c921042e82a28 ,
                        0xd57c539835a3708b5499350d781025d881f3952aadc22968a41f03e4809e3333 ,
                        0x6c723cb9c459ea688900954cbe3522d701729eb33d22b145322931571939ead8 ,
                        0x430ef24956b62fbe984b40d248f0743755384446a82d56126e3729342d063955 ,
                        0xa32a5557250da3745a298c0a0197fa31942eb1195ed7411d141f5e179aee9823 ,
                        0xaa01831d15852b509aadf99ab4f19aa66961acb1795e87109c519b9eb15ab122 ,
                        0x9baf491aa6119aba9910bab97973d241a9a19a73891dbe199a26b17bd5811ace ,
                        0x49100ce69a9f819b9e619ce7721ac9a91aab499cbeb919f1ff219a9bb753cde9 ,
                        0x99a6219aafe91c2bd09ec2e999bb919db1619dcb6910dc449ef08919fab99ffc ,
                        0xd99ffef99f001aa0023aa0045aa07e411ace711bf7791cc7415cc4d551101a9d ,
                        0xd5919aa951a13078a1189aa11abaa11ccaa116faa1141aa2123aa2246a1d06aa ,
                        0x18e9211abb871ca6d39d4da607d5653fd656213d328b0a868f81c623ebc730e1 ,
                        0xd6a337eaa340faa342ca930058a46d3652398aa44974a273911ef7f9a0c971a1 ,
                        0x86672132d623a7d7967d92a554f84a7f10a34df6077a606f1b5aa22d6a1d10ea ,
                        0xa0689aa60ecaa00ca6a66efaa62c5aa2236a3a742aa773caa463611ed919a579 ,
                        0x503fc703208b8975dff6893d335a191a9d6b1a1a8a5a02ceffc12578faa82391 ,
                        0x9ed421a532da4b34025d176a1d0c561f90daa9fd059da911a635b34353b2a562 ,
                        0x8aa8ac719d9ebaaa3df11a93da64803a8eded664aad951bdc1aab81a14af819c ,
                        0x30f8a73ee32398ea9da9daa82396abc6fa13e5811c30f8077ffa207d626fab71 ,
                        0xabc73aadbacaab6009a8c1ba9aaa4aaddcfa139b919c9929452fa29cdbdaade6 ,
                        0xca13a6711a3098329e535d1b949fe71aaf32a145a9b1a53c535d23c819bc29af ,
                        0xfc5a13d456af6aa51f585405d3d9af06db5fea4a96c0ba4107dbb039711c7d5a ,
                        0x23947495dfe9b0164b13a601b05847b1f579b11edb12d4962a0b9baa1f5bb22f ,
                        0xf1ad7b4785d04aac26dbb22ad141112b45aa511c2e5bb328ff819a5a892058e4 ,
                        0x8e36dbb31d01b367d2b13e3bb43fcb7bf848b1f04ab44a2b11a7e10999e9474b ,
                        0x1bb51d61b4c70224492bb558ab10a82935a3009259fbb515711a29a31afb0ab6 ,
                        0x66abb531d6332472b5676bb6adb62057c9b36d3bb74232992f22b7743bb76945 ,
                        0x7e6c9bb759fbb63647397e3bb85c9256d425b4844bb727d253a484b789fbb56f ,
                        0x9b5009f9b8846bb83fe2b8948bb5fa52245899b97e8b974fd4b79e3bb49bcb1f ,
                        0x653bba90bb73938bba738b6ac072baac8bb56cb6bab17bb62ad5b9b5ebb666c2 ,
                        0x63b9dbb642c894bd4bb4245045b81bbc5feb249065bc580b5ca2abbccefbbcd0 ,
                        0x1bbdd23bbdd45bbdd67bbdd89bbddabbbddcdbbddefbbde0047b110101003b
                    End

                    LayoutCachedLeft =9600
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =540
                    TabIndex =16
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2580
                    Top =540
                    Width =2640
                    Height =315
                    TabIndex =4
                    BackColor =15132391
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtsecKeywords"
                    GridlineColor =10921638

                    LayoutCachedLeft =2580
                    LayoutCachedTop =540
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =855
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextAlign =3
                            Left =900
                            Top =555
                            Width =1620
                            Height =315
                            BorderColor =8355711
                            Name ="Label10"
                            Caption ="2nd Keyword:"
                            GridlineColor =10921638
                            LayoutCachedLeft =900
                            LayoutCachedTop =555
                            LayoutCachedWidth =2520
                            LayoutCachedHeight =870
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =93
                    Left =5400
                    Top =540
                    Width =1320
                    Height =300
                    TabIndex =5
                    ForeColor =4210752
                    Name ="btnsecSearch"
                    Caption ="Search"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =5400
                    LayoutCachedTop =540
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =840
                    BackColor =15123357
                    BorderColor =15123357
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =93
                    Left =6900
                    Top =540
                    Width =300
                    Height =300
                    TabIndex =6
                    ForeColor =4210752
                    Name ="Command12"
                    Caption ="Command7"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xddddddddddddddddddddddddddddddddddddddddddd74ddddd44444ddddd47dd ,
                        0xdd4444ddddddd4dddd444dddddddd4dddd44d4ddddddd4dddd4ddd44dddd47dd ,
                        0xdddddddd44447ddddddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Undo Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =6900
                    LayoutCachedTop =540
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =840
                    BackColor =15123357
                    BorderColor =15123357
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CheckBox
                    OverlapFlags =223
                    Left =2220
                    Top =810
                    Width =240
                    TabIndex =7
                    BorderColor =10921638
                    Name ="Check13"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =2220
                    LayoutCachedTop =810
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =1050
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =2450
                            Top =780
                            Width =390
                            Height =315
                            BorderColor =8355711
                            Name ="chMB"
                            Caption ="MB"
                            GridlineColor =10921638
                            LayoutCachedLeft =2450
                            LayoutCachedTop =780
                            LayoutCachedWidth =2840
                            LayoutCachedHeight =1095
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =3060
                    Top =810
                    Width =240
                    TabIndex =8
                    BorderColor =10921638
                    Name ="Check18"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =3060
                    LayoutCachedTop =810
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =1050
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =3290
                            Top =780
                            Width =330
                            Height =315
                            BorderColor =8355711
                            Name ="chRb"
                            Caption ="RB"
                            GridlineColor =10921638
                            LayoutCachedLeft =3290
                            LayoutCachedTop =780
                            LayoutCachedWidth =3620
                            LayoutCachedHeight =1095
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =255
                    Left =3840
                    Top =810
                    Width =240
                    TabIndex =9
                    BorderColor =10921638
                    Name ="Check20"
                    DefaultValue ="0"
                    GridlineColor =10921638

                    LayoutCachedLeft =3840
                    LayoutCachedTop =810
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1050
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =4070
                            Top =780
                            Width =300
                            Height =315
                            BorderColor =8355711
                            Name ="chES"
                            Caption ="SP"
                            GridlineColor =10921638
                            LayoutCachedLeft =4070
                            LayoutCachedTop =780
                            LayoutCachedWidth =4370
                            LayoutCachedHeight =1095
                            ForeTint =100.0
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =255
                    Left =1380
                    Top =630
                    Width =4680
                    Height =570
                    TabIndex =10
                    BorderColor =10921638
                    Name ="Frame26"
                    GridlineColor =10921638

                    LayoutCachedLeft =1380
                    LayoutCachedTop =630
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1200
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =255
                            Left =1560
                            Top =480
                            Width =975
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label27"
                            Caption ="Site Filter"
                            GridlineColor =10921638
                            LayoutCachedLeft =1560
                            LayoutCachedTop =480
                            LayoutCachedWidth =2535
                            LayoutCachedHeight =795
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =4560
                    Top =810
                    Width =315
                    Height =270
                    TabIndex =11
                    BorderColor =10921638
                    Name ="Check28"
                    DefaultValue ="False"
                    GridlineColor =10921638

                    LayoutCachedLeft =4560
                    LayoutCachedTop =810
                    LayoutCachedWidth =4875
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4785
                            Top =780
                            Width =435
                            Height =315
                            BorderColor =8355711
                            Name ="Label29"
                            Caption ="RB2"
                            GridlineColor =10921638
                            LayoutCachedLeft =4785
                            LayoutCachedTop =780
                            LayoutCachedWidth =5220
                            LayoutCachedHeight =1095
                            ForeTint =100.0
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =1500
                    Top =810
                    Width =300
                    TabIndex =12
                    BorderColor =10921638
                    Name ="Check30"
                    DefaultValue ="False"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedTop =810
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =1050
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1725
                            Top =780
                            Width =300
                            Height =315
                            BorderColor =8355711
                            Name ="Label31"
                            Caption ="ES"
                            GridlineColor =10921638
                            LayoutCachedLeft =1725
                            LayoutCachedTop =780
                            LayoutCachedWidth =2025
                            LayoutCachedHeight =1095
                            ForeTint =100.0
                        End
                    End
                End
                Begin Image
                    PictureType =2
                    Left =8940
                    Top =480
                    Width =3420
                    Height =1020
                    Name ="Image39"
                    Picture ="NewMERMLogo"
                    HorizontalAnchor =1

                    LayoutCachedLeft =8940
                    LayoutCachedTop =480
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =1500
                    TabIndex =17
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7860
                    Top =540
                    Width =480
                    Height =420
                    TabIndex =13
                    Name ="Command40"
                    Caption ="Command40"
                    OnClick ="[Event Procedure]"
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

                    LayoutCachedLeft =7860
                    LayoutCachedTop =540
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =960
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =215
                    Left =6420
                    Top =600
                    Width =1560
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label41"
                    Caption ="Export Results:"
                    GridlineColor =10921638
                    LayoutCachedLeft =6420
                    LayoutCachedTop =600
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =915
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =5340
                    Top =840
                    Width =315
                    Height =270
                    TabIndex =14
                    BorderColor =10921638
                    Name ="Check42"
                    DefaultValue ="False"
                    GridlineColor =10921638

                    LayoutCachedLeft =5340
                    LayoutCachedTop =840
                    LayoutCachedWidth =5655
                    LayoutCachedHeight =1110
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =5580
                            Top =780
                            Width =405
                            Height =315
                            BorderColor =8355711
                            Name ="Label43"
                            Caption ="CTS"
                            GridlineColor =10921638
                            LayoutCachedLeft =5580
                            LayoutCachedTop =780
                            LayoutCachedWidth =5985
                            LayoutCachedHeight =1095
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7860
                    Top =1020
                    Width =720
                    Height =390
                    FontSize =14
                    FontWeight =700
                    TabIndex =15
                    Name ="Text44"
                    ControlSource ="=[subMERMList].[Form]![txtRecordCount]"

                    LayoutCachedLeft =7860
                    LayoutCachedTop =1020
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =1410
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6480
                            Top =1080
                            Width =1380
                            Height =315
                            Name ="Label45"
                            Caption ="Record Count:"
                            LayoutCachedLeft =6480
                            LayoutCachedTop =1080
                            LayoutCachedWidth =7860
                            LayoutCachedHeight =1395
                        End
                    End
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
Option Explicit

Private Sub btnSearch_Click()
Dim sql As String
Dim sFilter As String
sFilter = ""

If Me.Check13 Then BuildFilter sFilter, "MERM.Site = 'MB'"
If Me.Check18 Then BuildFilter sFilter, "MERM.Site = 'RB'"
If Me.Check20 Then BuildFilter sFilter, "MERM.Site = 'SP'"
If Me.Check30 Then BuildFilter sFilter, "MERM.Site = 'ES'"
If Me.Check28 Then BuildFilter sFilter, "MERM.Site = 'RB2'"
If Me.Check42 Then BuildFilter sFilter, "MERM.Site = 'CTS'"


sql = "SELECT MERM.EventID, MERM.Site, MERM.Group, MERM.CommonName, MERM.DocumentName, MERM.Asset, MERM.Building, MERM.AQMD_ID, " _
& "MERM.[Assigned To] FROM MERM " _
& "WHERE MERM.Group LIKE '*" & Me.txtKeywords & "*' OR MERM.DocumentName LIKE '*" & Me.txtKeywords & "*' OR MERM.AQMD_ID LIKE '*" & Me.txtKeywords & "*' OR MERM.Building LIKE '*" & Me.txtKeywords & "*' OR MERM.Asset LIKE '*" & Me.txtKeywords & "*' OR MERM.[Assigned To] LIKE '*" & Me.txtKeywords & "*' OR MERM.CommonName LIKE '*" & Me.txtKeywords & "*' " _
& "ORDER BY MERM.[Assigned To] DESC;"

Me.subMERMList.Form.RecordSource = sql
Me.subMERMList.Form.Requery

subMERMList.Form.Filter = sFilter
subMERMList.Form.FilterOn = Len(sFilter) > 0 ' test
'First attempt to count records
'Dim rst As Recordset
'Dim CallCount As Long
'Set rst = CurrentDb.OpenRecordset(subMERMList.Form.RecordSource) ' & IIf(subMERMList.Form.FilterOn, " WHERE " & subMERMList.Form.Filter, ""))
'rst.MoveLast
'CallCount = rst.RecordCount
'Set rst = Nothing
'Me.Text44 = CallCount
End Sub

Sub BuildFilter(ByRef sFilt As String, sAdd As String)
   If Len(sFilt) > 0 Then sFilt = sFilt & " OR "
   sFilt = sFilt & sAdd
End Sub

Private Sub btnsecSearch_Click()
'Dim SQL As String
'SQL = "SELECT MERM.EventID, MERM.Site, MERM.Group, MERM.DocumentName, MERM.Asset, MERM.AQMD_ID, " _
'& "MERM.[Assigned To] FROM MERM " _
'& "WHERE MERM.Site LIKE '*" & Me.txtsecKeywords & "*' OR MERM.Group LIKE '*" & Me.txtsecKeywords & "*' OR MERM.DocumentName LIKE '*" & Me.txtsecKeywords & "*' OR MERM.AQMD_ID LIKE '*" & Me.txtsecKeywords & "*' OR MERM.Asset LIKE '*" & Me.txtsecKeywords & "*' OR MERM.[Assigned To] LIKE '*" & Me.txtsecKeywords & "*' " _
'& "ORDER BY MERM.[Assigned To] DESC;"
'
'Me.subMERMList.Form.RecordSource = SQL
'Me.subMERMList.Form.Requery
MsgBox "This feature is still in development", vbOKOnly
End Sub
Sub BuildFilter1(ByRef nFilt As String, nAdd As String)
   If Len(nFilt) > 0 Then nFilt = nFilt & " OR "
   nFilt = nFilt & nAdd
End Sub
Private Sub Command40_Click()
On Error GoTo errHandler
Dim qdf As DAO.QueryDef
Dim qdfNew As DAO.QueryDef
Dim StrPathFile As String
Dim nFilter As String, nFilt As String
Dim strExport As String
Dim dbs As DAO.Database
Dim rsqs As DAO.Recordset
Dim Site As String
 
nFilter = ""

If Me.Check13 Then BuildFilter1 nFilter, "(sfrmMERMListQry.Site)= 'MB'"
If Me.Check18 Then BuildFilter1 nFilter, "(sfrmMERMListQry.Site)= 'RB'"
If Me.Check20 Then BuildFilter1 nFilter, "(sfrmMERMListQry.Site)= 'SP'"
If Me.Check30 Then BuildFilter1 nFilter, "(sfrmMERMListQry.Site)='ES'"
If Me.Check28 Then BuildFilter1 nFilter, "(sfrmMERMListQry.Site)='RB2'"
If Me.Check42 Then BuildFilter1 nFilter, "(sfrmMERMListQry.Site)='CTS'"
'If Len(nFilt) = 0 Then
'nFilter = "(sfrmMERMListQry.Site) Is Not Null"

Set dbs = CurrentDb

DoCmd.DeleteObject acQuery, "sfrmMERMListQry"
DoCmd.DeleteObject acQuery, "sfrmMERMListQryOut"


Set qdf = CurrentDb.CreateQueryDef("sfrmMERMListQry", Me.subMERMList.Form.RecordSource)
Set qdf = dbs.QueryDefs("sfrmMERMListQry")

If Len(nFilter) > 0 Then
strExport = "SELECT * FROM sfrmMERMListQry " _
& "WHERE (" & nFilter & ")"
Else
If Len(nFilt) = 0 Then
strExport = "SELECT * FROM sfrmMERMListQry "
End If
End If
Set qdfNew = CurrentDb.CreateQueryDef("sfrmMERMListQryOut", strExport)


StrPathFile = CreateObject("WScript.Shell").specialfolders("MyDocuments")
Debug.Print StrPathFile
'where the report is exported to
StrPathFile = StrPathFile & "\" & "User Report.xlsx"
DoCmd.OutputTo acOutputQuery, "sfrmMERMListQryOut", acFormatXLSX, StrPathFile, True, acExportQualityScreen
CurrentDb.QueryDefs.Delete qdfNew.name 'cleanup
    
exitHandler:
Exit Sub

errHandler:
If Err.Number = 7874 Then
Resume Next
Else
MsgBox Err.Number & " - " & Err.Description
Resume exitHandler
End If

End Sub

Sub Command7_Click()
Dim sql As String
Me.Check13 = False
Me.Check18 = False
Me.Check20 = False
Me.Check28 = False
Me.Check30 = False
Me.Check42 = False

Me.txtKeywords = ""
sql = "SELECT MERM.EventID, MERM.Site, MERM.Group, MERM.CommonName, MERM.DocumentName, MERM.Asset, MERM.AQMD_ID, " _
& "MERM.[Assigned To] FROM MERM " _
& "WHERE MERM.Site LIKE '*" & Me.txtKeywords & "*' OR MERM.Group LIKE '*" & Me.txtKeywords & "*' OR MERM.DocumentName LIKE '*" & Me.txtKeywords & "*' OR MERM.AQMD_ID LIKE '*" & Me.txtKeywords & "*' OR MERM.Asset LIKE '*" & Me.txtKeywords & "*' OR MERM.[Assigned To] LIKE '*" & Me.txtKeywords & "*' OR MERM.CommonName LIKE '*" & Me.txtKeywords & "*' " _
& "ORDER BY MERM.[Assigned To] DESC;"

Me.subMERMList.Form.RecordSource = sql
Me.subMERMList.Form.Requery

End Sub
Private Sub Command12_Click()
Dim sql As String
Me.txtsecKeywords = ""
sql = "SELECT MERM.EventID, MERM.Site, MERM.Group, MERM.CommonName, MERM.DocumentName, MERM.Asset, MERM.AQMD_ID, " _
& "MERM.[Assigned To] FROM MERM " _
& "WHERE MERM.Site LIKE '*" & Me.txtsecKeywords & "*' OR MERM.Group LIKE '*" & Me.txtsecKeywords & "*' OR MERM.DocumentName LIKE '*" & Me.txtsecKeywords & "*' OR MERM.AQMD_ID LIKE '*" & Me.txtsecKeywords & "*' OR MERM.Asset LIKE '*" & Me.txtsecKeywords & "*' OR MERM.[Assigned To] LIKE '*" & Me.txtsecKeywords & "*' OR MERM.CommonName LIKE '*" & Me.txtsecKeywords & "*' " _
& "ORDER BY MERM.[Assigned To] DESC;"

Me.subMERMList.Form.RecordSource = sql
Me.subMERMList.Form.Requery

End Sub


Sub Seach()
Dim strSearch, strSearch2 As String
Dim TempFilter, Tempfilter2, Tempfilter3, Tempfilter4, Tempfilter5, Tempfilter6, Tempfilter7 As String
Dim strTempfilter As String
Dim task, strCriteria As String
Dim varTempFilter As TempVars

On Error GoTo ErrorHandler

Me.Refresh
strSearch = Me.txtKeywords.Value
    
    'Checkbox filters; have check boxes for each field chSite, chCodumentName, etc
    If Me.chMB = True Then
        TempFilter = "[Site] Like ""*" & [strSearch] & "*"")"
    Else
        TempFilter = "([blank] Like ""*" & [strSearch] & "*"")"
    End If
    If Me.chRb = True Then
        TempFilter = "[Site] Like ""*" & [strSearch] & "*"")"
    Else
        TempFilter = "([blank] Like ""*" & [strSearch] & "*"")"
    End If
    If Me.chES = True Then
        TempFilter = "[Site] Like ""*" & [strSearch] & "*"")"
    Else
        TempFilter = "([blank] Like ""*" & [strSearch] & "*"")"
    End If
    
    'Second Filter
    If Not IsNull(Me.txtsecKeywords) Then
        strSearch2 = Me.txtsecKeywords.Value
        
    If Me.chMB = True Then
        Tempfilter2 = "([Site] Like ""*" & [strSearch2] & "*"")"
        strCriteria = "(" & Tempfilter2 & "And (" & TempFilter & "))"
      Else
        Tempfilter2 = "([blank] Like ""*" & [strSearch2] & "*"")"
    End If
End If
 ' Me.cmdShowAll.Enabled = True
  
ExitPoint:
    Exit Sub
ErrorHandler:
MsgBox "More Checkboxes may be needed", _
    vbNewLine & "Error Number: " & Err.Number _
    & vbNewLine & "error Description: " & Err.Description _
    , vbExclamation, "Unexpected Error"
    Exit Sub
End Sub
