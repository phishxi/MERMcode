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
    Width =8400
    DatasheetFontHeight =11
    ItemSuffix =23
    Left =1710
    Top =4815
    Right =13095
    Bottom =9900
    DatasheetGridlinesColor =0
    RecSrcDt = Begin
        0xbaea3b344b15e540
    End
    RecordSource ="SELECT QryAllMERMLate.Site, QryAllMERMLate.DocumentName, QryAllMERMLate.[Assigne"
        "d To], QryAllMERMLate.numday, QryAllMERMLate.Priority FROM QryAllMERMLate ORDER "
        "BY QryAllMERMLate.[Assigned To];"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
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
            Height =480
            BackColor =15064278
            Name ="FormHeader"
            AlternateBackShade =95.0
            BackTint =20.0
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =4860
                    Top =60
                    Width =390
                    FontSize =9
                    ForeColor =16777215
                    Name ="Command24"
                    Caption ="Export to Excel"
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
                    LeftPadding =105
                    TopPadding =60
                    RightPadding =120
                    BottomPadding =165
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    CursorOnHover =1
                    LayoutCachedLeft =4860
                    LayoutCachedTop =60
                    LayoutCachedWidth =5250
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =1
                    ForeShade =100.0
                    UseTheme =1
                    Shape =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    HoverColor =-2147483607
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =36
                    QuickStyleMask =-49
                    Overlaps =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =215
                    Left =4200
                    Top =120
                    Width =660
                    Height =285
                    FontSize =10
                    Name ="Label23"
                    Caption ="Export:"
                    HorizontalAnchor =1
                    LayoutCachedLeft =4200
                    LayoutCachedTop =120
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =405
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =5280
                    Top =60
                    Width =390
                    FontSize =9
                    TabIndex =1
                    ForeColor =16777215
                    Name ="Command25"
                    Caption ="Export to Excel"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddd87111ddddddd87111111111118d1111111111fff818 ,
                        0x1111111111fffff111f8118f1111118111ff11ff11fff811118ffff811fffff1 ,
                        0x111f88f111111181111f11f111fff8111118888111fffff11111ff1111111181 ,
                        0x1111881111fff8111111111111ffff811111111111fff818d87111111111118d ,
                        0xddddd87111dddddd000000000000000000000000000000000000000000000000 ,
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
                    LeftPadding =105
                    TopPadding =60
                    RightPadding =120
                    BottomPadding =165
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    CursorOnHover =1
                    LayoutCachedLeft =5280
                    LayoutCachedTop =60
                    LayoutCachedWidth =5670
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =1
                    ForeShade =100.0
                    UseTheme =1
                    Shape =1
                    BackColor =15921906
                    BackThemeColorIndex =1
                    BackShade =95.0
                    BorderThemeColorIndex =0
                    BorderShade =100.0
                    HoverColor =-2147483607
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeThemeColorIndex =1
                    PressedForeThemeColorIndex =1
                    Shadow =-1
                    QuickStyle =36
                    QuickStyleMask =-49
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1380
                    Width =525
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label9"
                    Caption ="Site í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =1380
                    LayoutCachedWidth =1905
                    LayoutCachedHeight =480
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2160
                    Width =1965
                    Height =465
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label11"
                    Caption ="Document Name \015\012í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =2160
                    LayoutCachedWidth =4125
                    LayoutCachedHeight =465
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =5820
                    Width =1620
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label14"
                    Caption ="Assigned Person    \015\012í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =5820
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =480
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Width =1185
                    Height =480
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label17"
                    Caption ="Priority Level\015\012 í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedWidth =1185
                    LayoutCachedHeight =480
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =7380
                    Width =960
                    Height =465
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="Label15"
                    Caption ="Days Late í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =7380
                    LayoutCachedWidth =8340
                    LayoutCachedHeight =465
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            Height =300
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackShade =95.0
            Begin
                Begin TextBox
                    FontUnderline = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1980
                    Width =4140
                    FontSize =9
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =12673797
                    Name ="Text5"
                    ControlSource ="DocumentName"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    HorizontalAnchor =2

                    LayoutCachedLeft =1980
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =240
                    DisplayAsHyperlink =1
                    ForeTint =100.0
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =6120
                    Width =1560
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text8"
                    ControlSource ="Assigned To"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =6120
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7680
                    Width =600
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text16"
                    ControlSource ="numday"
                    GridlineColor =10921638
                    HorizontalAnchor =1

                    LayoutCachedLeft =7680
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Width =1620
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text21"
                    ControlSource ="Priority"
                    ConditionalFormat = Begin
                        0x01000000dc000000030000000000000002000000000000001400000001010100 ,
                        0xba141900ffffff00000000000200000015000000290000000100000000000000 ,
                        0xffffff0000000000020000002a0000003d0000000100000072727200ffffff00 ,
                        0x2200310020002d00200056006500720079002000430072006900740069006300 ,
                        0x61006c002200000000002200320020002d002000480069006700680020005000 ,
                        0x720069006f0072006900740079002200000000002200330020002d0020004c00 ,
                        0x6f00770020005000720069006f007200690074007900220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedWidth =1620
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x010003000000000000000200000001010100ba141900ffffff00130000002200 ,
                        0x310020002d002000560065007200790020004300720069007400690063006100 ,
                        0x6c00220000000000000000000000000000000000000000000000000000020000 ,
                        0x000100000000000000ffffff00130000002200320020002d0020004800690067 ,
                        0x00680020005000720069006f0072006900740079002200000000000000000000 ,
                        0x00000000000000000000000000000000020000000100000072727200ffffff00 ,
                        0x120000002200330020002d0020004c006f00770020005000720069006f007200 ,
                        0x6900740079002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =1560
                    Width =420
                    FontSize =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text3"
                    ControlSource ="Site"
                    GridlineColor =10921638

                    LayoutCachedLeft =1560
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =240
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


Private Sub Command24_Click()
Dim lngColumn As Long
Dim xlx As Object, xlw As Object, xls As Object, xlc As Object
Dim dbs As DAO.Database
Dim rst As DAO.Recordset
Dim strPathFileName As String, strWorksheetName As String
Dim strRecordsetDataSource As String
Dim blnEXCEL As Boolean, blnHeaderRow As Boolean

blnEXCEL = False
    Dim answer As Integer
answer = MsgBox("Are you sure you want to export the list to Excel?", vbYesNo + vbQuestion, "Export To Excel")
If answer = vbYes Then



' Replace C:\Filename.xls with the actual path and filename
' that will be used to save the new EXCEL file into which you
' will write the data
strPathFileName = "C:\Filename.xls"

' Replace QueryOrTableName with the real name of the table or query
' whose data are to be written into the worksheet
strRecordsetDataSource = "qryAllMERMLate"

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
Else
'do nothing
End If

End Sub

Private Sub Command25_Click()
'Access Export button opens subform query
DoCmd.OpenQuery "qryAllMERMLate"

End Sub



Private Sub Label11_Click()

Me.OrderByOn = True

If Me.OrderBy = "[DocumentName] DESC" Then
    Me.OrderBy = "[DocumentName] ASC"
Else
    Me.OrderBy = "[DocumentName] DESC"
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

Private Sub Label15_Click()
Me.OrderByOn = True

If Me.OrderBy = "[numday] DESC" Then
    Me.OrderBy = "[numday] ASC"
Else
    Me.OrderBy = "[numday] DESC"
End If

Me.Refresh
End Sub
Private Sub Label17_Click()
Me.OrderByOn = True

If Me.OrderBy = "[Priority] DESC" Then
    Me.OrderBy = "[Priority] ASC"
Else
    Me.OrderBy = "[Priority] DESC"
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
