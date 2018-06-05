Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14640
    DatasheetFontHeight =11
    ItemSuffix =49
    Left =1905
    Top =4515
    Right =15030
    Bottom =12660
    DatasheetGridlinesColor =15062992
    Filter ="[QA_Person]=\"Matt Kent\""
    OrderBy ="[DocumentName]"
    RecSrcDt = Begin
        0x95fcb06de90ce540
    End
    RecordSource ="qryQAreview"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    NavigationCaption ="Occurances"
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin FormHeader
            Height =480
            BackColor =4144959
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Width =900
                    Height =435
                    FontSize =9
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label4"
                    Caption ="QA Person\015\012í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    LayoutCachedWidth =900
                    LayoutCachedHeight =435
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1680
                    Width =870
                    Height =420
                    FontSize =9
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label5"
                    Caption ="Due Date\015\012í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =1680
                    LayoutCachedWidth =2550
                    LayoutCachedHeight =420
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3480
                    Width =1425
                    Height =435
                    FontSize =9
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label31"
                    Caption ="MERM document\015\012í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =3480
                    LayoutCachedWidth =4905
                    LayoutCachedHeight =435
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =10980
                    Width =1110
                    Height =435
                    FontSize =9
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label39"
                    Caption ="Upload Date\015\012í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1
                    LayoutCachedLeft =10980
                    LayoutCachedWidth =12090
                    LayoutCachedHeight =435
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =13320
                    Width =1170
                    Height =480
                    FontSize =9
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label46"
                    Caption ="Days in Queue\015\012í ½í´¼/í ½í´½"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =1
                    LayoutCachedLeft =13320
                    LayoutCachedWidth =14490
                    LayoutCachedHeight =480
                End
            End
        End
        Begin Section
            Height =360
            Name ="Detail"
            AlternateBackColor =15921906
            Begin
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =11580
                    Top =60
                    Width =303
                    Height =180
                    RightMargin =144
                    BorderColor =14211288
                    Name ="InstanceID"
                    ControlSource ="tblSubmissions.InstanceID"

                    LayoutCachedLeft =11580
                    LayoutCachedTop =60
                    LayoutCachedWidth =11883
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1620
                    Width =1326
                    Height =360
                    TabIndex =1
                    RightMargin =144
                    BorderColor =14211288
                    Name ="EventDate"
                    ControlSource ="EventDate"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001010000 ,
                        0x1f497d00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0049007300430061006e006e00650064005d0020004900730020004e006f00 ,
                        0x740020004e0075006c006c0000000000
                    End
                    ShowDatePicker =0

                    LayoutCachedLeft =1620
                    LayoutCachedWidth =2946
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010100001f497d00ffffff00160000005b00 ,
                        0x49007300430061006e006e00650064005d0020004900730020004e006f007400 ,
                        0x20004e0075006c006c00000000000000000000000000000000000000000000
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =4212
                    Top =60
                    Width =360
                    TabIndex =2
                    Name ="Check32"
                    ControlSource ="Complete"
                    DefaultValue ="=False"

                    LayoutCachedLeft =4212
                    LayoutCachedTop =60
                    LayoutCachedWidth =4572
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =13260
                    Top =60
                    Width =1251
                    Height =300
                    TabIndex =3
                    RightMargin =144
                    BorderColor =14211288
                    Name ="Text35"
                    ControlSource ="EventID"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001010000 ,
                        0x1f497d00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0049007300430061006e006e00650064005d0020004900730020004e006f00 ,
                        0x740020004e0075006c006c0000000000
                    End
                    ShowDatePicker =0

                    LayoutCachedLeft =13260
                    LayoutCachedTop =60
                    LayoutCachedWidth =14511
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010100001f497d00ffffff00160000005b00 ,
                        0x49007300430061006e006e00650064005d0020004900730020004e006f007400 ,
                        0x20004e0075006c006c00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    FontUnderline = NotDefault
                    IsHyperlink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3000
                    Width =2880
                    Height =360
                    TabIndex =4
                    ForeColor =16711680
                    Name ="Text37"
                    ControlSource ="DocumentName"
                    OnClick ="[Event Procedure]"
                    HorizontalAnchor =2

                    LayoutCachedLeft =3000
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9960
                    Width =960
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =5026082
                    Name ="Command36"
                    Caption ="Approve"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe UI"
                    HorizontalAnchor =1

                    LayoutCachedLeft =9960
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =360
                    HoverForeColor =5026082
                    PressedForeColor =5026082
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =12840
                    Width =660
                    Height =300
                    TabIndex =6
                    Name ="Text41"
                    ControlSource ="SubmitID"

                    LayoutCachedLeft =12840
                    LayoutCachedWidth =13500
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10980
                    Width =2706
                    Height =360
                    TabIndex =7
                    RightMargin =144
                    BorderColor =14211288
                    Name ="Text40"
                    ControlSource ="DateComplete"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001010000 ,
                        0x1f497d00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0049007300430061006e006e00650064005d0020004900730020004e006f00 ,
                        0x740020004e0075006c006c0000000000
                    End
                    HorizontalAnchor =1
                    ShowDatePicker =0

                    LayoutCachedLeft =10980
                    LayoutCachedWidth =13686
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010100001f497d00ffffff00160000005b00 ,
                        0x49007300430061006e006e00650064005d0020004900730020004e006f007400 ,
                        0x20004e0075006c006c00000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =60
                    Width =1560
                    Height =360
                    TabIndex =8
                    Name ="Text43"
                    ControlSource ="QA_Person"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =0
                    Left =8940
                    Width =960
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    ForeColor =1643706
                    Name ="Command43"
                    Caption ="Reject"
                    OnClick ="[Event Procedure]"
                    FontName ="Segoe"
                    HorizontalAnchor =1

                    LayoutCachedLeft =8940
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =360
                    HoverForeColor =1643706
                    PressedForeColor =1643706
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13560
                    Width =840
                    Height =360
                    TabIndex =10
                    Name ="Text44"
                    ControlSource ="numday"
                    HorizontalAnchor =1

                    LayoutCachedLeft =13560
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =360
                End
            End
        End
        Begin FormFooter
            Height =315
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Width =960
                    Height =315
                    Name ="Text47"
                    ControlSource ="=IIf(Count(*)>=1,Count(*),0)"

                    LayoutCachedLeft =540
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =315
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
Option Explicit
Private Const conMod = "Form_frmReviewSub1"

Private Sub Command36_Click()
Dim db1 As Database
Dim rs As DAO.Recordset
Dim FS As Object
Dim fd As FileDialog, fileName As String
On Error GoTo Err_Handler

Dim Msg, Style, Title, Response, MyString
Msg = "Do you confirm this item been reviewed?"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Title = "Approve Submission?"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Yes"    ' Perform some action.

 
'If MsgBox("Are you sure you wish to mark this item as 'Reviewed'", vbOKCancel) = vbOK Then

'Write a new record to the Review table
Set db1 = CurrentDb
Set rs = db1.OpenRecordset("tblReview")
rs.AddNew
rs("SubmitID").Value = Me.[SubmitID]
rs("EventID").Value = Me.EventID
rs("InstanceID").Value = Me.InstanceID
rs("DocumentName").Value = Me.DocumentName
rs("EventDate").Value = Me.EventDate
rs("Complete").Value = True

rs.Update

Forms!MERMHome!frmReviewSub1.Requery
  
Else    ' User chose No.
    MyString = "No"    ' Perform some action.
End If
Exit Sub

Exit_Handler:
    Exit Sub
Err_Handler:
Set fd = Nothing
    Call LogError(Err.Number, Err.Description, conMod)
    Resume Exit_Handler
'End Function
End Sub


Private Sub Command43_Click()
'Reject Button
    'Creates an email to be sent to the person who submitted the record
    'puts the item back to it's original state and requires the user to either "skip" or resubmit document.
    
Dim oOutlook As Outlook.Application
Dim oEmailItem As MailItem
Dim rs As DAO.Recordset
Dim Names As String
Dim Emails As String
Dim fd As FileDialog, fileName As String
On Error Resume Next
Err.Clear
Dim Msg, Style, Title, Response, MyString

Msg = "Are you sure you want to reject this submittal?"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Title = "Reject this submission?"
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Yes"    ' Perform some action.
 
Set oOutlook = GetObject(, "Outlook.Application")
If Err.Number <> 0 Then
Set oOutlook = New Outlook.Application
End If

CurrentDb.Execute "qryQAreview", dbFailOnError

Set oEmailItem = oOutlook.CreateItem(olMailItem)
With oEmailItem
        Names = [Assigned To]
        Emails = [newEmail]
        .To = Emails
        .Subject = "MERM Item Rejected - Immediate Attention Required"
        
        .HTMLBody = "<p style='font-family:Calibri'>Dear " & Names & ",<br/> QA Reviewer, " & [QA_Person] & ", has rejected and the following MERM item: <b>" & DocumentName & "</b>" _
        & "<p style='font-family:Calibri'>Your item has been returned to its original state and it listed on your task list.</p>" _
        & " <a href='\\rsic30-db0016\InjuryandIllnessDatabase\MERM\MERM_fe v2.0.accde'>Please view the MERM to address this issue and resubmit your documentation</a>" & "</b><br/><p style='font-family:Calibri'>Additional notes regarding rejection: </p><br/><br/> "
        '.Send
        .Display
End With


Dim eventlog As String, submitlog As String
Dim i As Variant
i = -1
submitlog = "DELETE FROM tblSubmissions where InstanceID = " & (InstanceID) & " AND EventID = " & Me.EventID & ";"
eventlog = "DELETE FROM tblEventException where InstanceID = " & (InstanceID) & " AND EventID = " & Me.EventID & ";"


DoCmd.SetWarnings False
DoCmd.RunSQL eventlog
DoCmd.RunSQL submitlog
DoCmd.SetWarnings True
DoCmd.Requery
Me.Requery

Else    ' User chose No.
    MyString = "No"    ' Perform some action.
        
End If
Exit Sub

Exit_Handler:
    Exit Sub
Err_Handler:
Set fd = Nothing
    Call LogError(Err.Number, Err.Description, conMod)
    Resume Exit_Handler
'End Function

End Sub

Private Sub Label39_Click()
Me.OrderByOn = True

If Me.OrderBy = "[DateComplete] DESC" Then
    Me.OrderBy = "[DateComplete] ASC"
Else
    Me.OrderBy = "[DateComplete] DESC"
End If

Me.Refresh
End Sub

Private Sub Label4_Click()
Me.OrderByOn = True

If Me.OrderBy = "[QA_Person] DESC" Then
    Me.OrderBy = "[QA_Person] ASC"
Else
    Me.OrderBy = "[QA_Person] DESC"
End If

Me.Refresh
End Sub

Private Sub Label46_Click()
Me.OrderByOn = True

If Me.OrderBy = "[numday] DESC" Then
    Me.OrderBy = "[numday] ASC"
Else
    Me.OrderBy = "[numday] DESC"
End If

Me.Refresh
End Sub

Private Sub Label5_Click()
Me.OrderByOn = True

If Me.OrderBy = "[EventDate] DESC" Then
    Me.OrderBy = "[EventDate] ASC"
Else
    Me.OrderBy = "[EventDate] DESC"
End If

Me.Refresh
End Sub
Private Sub Label31_Click()
Me.OrderByOn = True

If Me.OrderBy = "[DocumentName] DESC" Then
    Me.OrderBy = "[DocumentName] ASC"
Else
    Me.OrderBy = "[DocumentName] DESC"
End If

Me.Refresh
End Sub

Private Sub Text37_Click()

DoCmd.OpenForm "EquipForm", , , "DocumentName = '" & Me.DocumentName & "'", , acHidden

Dim SharepointAddress As String

SharepointAddress = Forms![EquipForm]![Attachment] & "\"
Shell "C:\WINDOWS\explorer.exe " & SharepointAddress & "", vbNormalFocus


End Sub
