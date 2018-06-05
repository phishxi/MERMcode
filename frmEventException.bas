Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =204
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =7992
    DatasheetFontHeight =11
    ItemSuffix =15
    Left =28710
    Top =2505
    Right =-15391
    Bottom =15150
    DatasheetGridlinesColor =15062992
    OnUnload ="[Event Procedure]"
    AfterDelConfirm ="[Event Procedure]"
    Filter ="(EventID = 11) AND (InstanceID = 1)"
    RecSrcDt = Begin
        0x2460aff26bffe440
    End
    RecordSource ="SELECT tblEventException.*, MERM.DocumentName, IIf([MERM].[PeriodTypeID] Is Null"
        ",Null,DateAdd([MERM].[PeriodTypeID],[tblEventException].[InstanceID]*[MERM].[Per"
        "iodFreq],[MERM].[EventStart])) AS UsualDate, tblEventException.InstanceComment, "
        "tblEventException.ActualHours FROM MERM INNER JOIN tblEventException ON MERM.Eve"
        "ntID = tblEventException.EventID;"
    Caption ="Manage an exception to a recurring event"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
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
            Height =1080
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4392
                    Top =144
                    Width =1008
                    Height =315
                    ColumnWidth =1395
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =15527148
                    Name ="InstanceID"
                    ControlSource ="InstanceID"
                    StatusBarText ="Instance number. Zero for original. Required. Combination of EventID+CountID mus"
                        "t be unique."

                    LayoutCachedLeft =4392
                    LayoutCachedTop =144
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =459
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3456
                            Top =144
                            Width =915
                            Height =315
                            Name ="Label0"
                            Caption ="Instance:"
                            LayoutCachedLeft =3456
                            LayoutCachedTop =144
                            LayoutCachedWidth =4371
                            LayoutCachedHeight =459
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1944
                    Top =576
                    Width =5760
                    Height =330
                    ColumnWidth =2805
                    ColumnOrder =2
                    TabIndex =2
                    BackColor =15527148
                    Name ="EventDescrip"
                    ControlSource ="DocumentName"
                    StatusBarText ="Description of what this event is."

                    LayoutCachedLeft =1944
                    LayoutCachedTop =576
                    LayoutCachedWidth =7704
                    LayoutCachedHeight =906
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =600
                            Width =1695
                            Height =315
                            Name ="Label4"
                            Caption ="Document Name:"
                            LayoutCachedLeft =180
                            LayoutCachedTop =600
                            LayoutCachedWidth =1875
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1944
                    Top =144
                    Width =1008
                    Height =315
                    ColumnOrder =0
                    BackColor =15527148
                    Name ="EventID"
                    ControlSource ="EventID"
                    StatusBarText ="Relates to tblEvent.EventID. Required."
                    DefaultValue ="1"

                    LayoutCachedLeft =1944
                    LayoutCachedTop =144
                    LayoutCachedWidth =2952
                    LayoutCachedHeight =459
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            TextAlign =3
                            Left =288
                            Top =144
                            Width =1545
                            Height =315
                            Name ="Label5"
                            Caption ="Event ID:"
                            LayoutCachedLeft =288
                            LayoutCachedTop =144
                            LayoutCachedWidth =1833
                            LayoutCachedHeight =459
                        End
                    End
                End
            End
        End
        Begin Section
            Height =2679
            Name ="Detail"
            Begin
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    AccessKey =82
                    IMESentenceMode =3
                    Left =1941
                    Top =144
                    Width =1152
                    Height =315
                    ColumnWidth =1635
                    Name ="InstanceDate"
                    ControlSource ="InstanceDate"
                    StatusBarText ="Date this is rescheduled to."
                    UnicodeAccessKey =82

                    LayoutCachedLeft =1941
                    LayoutCachedTop =144
                    LayoutCachedWidth =3093
                    LayoutCachedHeight =459
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =285
                            Top =144
                            Width =1545
                            Height =315
                            Name ="Label1"
                            Caption ="&Reschedule to:"
                            LayoutCachedLeft =285
                            LayoutCachedTop =144
                            LayoutCachedWidth =1830
                            LayoutCachedHeight =459
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1941
                    Top =778
                    ColumnWidth =1275
                    TabIndex =2
                    Name ="IsCanned"
                    ControlSource ="IsCanned"
                    StatusBarText ="Check the box if this instance is cancelled. Required."
                    DefaultValue ="0"

                    LayoutCachedLeft =1941
                    LayoutCachedTop =778
                    LayoutCachedWidth =2201
                    LayoutCachedHeight =1018
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =285
                            Top =778
                            Width =1545
                            Height =315
                            Name ="Label2"
                            Caption ="Cancel this one:"
                            LayoutCachedLeft =285
                            LayoutCachedTop =778
                            LayoutCachedWidth =1830
                            LayoutCachedHeight =1093
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1896
                    Top =1500
                    Width =5760
                    Height =1125
                    ColumnWidth =5715
                    TabIndex =3
                    Name ="InstanceComment"
                    ControlSource ="tblEventException.InstanceComment"
                    StatusBarText ="Why this was cancelled/rescheduled. 255-char max."

                    LayoutCachedLeft =1896
                    LayoutCachedTop =1500
                    LayoutCachedWidth =7656
                    LayoutCachedHeight =2625
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =240
                            Top =1500
                            Width =1545
                            Height =315
                            Name ="Label3"
                            Caption ="Comment:"
                            LayoutCachedLeft =240
                            LayoutCachedTop =1500
                            LayoutCachedWidth =1785
                            LayoutCachedHeight =1815
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5688
                    Top =144
                    Width =1152
                    Height =315
                    TabIndex =1
                    BackColor =15527148
                    Name ="UsualDate"
                    ControlSource ="UsualDate"

                    LayoutCachedLeft =5688
                    LayoutCachedTop =144
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =459
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =204
                            Left =3456
                            Top =144
                            Width =2175
                            Height =285
                            FontSize =10
                            ForeColor =12349952
                            Name ="Label6"
                            Caption ="instead of the usual date:"
                            LayoutCachedLeft =3456
                            LayoutCachedTop =144
                            LayoutCachedWidth =5631
                            LayoutCachedHeight =429
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =2229
                    Top =778
                    Width =5190
                    Height =285
                    FontSize =10
                    ForeColor =12349952
                    Name ="Label7"
                    Caption ="(Check the box if this instance of the sequence will not occur.)"
                    LayoutCachedLeft =2229
                    LayoutCachedTop =778
                    LayoutCachedWidth =7419
                    LayoutCachedHeight =1063
                End
                Begin Label
                    OverlapFlags =87
                    TextFontCharSet =204
                    TextAlign =2
                    Left =285
                    Top =461
                    Width =1545
                    Height =314
                    ForeColor =12349952
                    Name ="Label11"
                    Caption ="OR"
                    LayoutCachedLeft =285
                    LayoutCachedTop =461
                    LayoutCachedWidth =1830
                    LayoutCachedHeight =775
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7596
                    Top =420
                    TabIndex =4
                    Name ="Check12"
                    ControlSource ="Complete?"
                    StatusBarText ="Check the box if this instance is cancelled. Required."
                    DefaultValue ="0"

                    LayoutCachedLeft =7596
                    LayoutCachedTop =420
                    LayoutCachedWidth =7856
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =5940
                            Top =420
                            Width =1545
                            Height =315
                            ForeColor =12566463
                            Name ="Label13"
                            Caption ="Complete?:"
                            LayoutCachedLeft =5940
                            LayoutCachedTop =420
                            LayoutCachedWidth =7485
                            LayoutCachedHeight =735
                            ForeShade =75.0
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1980
                    Top =1140
                    Width =720
                    Height =315
                    TabIndex =5
                    BackColor =15523798
                    Name ="Text12"
                    ControlSource ="tblEventException.ActualHours"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =1140
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =1455
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1140
                            Width =1560
                            Height =315
                            Name ="Label14"
                            Caption ="Hours Spent:"
                            LayoutCachedLeft =300
                            LayoutCachedTop =1140
                            LayoutCachedWidth =1860
                            LayoutCachedHeight =1455
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1224
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =68
                    Left =1944
                    Top =144
                    Width =4104
                    ForeColor =3751056
                    Name ="cmdDelete"
                    Caption ="&Delete this entry (Revert to usual date.)"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =68

                    LayoutCachedLeft =1944
                    LayoutCachedTop =144
                    LayoutCachedWidth =6048
                    LayoutCachedHeight =504
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =79
                    Left =1440
                    Top =720
                    TabIndex =1
                    Name ="cmdOk"
                    Caption ="&Ok"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =79

                    LayoutCachedLeft =1440
                    LayoutCachedTop =720
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =1080
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =5112
                    Top =720
                    TabIndex =2
                    Name ="cmdCancel"
                    Caption ="&Cancel"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =67

                    LayoutCachedLeft =5112
                    LayoutCachedTop =720
                    LayoutCachedWidth =6552
                    LayoutCachedHeight =1080
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private mbHasChanged As Boolean
Private Const conMod = "frmEventException"

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler
    'Purpose:   Undo and close
    
    If Me.Dirty Then
        Me.Undo
    End If
    
    DoCmd.Close acForm, Me.name
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".cmdCancel_Click")
    Resume Exit_Handler
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Err_Handler
    'Purpose:   Cancel new entry, or delete existing entry, and close form.
    
    If Me.Dirty Then
        Me.Undo
    End If
    
    If Not Me.NewRecord Then
        RunCommand acCmdDeleteRecord
    End If
    
    DoCmd.Close acForm, Me.name
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".cmdDelete_Click")
    Resume Exit_Handler
End Sub

Private Sub cmdOk_Click()
On Error GoTo Err_Handler
    'Purpose:   Save and close.
    
    If IsCanned = -1 Then 'And Me.Dirty Then
      Me.Requery
      'Me.Dirty = False
    End If

    DoCmd.Close acForm, Me.name

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".cmdOk_Click")
    Resume Exit_Handler
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    If IsCanned = -1 And Status = acDeleteOK Then
        mbHasChanged = True
        
    End If
End Sub

Private Sub Form_AfterUpdate()
If IsCanned = -1 Then
    mbHasChanged = True
End If
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
    'Purpose:
    Dim lngEventID As Long
    
    If Me.FilterOn Then
        lngEventID = ParseLongFromFilter(Me.Filter, "EventID")
        If lngEventID <> 0& Then
            Me.EventID.DefaultValue = lngEventID
        End If
        Me.InstanceID.DefaultValue = ParseLongFromFilter(Me.Filter, "InstanceID")
    End If
    Me.Requery
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".Form_Load")
    Resume Exit_Handler
End Sub

Private Function ParseLongFromFilter(strWhere As String, strField As String) As Long
    Dim lngPos As Long
    Const strcCharList = ")] ="
    
    lngPos = InStr(strWhere, strField)
    If lngPos > 0& Then
        ParseLongFromFilter = Val(StripChar(Mid$(strWhere, lngPos + Len(strField) + 1&), strcCharList))
    End If
End Function

Private Function StripChar(ByVal strSource, strCharList) As String
    Dim i As Long
    
    For i = 1 To Len(strCharList)
        strSource = Replace(strSource, Mid(strCharList, i, 1), vbNullString)
    Next
    StripChar = strSource
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
    'Purpose:   Requery the previous form, and return the the same record.
    Dim strWhere As String
    Dim frm As Form
    Dim rs As DAO.Recordset
    
    If CurrentProject.AllForms("EquipForm").IsLoaded Then
        Set frm = Forms!EquipForm!frmEventSub.Form
        If Not frm.NewRecord Then
            strWhere = "(EventID = " & frm!EventID & ") AND (InstanceID = " & frm!InstanceID & ")"
            frm.Requery
            Set rs = frm.RecordsetClone
            rs.FindFirst strWhere
            If Not rs.NoMatch Then
                frm.Bookmark = rs.Bookmark
            End If
        End If
    End If
    
Exit_Handler:
    Set rs = Nothing
    Set frm = Nothing
    Exit Sub
    
Err_Handler:
    Call LogError(Err.Number, Err.Description, conMod & ".Form_Unload")
    Resume Exit_Handler
End Sub
