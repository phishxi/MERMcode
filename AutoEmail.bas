Option Compare Database

'First Email function
'Checks number of days requested for first email against the number of days until the item is due
'Relies on qryfirstEmail query
Sub SendEmail()
 
Dim oOutlook As Outlook.Application
Dim oEmailItem As MailItem
Dim rs As DAO.Recordset
Dim StrItemsdue As String
Dim Names As String
Dim numday As Integer
On Error Resume Next
Err.Clear
 
Set oOutlook = GetObject(, "Outlook.Application")
If Err.Number <> 0 Then
Set oOutlook = New Outlook.Application
End If
CurrentDb.Execute "qryfirstEmail", dbFailOnError

Set rs = CurrentDb.OpenRecordset("SELECT * FROM qryfirstEmail")

If Not (rs.BOF And rs.EOF) Then
 
rs.MoveFirst
 
While (Not rs.EOF)

If rs!rcvEmail = -1 Then
'rs!Email = "Matthew.Kent@ngc.com" _
Or rs!Email = "Renee.Ballinger@ngc.com" _
Or rs!Email = "Helene.Jouin@ngc.com" _
Or rs!Email = "Ron.Frazer@ngc.com" Then

 Set oEmailItem = oOutlook.CreateItem(olMailItem)

With oEmailItem
Names = rs![Assigned To]
Items = DCount("[EventID]", "qryfirstEmail", "[Assigned To] = '" & Names & "'")
.To = rs!Email
If .To = "Helene.Jouin@ngc.com" Then
.CC = rs!MgrEmail
End If

.Subject = "MERM Reminders for " & Items & " Upcoming Tasks"

'Creates a table with records due
StrItemsdue = "<HTML><table font color='white'border='1'style='width:75%'><tr bgcolor='800000'><th> Priority </th><th> Assigned to </th><th> Common Name </th><th> Document Name </th><th> Due Date </th><th> Days until due </th></tr>"
Do While rs.EOF = False And rs!Email = .To And rs![priority] = "1 - Very Critical"

                StrItemsdue = StrItemsdue & "<tr><td align='left'>" & rs("Priority") & "<td align='center'>" & rs("Assigned To") & "</td><td>" & rs("CommonName") & "</td><td>" & rs("DocumentName") & "</td><td>" & rs("EventDate") & "</td><td>" & _
                    "Due in " & rs("NumDay") & " days" & "</td></tr>"
                rs.MoveNext
    If rs!Email <> .To Then
         Exit Do
    End If
Loop

StrItemsdue = StrItemsdue & "</table>"

'creates a new table for non-critical records
StrItemsdue2 = "<table border='1'style='width:75%'><tr bgcolor='a3a3c2'><th> Priority </th><th> Assigned to </th><th> Common Name </th><th> Document Name </th><th> Due Date </th><th> Days until due </th></tr>"
Do While rs!Email = .To And rs!priority <> "1 - Very Critical"
        StrItemsdue2 = StrItemsdue2 & "<tr><td align='left'>" & rs("Priority") & "<td align='center'>" & rs("Assigned To") & "</td><td>" & rs("CommonName") & "</td><td>" & rs("DocumentName") & "</td><td>" & rs("EventDate") & "</td><td>" & _
                    "Due in " & rs("NumDay") & " days" & "</td></tr>"
            
                rs.MoveNext
    If rs!Email <> .To Then
        Exit Do
    End If
Loop

StrItemsdue2 = StrItemsdue2 & "</table></HTML>"

.HTMLBody = "<p>Dear " & Names & ",</p>" & "<p>Below is a list of the MERM items you have coming due:" & StrItemsdue & "<br/>" & StrItemsdue2 & "<br/><a href='\\rsic30-db0016\InjuryandIllnessDatabase\MERM\MERM_fe v2.0.accde'>Please view the MERM to address these items</a>"
.Send
'.Display

End With
Else
rs.MoveNext
End If

Wend

Else
    'MsgBox "No items are pending notification and no emails were sent."
Exit Sub
End If

rs.Close
Set rs = Nothing
 
End Sub
'Second Email function
'Checks number of days requested for second email against the number of days until the item is due
'Relies on qrysecondEmail query
Sub SendEmail2()
 
Dim oOutlook As Outlook.Application
Dim oEmailItem As MailItem
Dim rs As DAO.Recordset
Dim StrItemsdue As String
Dim Names As String
On Error Resume Next
Err.Clear
 
Set oOutlook = GetObject(, "Outlook.Application")
If Err.Number <> 0 Then
Set oOutlook = New Outlook.Application
End If
CurrentDb.Execute "qrysecondEmail", dbFailOnError

Set rs = CurrentDb.OpenRecordset("SELECT * FROM qrysecondEmail")

If Not (rs.BOF And rs.EOF) Then
 
rs.MoveFirst
 
While (Not rs.EOF)

If rs!rcvEmail = -1 Then
'rs!Email = "Matthew.Kent@ngc.com" _
Or rs!Email = "Renee.Ballinger@ngc.com" _
Or rs!Email = "Helene.Jouin@ngc.com" _
Or rs!Email = "Ron.Frazer@ngc.com" Then

 Set oEmailItem = oOutlook.CreateItem(olMailItem)

With oEmailItem
Names = rs![Assigned To]
Items = DCount("[EventID]", "qrysecondEmail", "[Assigned To] = '" & Names & "'")
.To = rs!Email
If .To = "Helene.Jouin@ngc.com" Then
.CC = rs!MgrEmail
End If
.Subject = "Last Reminders for " & Items & " Upcoming MERM Tasks"

'Creates a table with records due
StrItemsdue = "<HTML><table font color='white'border='1'style='width:75%'><tr bgcolor='800000'><th> Priority </th><th> Assigned to </th><th> Common Name </th><th> Document Name </th><th> Due Date </th><th> Days until due </th></tr>"
Do While rs.EOF = False And rs!Email = .To And rs![priority] = "1 - Very Critical"

                StrItemsdue = StrItemsdue & "<tr><td align='left'>" & rs("Priority") & "</td><td align='center'>" & rs("Assigned To") & "</td><td>" & rs("CommonName") & "</td><td>" & rs("DocumentName") & "</td><td>" & rs("EventDate") & "</td><td>" & _
                    "Due in " & rs("NumDay") & " days" & "</td></tr>"
                rs.MoveNext
    If rs!Email <> .To Then
        Exit Do
    End If
Loop

StrItemsdue = StrItemsdue & "</table>"

'creates a new table for non-critical records
StrItemsdue2 = "<table border='1'style='width:75%'><tr bgcolor='a3a3c2'><th> Priority </th><th> Assigned to </th><th> Common Name </th><th> Document Name </th><th> Due Date </th><th> Days until due </th></tr>"
Do While rs!Email = .To And rs!priority <> "1 - Very Critical"
        StrItemsdue2 = StrItemsdue2 & "<tr><td align='left'>" & rs("Priority") & "<td align='center'>" & rs("Assigned To") & "</td><td>" & rs("CommonName") & "</td><td>" & rs("DocumentName") & "</td><td>" & rs("EventDate") & "</td><td>" & _
                    "Due in " & rs("NumDay") & " days" & "</td></tr>"
            
                rs.MoveNext
    If rs!Email <> .To Then
         Exit Do
    
    End If
Loop

StrItemsdue2 = StrItemsdue2 & "</table></HTML>"

.HTMLBody = "<p style='color:red'><b>Last Reminder! </b></p>" & "<p>Hey <b>" & Names & "</b></p>" & "<p> this is your final reminder to complete the following items prior to the specified due date." & StrItemsdue & "<br/>" & StrItemsdue2 & "<br/><a href='\\rsic30-db0016\InjuryandIllnessDatabase\MERM\MERM_fe v2.0.accde'>Please view the MERM to address these items</a>"
.Send
'.Display

End With
Else
rs.MoveNext
End If

Wend

Else
    'MsgBox "No items are pending notification and no emails were sent."
Exit Sub
End If

rs.Close
Set rs = Nothing
 
End Sub
'Manager Email function
'Checks if we are passed due date(-numday)
'Relies on qryMgrEmail query
Sub SendMgrEmail()
 
Dim oOutlook As Outlook.Application
Dim oEmailItem As MailItem
Dim rs As DAO.Recordset
Dim StrItemsdue As String
Dim Names As String
Dim Items As Integer
On Error Resume Next
Err.Clear
 
Set oOutlook = GetObject(, "Outlook.Application")
If Err.Number <> 0 Then
Set oOutlook = New Outlook.Application
End If
CurrentDb.Execute "qryMgrEmail", dbFailOnError

Set rs = CurrentDb.OpenRecordset("SELECT * FROM qryMgrEmail")

If Not (rs.BOF And rs.EOF) Then
 
rs.MoveFirst
 
While (Not rs.EOF)

If rs!rcvEmail = -1 Then
'rs!Email = "Matthew.Kent@ngc.com" _
Or rs!Email = "Renee.Ballinger@ngc.com" _
Or rs!Email = "Mark.Bordelon@ngc.com" _
Or rs!Email = "Doug.Hill@ngc.com" _
Or rs!Email = "Arlen.Fuhrman@ngc.com" _
Or rs!Email = "Ron.Frazer@ngc.com" Then
'Or rs!Email = "Helene.Jouin@ngc.com" _

 Set oEmailItem = oOutlook.CreateItem(olMailItem)

With oEmailItem
Names = rs![Assigned To]
Items = DCount("[EventID]", "qryMgrEmail", "[Assigned To] = '" & Names & "'")
.To = rs!Email
.CC = rs!MgrEmail
.Importance = olImportanceHigh
.Subject = "IMMEDIATE ACTION REQUIRED: You have " & Items & " late MERM items!"

'Creates a table with records due
StrItemsdue = "<HTML><table font color='white'border='1'style='width:75%'><tr bgcolor='800000'><th> Priority </th><th> Assigned to </th><th> Common Name </th><th> Document Name </th><th> Due Date </th><th> Days Late </th></tr>"
Do While rs.EOF = False And rs!Email = .To And rs![priority] = "1 - Very Critical"

                StrItemsdue = StrItemsdue & "<tr><td align='left'>" & rs("Priority") & "</td><td align='center'>" & rs("Assigned To") & "</td><td>" & rs("CommonName") & "</td><td>" & rs("DocumentName") & "</td><td>" & rs("EventDate") & "</td><td>" & _
                    rs("NumDay") & " days past due" & "</td></tr>"
                rs.MoveNext
    If rs!Email <> .To Then
        Exit Do
    End If
Loop

StrItemsdue = StrItemsdue & "</table>"

'creates a new table for non-critical records
StrItemsdue2 = "<table border='1'style='width:75%'><tr bgcolor='a3a3c2'><th> Priority </th><th> Assigned to </th><th> Common Name </th><th> Document Name </th><th> Due Date </th><th> Days Late </th></tr>"
Do While rs!Email = .To And rs!priority <> "1 - Very Critical"
        StrItemsdue2 = StrItemsdue2 & "<tr><td align='left'>" & rs("Priority") & "<td align='center'>" & rs("Assigned To") & "</td><td>" & rs("CommonName") & "</td><td>" & rs("DocumentName") & "</td><td>" & rs("EventDate") & "</td><td>" & _
                    rs("NumDay") & " days past due" & "</td></tr>"
            
                rs.MoveNext
    If rs!Email <> .To Then
         Exit Do
    
    End If
Loop

StrItemsdue2 = StrItemsdue2 & "</table></HTML>"

.HTMLBody = "<p style='color:red'><b>EXTREMELY URGENT! </b></p>" & "<p>The following <b>" & Items & "</b> items are past due <b>" & Names & "</b></p>" & "<p>Please visit the MERM and immediately address these issues." & StrItemsdue & "<br/>" & StrItemsdue2 & "<br/><a href='\\rsic30-db0016\InjuryandIllnessDatabase\MERM\MERM_fe v2.0.accde'>Please view the MERM to address these items</a>"
.Send
'.Display

End With
Else
rs.MoveNext
End If

Wend

Else
    'MsgBox "No items are pending notification and no emails were sent."
Exit Sub
End If

rs.Close
Set rs = Nothing
 
End Sub

'MERMaide Emails
Sub SendMonthlyEmail()

'DoCmd.OpenQuery "qrymonthlyEmail", , acReadOnly
Dim oOutlook As Outlook.Application
Dim oEmailItem As MailItem
Dim rs As DAO.Recordset
Dim StrItemsdue As String
Dim Names As String
Dim monthdue As String
On Error Resume Next
Err.Clear

Set oOutlook = GetObject(, "Outlook.Application")
If Err.Number <> 0 Then
Set oOutlook = New Outlook.Application
End If
CurrentDb.Execute "tempQuery", dbFailOnError

Set rs = CurrentDb.OpenRecordset("SELECT * FROM tempQuery")

If Not (rs.BOF And rs.EOF) Then
 
rs.MoveFirst
monthdue = rs!month
While (Not rs.EOF)

If rs!Email = rs!Email Then

 Set oEmailItem = oOutlook.CreateItem(olMailItem)
With oEmailItem
Names = rs![Assigned To]
.To = rs!Email
.Subject = "MERM Reminders for Upcoming Tasks"

'Creates a table with records due
StrItemsdue = "<HTML><table font color='white'border='1'style='width:75%'><tr bgcolor='800000'><th> Priority </th><th> Assigned to </th><th> Common Name </th><th> Document Name </th><th> Due Date </th><th> Days until due </th></tr>"
Do While rs.EOF = False And rs!Email = .To And rs![priority] = "1 - Very Critical"

                StrItemsdue = StrItemsdue & "<tr><td align='left'>" & rs("Priority") & "<td align='center'>" & rs("Assigned To") & "</td><td>" & rs("CommonName") & "</td><td>" & rs("DocumentName") & "</td><td>" & rs("EventDate") & "</td><td>" & _
                    "Due in " & rs("NumDay") & " days" & "</td></tr>"
                rs.MoveNext
                
Loop

StrItemsdue = StrItemsdue & "</table>"

'creates a new table for non-critical records
StrItemsdue2 = "<table border='1'style='width:75%'><tr bgcolor='a3a3c2'><th> Priority </th><th> Assigned to </th><th> Common Name </th><th> Document Name </th><th> Due Date </th><th> Days until due </th></tr>"
Do While rs!priority <> "1 - Very Critical"
        StrItemsdue2 = StrItemsdue2 & "<tr><td align='left'>" & rs("Priority") & "<td align='center'>" & rs("Assigned To") & "</td><td>" & rs("CommonName") & "</td><td>" & rs("DocumentName") & "</td><td>" & rs("EventDate") & "</td><td>" & _
                    "Due in " & rs("NumDay") & " days" & "</td></tr>"
            
                rs.MoveNext
        If rs!Email <> .To Then
    Exit Do

        End If
Loop

StrItemsdue2 = StrItemsdue2 & "</table></HTML>"

.HTMLBody = "<p>Dear " & Names & ",</p>" & "<p>Below is a list of the MERM items you need to complete during the month of <b>" & monthdue & "</b>.</p>" & StrItemsdue & "<br/>" & StrItemsdue2 & "<br/>"
.Send
'.Display

End With
Else
rs.MoveNext
End If

Wend

Else
    'MsgBox "No items are pending notification and no emails were sent."
Exit Sub
End If
MsgBox "SUCCESS!" & vbCrLf & "Sending monthly emails is complete."
rs.Close
Set rs = Nothing
 
End Sub
'Create Group Emails
Sub SendGroupEmail()

Dim MyDb As DAO.Database
Dim rsEmail As DAO.Recordset
Dim sToName As String
Dim sSubject As String

 
Set MyDb = CurrentDb()
Set rsEmail = MyDb.OpenRecordset("tempQuery2", dbOpenSnapshot)
 
With rsEmail
        .MoveFirst
        Do Until rsEmail.EOF
            If IsNull(.Fields(2)) = False Then
                sToName = sToName & .Fields(2) & ";"
                'sSubject = "Hello World"

            End If
            .MoveNext
        Loop
        On Error Resume Next
        'Set True to False to send without displaying first
        DoCmd.SendObject acSendNoObject, , , _
            sToName, , , sSubject, , True, False

End With
 
Set MyDb = Nothing
Set rsEmail = Nothing

End Sub
'Quality Assurance email; Items that require review
Sub SendQAEmail()
 
Dim oOutlook As Outlook.Application
Dim oEmailItem As MailItem
Dim rs As DAO.Recordset
Dim StrItemsdue As String
Dim Names As String
Dim numday As Integer
Dim Items As Integer
On Error Resume Next
Err.Clear
 
Set oOutlook = GetObject(, "Outlook.Application")
If Err.Number <> 0 Then
Set oOutlook = New Outlook.Application
End If
CurrentDb.Execute "qryQAreview", dbFailOnError

Set rs = CurrentDb.OpenRecordset("SELECT * FROM qryQAreview")

If Not (rs.BOF And rs.EOF) Then
 
rs.MoveFirst
 
While (Not rs.EOF)

If rs!rcvEmail = -1 Then
'rs!QAemail = "Matthew.Kent@ngc.com" _
Or rs!QAemail = "Renee.Ballinger@ngc.com" _
Or rs!QAemail = "Helene.Jouin@ngc.com" _
Or rs!QAemail = "Arlen.Fuhrman@ngc.com" _
Or rs!QAemail = "Ron.Frazer@ngc.com" Then


Set oEmailItem = oOutlook.CreateItem(olMailItem)

With oEmailItem
Names = rs![QA_Person]
Items = DCount("[EventID]", "qryQAreview", "[QA_Person] = '" & Names & "'")
.To = rs!QAemail
'If .To = "Helene.Jouin@ngc.com" Then
'.CC = rs!MgrEmail
'End If

.Subject = Items & " pending MERM QA items in your queue"

'Creates a table with records due
StrItemsdue = "<HTML><table font color='black'border='1'style='width:75%'><tr bgcolor='6699cc'><th> QA Person </th><th> Common Name </th><th> Document Name </th><th> Submission Date </th><th> Days in Queue </th></tr>"
Do While rs.EOF = False And rs!QAemail = .To

                StrItemsdue = StrItemsdue & "<tr><td align='center'>" & rs("QA_Person") & "</td><td>" & rs("CommonName") & "</td><td>" & rs("DocumentName") & "</td><td>" & rs("DateComplete") & "</td><td>" & _
                    "Submitted " & rs("numday") & " days ago" & "</td></tr>"
                rs.MoveNext
    If rs!QAemail <> .To Then
         Exit Do
    End If
Loop

StrItemsdue = StrItemsdue & "</table></HTML>"


.HTMLBody = "<p>Dear " & Names & ",</p>" & "<p>Below is a list of the " & Items & " MERM items you need to QA review:" & StrItemsdue & "<br/><a href='\\rsic30-db0016\InjuryandIllnessDatabase\MERM\MERM_fe v2.0.accde'>Please view the MERM to address these items</a>"
.Send
'.Display

End With
Else
rs.MoveNext
End If

Wend

Else
    'MsgBox "No items are pending notification and no emails were sent."
Exit Sub
End If

rs.Close
Set rs = Nothing
 
End Sub