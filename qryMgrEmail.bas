SELECT MERM.EventID, MERM.CommonName, MERM.DocumentName, MERM.[Assigned To], qryEventDates.EventDate, Contacts.Email, IIf([DelayedSubmission] Is Not Null,([EventDate]-Date())+[DelayedSubmission],[EventDate]-Date()) AS numday, MERM.Priority, Contacts.rcvEmail, MERM.DelayedSubmission, Contacts.MgrName, Contacts.MgrEmail, MERM.DelayedSubmission
FROM (Contacts INNER JOIN MERM ON Contacts.ID = MERM.Email) INNER JOIN qryEventDates ON MERM.EventID = qryEventDates.EventID
WHERE (((IIf([DelayedSubmission] Is Not Null,([EventDate]-Date())+[DelayedSubmission],[EventDate]-Date()))<=0) AND ((Contacts.rcvEmail)=Yes) AND ((Contacts.MgrEmail) Is Not Null))
ORDER BY Contacts.Email, MERM.Priority;

