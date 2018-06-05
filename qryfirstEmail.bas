SELECT MERM.CommonName, qryEventDates.DocumentName, MERM.[Assigned To], qryEventDates.EventDate, Contacts.Email, [EventDate]-Date() AS numday, MERM.firstEmail, ([EventDate]-Date())-[firstEmail] AS email1, Contacts.rcvEmail, MERM.Priority, Contacts.MgrEmail
FROM (Contacts INNER JOIN MERM ON Contacts.ID = MERM.Email) INNER JOIN qryEventDates ON MERM.EventID = qryEventDates.EventID
WHERE (((([EventDate]-Date())-[firstEmail])=0) AND ((Contacts.rcvEmail)=Yes))
ORDER BY Contacts.Email, MERM.Priority;

