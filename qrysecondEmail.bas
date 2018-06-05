SELECT MERM.CommonName, qryEventDates.DocumentName, MERM.[Assigned To], qryEventDates.EventDate, Contacts.Email, [EventDate]-Date() AS numday, ([EventDate]-Date())-[secondEmail] AS email2, MERM.secondEmail, MERM.Priority, Contacts.rcvEmail, Contacts.MgrEmail
FROM (Contacts INNER JOIN MERM ON Contacts.ID = MERM.Email) INNER JOIN qryEventDates ON MERM.EventID = qryEventDates.EventID
WHERE (((([EventDate]-Date())-[secondEmail])=0) AND ((Contacts.rcvEmail)=Yes))
ORDER BY Contacts.Email, MERM.Priority;

