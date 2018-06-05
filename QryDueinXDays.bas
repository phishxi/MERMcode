SELECT MERM.CommonName, qryEventDates.DocumentName, MERM.[Assigned To], qryEventDates.EventDate, Contacts.Email, [EventDate]-Date() AS NumDay, MERM.Email_Sent_Date, MERM.firstEmail
FROM (Contacts INNER JOIN MERM ON Contacts.ID = MERM.Email) INNER JOIN qryEventDates ON MERM.EventID = qryEventDates.EventID
WHERE ((([EventDate]-Date())>0 And ([EventDate]-Date())<30))
ORDER BY MERM.[Assigned To], [EventDate]-Date();

