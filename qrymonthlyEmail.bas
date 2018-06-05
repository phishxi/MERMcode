SELECT MERM.CommonName, qryEventDates.DocumentName, MERM.[Assigned To], qryEventDates.EventDate, Contacts.Email, [EventDate]-Date() AS numday, MERM.Priority, Format([EventDate],"yyyy") AS [Year], Format([EventDate],"mmmm") AS [Month]
FROM (Contacts INNER JOIN MERM ON Contacts.ID = MERM.Email) INNER JOIN qryEventDates ON MERM.EventID = qryEventDates.EventID
WHERE (((Format([EventDate],"yyyy"))=IIf(Month(Date())>=12,Format(Now(),"yyyy")+1,Format(Now(),"yyyy"))))
ORDER BY Contacts.Email, MERM.Priority;

