SELECT [Contacts Extended].ID, qryEventDates.EventID, qryEventDates.InstanceID, MERM.DocumentName, MERM.[Assigned To], qryEventDates.EventDate, MERM.Email, [EventDate]-Date() AS numday, Format([EventDate],"yyyy") AS [Year], MERM.CommonName, MERM.QA_Person
FROM ([Contacts Extended] INNER JOIN (MERM INNER JOIN qryEventDates ON MERM.EventID = qryEventDates.EventID) ON [Contacts Extended].[Contact Name] = MERM.[Assigned To]) INNER JOIN Contacts ON [Contacts Extended].ID = Contacts.ID
ORDER BY qryEventDates.EventDate, [EventDate]-Date();

