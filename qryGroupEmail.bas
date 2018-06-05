SELECT DISTINCT [Contacts Extended].[Contact Name], MERM.Group, Contacts.Email
FROM ([Contacts Extended] INNER JOIN Contacts ON [Contacts Extended].ID = Contacts.ID) INNER JOIN MERM ON [Contacts Extended].[Contact Name] = MERM.[Assigned To];

