SELECT IIf(IsNull([Last Name]),IIf(IsNull([First Name]),[Company],[First Name]),IIf(IsNull([First Name]),[Last Name],[First Name] & " " & [Last Name])) AS [Contact Name], IIf(IsNull([Last Name]),IIf(IsNull([First Name]),[Company],[First Name]),IIf(IsNull([First Name]),[Last Name],[Last Name] & ", " & [First Name])) AS [File As], Contacts.*, Contacts.Admin
FROM Contacts
ORDER BY IIf(IsNull([Last Name]),IIf(IsNull([First Name]),[Company],[First Name]),IIf(IsNull([First Name]),[Last Name],[First Name] & " " & [Last Name])), IIf(IsNull([Last Name]),IIf(IsNull([First Name]),[Company],[First Name]),IIf(IsNull([First Name]),[Last Name],[Last Name] & ", " & [First Name]));

