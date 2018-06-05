SELECT tblSubmissions.EventID, ([tblReview].[EventDate]-[tblSubmissions].[DateComplete]+7) AS LateItems, IIf([LateItems]<0,1,"") AS Late, Format([tblSubmissions].[DateComplete],"mmmm") AS [Month], Format([tblSubmissions].[DateComplete],"yyyy") AS [Year], Month([tblSubmissions].[DateComplete]) AS MonthNum, [MonthNum] & "_" & [Month] AS [Order], tblSubmissions.DateComplete
FROM (tblReview INNER JOIN (MERM INNER JOIN (Contacts INNER JOIN [Contacts Extended] ON Contacts.ID = [Contacts Extended].ID) ON MERM.QA_Person = [Contacts Extended].[Contact Name]) ON tblReview.EventID = MERM.EventID) INNER JOIN tblSubmissions ON tblReview.SubmitID = tblSubmissions.SubmitID
WHERE (((Format([tblSubmissions].[DateComplete],"yyyy"))=Year(Date())))
ORDER BY Month([tblSubmissions].[DateComplete]);

