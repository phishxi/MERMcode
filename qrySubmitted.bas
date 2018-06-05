SELECT tblReview.SubmitID, tblReview.EventID, tblReview.InstanceID, MERM.Site, MERM.CommonName, tblReview.DocumentName, MERM.Email, tblReview.EventDate, tblReview.Complete, tblReview.DateComplete, Format([tblSubmissions].[DateComplete],"yyyy") AS [Year], tblSubmissions.DateComplete, MERM.[Assigned To], MERM.QA_Person, [Contacts Extended].Email AS QAemail, Round([tblSubmissions]![DateComplete]-[tblSubmissions]![EventDate],1) AS DaysToSubmit, ([tblReview].[EventDate]-[tblSubmissions].[DateComplete]) AS LateItems
FROM (tblReview INNER JOIN (MERM INNER JOIN (Contacts INNER JOIN [Contacts Extended] ON Contacts.ID = [Contacts Extended].ID) ON MERM.QA_Person = [Contacts Extended].[Contact Name]) ON tblReview.EventID = MERM.EventID) INNER JOIN tblSubmissions ON tblReview.SubmitID = tblSubmissions.SubmitID
WHERE (((tblSubmissions.DateComplete) Between DateSerial(Year(Date()),1,1) And DateSerial(Year(Date()),12,31)))
ORDER BY MERM.QA_Person;

