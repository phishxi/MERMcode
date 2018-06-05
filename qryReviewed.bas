SELECT tblReview.SubmitID, tblReview.EventID, tblReview.InstanceID, MERM.Site, MERM.CommonName, tblReview.DocumentName, MERM.Email, tblReview.EventDate, tblReview.Complete, tblReview.DateComplete, tblSubmissions.DateComplete, MERM.[Assigned To], MERM.QA_Person, [Contacts Extended].Email AS QAemail, Round([tblReview]![DateComplete]-[tblSubmissions]![DateComplete],0) AS TimeToReview
FROM (tblReview INNER JOIN (MERM INNER JOIN (Contacts INNER JOIN [Contacts Extended] ON Contacts.ID = [Contacts Extended].ID) ON MERM.QA_Person = [Contacts Extended].[Contact Name]) ON tblReview.EventID = MERM.EventID) INNER JOIN tblSubmissions ON tblReview.SubmitID = tblSubmissions.SubmitID
ORDER BY MERM.QA_Person;

