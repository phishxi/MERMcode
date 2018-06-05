SELECT tblSubmissions.SubmitID, tblSubmissions.EventID, tblSubmissions.InstanceID, MERM.Site, MERM.CommonName, tblSubmissions.DocumentName, MERM.newEmail, Contacts.rcvEmail, tblSubmissions.EventDate, tblSubmissions.Complete, tblSubmissions.DateComplete, MERM.[Assigned To], MERM.QA_Person, [Contacts Extended].Email AS QAemail, Round(Date()-[tblSubmissions]![DateComplete],0) AS numday
FROM ((tblSubmissions LEFT JOIN tblReview ON tblSubmissions.[SubmitID] = tblReview.[SubmitID]) INNER JOIN MERM ON tblSubmissions.EventID = MERM.EventID) INNER JOIN (Contacts INNER JOIN [Contacts Extended] ON Contacts.ID = [Contacts Extended].ID) ON MERM.QA_Person = [Contacts Extended].[Contact Name]
WHERE (((Contacts.rcvEmail)=Yes) AND ((tblReview.SubmitID) Is Null))
ORDER BY MERM.QA_Person, Round(Date()-[tblSubmissions]![DateComplete],0) DESC;

