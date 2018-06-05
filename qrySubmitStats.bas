SELECT [Contacts Extended].[Contact Name], tblSubmissions.EventID, tblSubmissions.DateComplete
FROM (MERM INNER JOIN tblSubmissions ON MERM.EventID = tblSubmissions.EventID) INNER JOIN [Contacts Extended] ON tblSubmissions.SubmittedBy = [Contacts Extended].MyID
GROUP BY [Contacts Extended].[Contact Name], tblSubmissions.EventID, tblSubmissions.DateComplete
HAVING (((tblSubmissions.[DateComplete]) Between DateSerial(Year(Date()),Month(Date()),1) And DateSerial(Year(Date()),Month(Date())+1,0)));

