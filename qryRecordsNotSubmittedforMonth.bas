SELECT qryEventDates.EventID, qryEventDates.InstanceID, qryEventDates.EventDate, MERM.[Assigned To], qryEventDates.DocumentName, qryEventDates.Comment, qryEventDates.IsCanned
FROM (qryEventDates LEFT JOIN tblSubmissions ON (qryEventDates.[EventID] = tblSubmissions.[EventID]) AND (qryEventDates.[InstanceID] = tblSubmissions.[InstanceID])) INNER JOIN MERM ON qryEventDates.EventID = MERM.EventID
WHERE (((tblSubmissions.EventID) Is Null) AND ((Month([EventDate])) Between Month(Now()) And Month(Now())+1));

