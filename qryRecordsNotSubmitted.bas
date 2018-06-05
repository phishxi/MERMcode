SELECT qryEventDates.EventID, qryEventDates.InstanceID, qryEventDates.EventDate, MERM.[Assigned To], qryEventDates.DocumentName, qryEventDates.Comment, qryEventDates.IsCanned
FROM (qryEventDates LEFT JOIN tblSubmissions ON (qryEventDates.[InstanceID] = tblSubmissions.[InstanceID]) AND (qryEventDates.[EventID] = tblSubmissions.[EventID])) INNER JOIN MERM ON qryEventDates.EventID = MERM.EventID
WHERE (((qryEventDates.EventDate) Between [Forms]![MERMHome]![txtBeginDate] And [Forms]![MERMHome]![txtEndDate]) AND ((tblSubmissions.EventID) Is Null));

