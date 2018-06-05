SELECT MERM.Priority, MERM.DocumentName, MERM.Site, MERM.[Assigned To], qryEventDates.EventDate, IIf([DelayedSubmission] Is Not Null,([EventDate]-Date())+[DelayedSubmission],[EventDate]-Date()) AS numday, MERM.DelayedSubmission, MERM.EventID
FROM MERM INNER JOIN qryEventDates ON MERM.EventID = qryEventDates.EventID
WHERE (((IIf([DelayedSubmission] Is Not Null,([EventDate]-Date())+[DelayedSubmission],[EventDate]-Date()))<=0))
ORDER BY MERM.Priority, MERM.[Assigned To], IIf([DelayedSubmission] Is Not Null,([EventDate]-Date())+[DelayedSubmission],[EventDate]-Date()) DESC;

