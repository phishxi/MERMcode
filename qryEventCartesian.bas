SELECT MERM.EventID, tblCount.CountID AS InstanceID, MERM.DocumentName, MERM.EventStart, MERM.RecurCount, MERM.PeriodFreq, MERM.PeriodTypeID, MERM.Comment
FROM tblCount, MERM
WHERE (((MERM.RecurCount) Is Null)) OR (((tblCount.CountID)<=[MERM].[RecurCount]))
ORDER BY MERM.EventID;

