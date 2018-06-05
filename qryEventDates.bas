SELECT qryEventCartesian.EventID, IIf([tblEventException].[EventID] Is Null,IIf(([qryEventCartesian].[PeriodTypeID] Is Null) Or ([qryEventCartesian].[PeriodFreq] Is Null) Or ([qryEventCartesian].[InstanceID] Is Null),[qryEventCartesian].[EventStart],DateAdd([qryEventCartesian].[PeriodTypeID],[qryEventCartesian].[InstanceID]*[qryEventCartesian].[PeriodFreq],[qryEventCartesian].[EventStart])),IIf([tblEventException].[IsCanned],Null,[tblEventException].[InstanceDate])) AS EventDate, qryEventCartesian.DocumentName, qryEventCartesian.InstanceID, tblEventException.IsCanned, qryEventCartesian.Comment, tblEventException.[Complete?], tblEventException.InstanceComment, qryEventCartesian.EventStart, qryEventCartesian.RecurCount, qryEventCartesian.PeriodFreq, ltPeriodType.PeriodType, tblEventException.ActualHours
FROM (qryEventCartesian LEFT JOIN tblEventException ON (qryEventCartesian.InstanceID = tblEventException.InstanceID) AND (qryEventCartesian.EventID = tblEventException.EventID)) LEFT JOIN ltPeriodType ON qryEventCartesian.PeriodTypeID = ltPeriodType.PeriodTypeId
WHERE (((IIf([tblEventException].[EventID] Is Null,IIf(([qryEventCartesian].[PeriodTypeID] Is Null) Or ([qryEventCartesian].[PeriodFreq] Is Null) Or ([qryEventCartesian].[InstanceID] Is Null),[qryEventCartesian].[EventStart],DateAdd([qryEventCartesian].[PeriodTypeID],[qryEventCartesian].[InstanceID]*[qryEventCartesian].[PeriodFreq],[qryEventCartesian].[EventStart])),IIf([tblEventException].[IsCanned],Null,[tblEventException].[InstanceDate])))>Date()-75))
ORDER BY qryEventCartesian.EventID, qryEventCartesian.InstanceID;
