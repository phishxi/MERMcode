SELECT qryEventDates.EventID, MERM.MERMType, MERM.Group, MERM.[Assigned To], qryEventDates.EventDate, qryEventDates.DocumentName, MERM.Asset, MERM.AQMD_ID, MERM.Site, MERM.Building, MERM.Priority, DateDiff("d",Now(),[EventDate]) AS DaysUntil
FROM MERM INNER JOIN qryEventDates ON MERM.EventID = qryEventDates.EventID
WHERE (((MERM.Group) Like IIf(IsNull([Forms]![MERMHome]![ComboOrg]),"*",[Forms]![MERMHome]![ComboOrg])) AND ((MERM.[Assigned To]) Like IIf(IsNull([Forms]![MERMHome]![ComboName]),"*",[Forms]![MERMHome]![ComboName])) AND ((qryEventDates.EventDate) Between [Forms]![MERMHome]![BeginDate] And [Forms]![MERMHome]![EndDate]) AND ((MERM.Site) Like IIf(IsNull([Forms]![MERMHome]![ComboSite]),"*",[Forms]![MERMHome]![ComboSite])))
ORDER BY qryEventDates.EventDate, DateDiff("d",Now(),[EventDate]);

