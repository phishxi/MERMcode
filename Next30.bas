SELECT MERM.Priority, MERM.Building, QryAllMERM.EventDate, MERM.Site, MERM.MERMType, MERM.DocumentName, MERM.[Assigned To], MERM.Asset, MERM.AQMD_ID
FROM MERM INNER JOIN QryAllMERM ON MERM.EventID = QryAllMERM.EventID
WHERE (((QryAllMERM.EventDate)<DateAdd("m",+1,Date())))
ORDER BY QryAllMERM.EventDate, MERM.DocumentName;

