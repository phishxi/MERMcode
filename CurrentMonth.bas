SELECT QryAllMERM.EventID, QryAllMERM.EventDate, MERM.Site, MERM.MERMType, QryAllMERM.[Assigned To], MERM.DocumentName, MERM.Asset, MERM.AQMD_ID
FROM MERM INNER JOIN QryAllMERM ON MERM.EventID = QryAllMERM.EventID
WHERE (((QryAllMERM.EventDate)<DateAdd("m",2,Date())))
ORDER BY QryAllMERM.EventDate, QryAllMERM.[Assigned To];

