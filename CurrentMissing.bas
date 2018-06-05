SELECT MERM.Site, QryAllMERM.EventDate, MERM.[Assigned To], MERM.DocumentName, MERM.Asset, MERM.AQMD_ID
FROM MERM INNER JOIN QryAllMERM ON MERM.EventID = QryAllMERM.EventID
WHERE (((QryAllMERM.EventDate)<=Date()+60))
ORDER BY QryAllMERM.EventDate, MERM.[Assigned To];

