SELECT MERM.EventID, MERM.Site, MERM.Group, MERM.CommonName, MERM.DocumentName, MERM.Asset, MERM.AQMD_ID, MERM.[Assigned To]
FROM MERM
WHERE (((MERM.Group) Like '*ron*')) OR (((MERM.DocumentName) Like '*ron*')) OR (((MERM.AQMD_ID) Like '*ron*')) OR (((MERM.Asset) Like '*ron*')) OR (((MERM.[Assigned To]) Like '*ron*')) OR (((MERM.CommonName) Like '*ron*'))
ORDER BY MERM.[Assigned To] DESC;

