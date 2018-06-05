SELECT MERM.EventID, MERM.Site, MERM.Group, MERM.CommonName, MERM.DocumentName, MERM.Asset, MERM.AQMD_ID, MERM.[Assigned To]
FROM MERM
WHERE (((MERM.Group) Like '**')) OR (((MERM.DocumentName) Like '**')) OR (((MERM.AQMD_ID) Like '**')) OR (((MERM.Asset) Like '**')) OR (((MERM.[Assigned To]) Like '**')) OR (((MERM.CommonName) Like '**'))
ORDER BY MERM.[Assigned To] DESC;

