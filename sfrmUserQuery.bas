SELECT UserQuery.Priority, UserQuery.EventDate, UserQuery.Site, UserQuery.DocumentName, UserQuery.[Assigned To], UserQuery.EventDate, UserQuery.DaysUntil
FROM UserQuery
ORDER BY UserQuery.[DaysUntil];

