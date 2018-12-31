use Guardian

SELECT * FROM [Guardian].[Guardian].[MailQueue] WHERE CHARINDEX('UniSeries has detected a condition where',MailBody)  > 0   