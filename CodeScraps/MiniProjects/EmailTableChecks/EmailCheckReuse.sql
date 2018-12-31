use Guardian

SELECT * FROM [Guardian].[Guardian].[MailQueue] WHERE CHARINDEX('Warning:  UniSeries has detected an attempt to',MailBody)  > 0   