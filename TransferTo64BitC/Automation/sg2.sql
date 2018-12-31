use Guardian

declare @LotId int;

select @LotId = max(LotId) from Guardian.Lots where LotId > 0;

select a.SPTNumber, a.SPTFormatId, a.SPTObjectTypeId
from Guardian.SPTDuplicatesRemoved a
inner join Guardian.Lots b on b.LotId = a.LotId
where a.LotId = @LotId;
go