use Guardian

declare @LotId int;

select @LotId = max(LotId) from Guardian.Lots where LotId > 0;

select a.SPTNumber, a.SPTFormatId, a.SPTObjectTypeId
from Guardian.SPTDuplicatesCrossLots a
inner join Guardian.Lots b on b.LotId = a.CurrentLotId
inner join Guardian.Lots c on c.LotId = a.LotId
where a.CurrentLotId = @LotId;
go