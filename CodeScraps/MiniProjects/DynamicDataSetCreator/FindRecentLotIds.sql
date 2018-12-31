/****** Script for SelectTopNRows command from SSMS  ******/
/* Find the desired LotId */
/*with Guardian*/

SELECT TOP 1000 [LotId]
      ,[LineId]
      ,[ProductId]
      ,[LotStateId]
/*      ,[LineLotId]*/
/*      ,[FingerPrint]*/
      ,[StartDate]
/*      ,[ExpDate]*/
      ,[CreateDate]
      ,[LotCode]
/*      ,[EndLot]*/
/*      ,[G2GId]*/
      ,[TimeDiff]
/*      ,[IsAggregationOn]*/
      ,[TotalComItems]
      ,[TotalDecomItems]
      ,[TotalLotItemsInLot]
      ,[TotalLotItems]
      ,[TotalSPTinLot]
      ,[TotalLotSPT]
/*      ,[SingleLevelHierarchy]*/
      ,[NotificationStateId]
/*      ,[AdvisorProductId]*/
      ,[CrossLotSPTCheck]
      ,[IsQuarantined]
      ,[IsOverrideQuarantine]
/*      ,[UserName]*/
/*      ,[SecondUserName]*/
  FROM [Guardian].[Guardian].[Lots]