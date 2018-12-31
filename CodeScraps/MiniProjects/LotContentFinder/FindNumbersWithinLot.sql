/****** Script for SelectTopNRows command from SSMS  ******/
/*Pulls the full SPTNumber contents with the specified LotId ($(DesiredID))*/
use Guardian

SELECT TOP(100) PERCENT [SPTNumber]
  FROM [Guardian].[Guardian].[ProcessedSPTNumbers] WHERE LotId = $(DesiredID) /*<- Replace with found LotId*/