USE [INTDB_RankingSports]
GO
/****** Object:  StoredProcedure [dbo].[AE_SP001_GetINTDBInformation]    Script Date: 6/22/2015 2:54:47 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--[AE_SP001_GetINTDBInformation]'SBODemoSG'



CREATE Procedure [dbo].[AE_SP001_GetINTDBInformation]
@Entity as varchar(30)
as
begin
Declare @SQL varchar(max)

create table #FINAL 
(
    HTransID [int] NOT NULL,
	HOutlet [nvarchar](50) NOT NULL,
	HPOSTxNo [nvarchar](100) NOT NULL,
	HPOSTillId [nvarchar](100) NULL,
	PHOSTxDate [datetime]   NULL,
	HPOSTxDatetime [datetime]   NULL,
	[HPOSTxType] [nvarchar](5)   NULL,
	HCardCode [nvarchar](50)   NULL,	

	DTransID [int]   NULL,
	DHeaderID [nvarchar](50)   NULL,
	DOutlet [nvarchar](50)   NULL,
	DItemCode [nvarchar](100)   NULL,
	VatGourpSa [nvarchar] (20) NULL,
	DPriceBefDi [numeric](19, 6) NULL,
	DDiscPrcnt [numeric](19, 6)   NULL,
	DPrice [numeric](19, 6)   NULL,
	DQuantity [numeric](19, 6)  NULL,
	DLineTotal [numeric](19, 6) NULL,
	
	PPaymentAmount [numeric](19, 6)   NULL,
	DNetAmount [numeric](19, 6) NULL,
	[Validation2 Msg] [nvarchar](300) NULL
			)

		
set @SQL  = '
SELECT T0.[ID] [HTransID],T0.[Outlet] [HOutlet],T0.[POSTxNo] [HPOSTxNo],T0.[POSTillId] [HPOSTillId],T0.[POSTxDate] [PHOSTxDate],
T0.[POSTxDatetime] [HPOSTxDatetime] ,T0.[POSTxType] [HPOSTxType] ,T0.[CardCode] [CardCode], 
T2.[ID] [DTransID],T2.[HeaderID] [DHeaderID] ,T2.[Outlet] [DOutlet],T2.[ItemCode] [DItemCode], T4.[VatGourpSa], T2.[PriceBefDi] [DPriceBefDi],
T2.[DiscPrcnt] [DDiscPrcnt],T2.[Price] [DPrice],T2.[Quantity] [DQuantity], T2.[LineTotal] [DLineTotal], 
(select sum(TT1.PaymentAmount)  From [AB_Payment] TT1 where TT1.HeaderID = T0.ID) [PPaymentAmount],
(select round(sum(TT.LineTotal + (TT.LineTotal * 0.07)),2)   
from [AB_SalesTransDetail] TT 
where TT.HeaderID  = T0.ID 
 ) [DNetAmount],
case 
   when 
      isnull(T1.[U_POS_RefNo],'''') <> '''' and T0.[POSTxType] = ''S'' then ''Receipt # '' + T1.[U_POS_RefNo] + '' already has an AR Invoice. {''+ cast(T1.DocNum as varchar) +''}'' 
   else '''' end [Validation2 Msg] 

  FROM [AB_SalesTransHeader] T0 
left outer join ' + @Entity + '.. OINV T1 ON T1.[U_POS_RefNo] = T0.[POSTxNo] 
JOIN [AB_SalesTransDetail] T2 ON T2.HeaderID = T0.ID 
LEFT OUTER JOIN ' + @Entity + '.. OITM T4 ON T4.ITEMCODE = T2.ITEMCODE
LEFT OUTER JOIN ' + @Entity + '.. OITT T6 ON T6.Code = T2.ItemCode 
WHERE (isnull([Status], '''') = '''' OR [Status] <> ''SUCCESS'')
ORDER BY T0.ID , T2.ID '

insert into #FINAL 
 execute(@SQL)

SELECT #FINAL.HTransID , COUNT(#FINAL.[Validation2 Msg] )[Validation2]  into #Validation2 FROM #FINAL WHERE ISNULL(#FINAL.[Validation2 Msg],'') <> ''
GROUP BY #FINAL.HTransID

SELECT #Final.* ,
CASE WHEN ISNULL(V2.Validation2 ,'') = '' THEN 0 ELSE V2.Validation2 END [Validation2Count],
ltrim(#final.[Validation2 Msg] ) [DetailsErrMsg]
 FROM #FINAL 
LEFT OUTER JOIN #Validation2 V2 ON V2.HTransID = #FINAL.HTransID
order by cast(#Final.DHeaderID as integer) , cast(#Final.DTransID as integer)

drop table #FINAL
drop table #Validation2
End








 
 
 
 

GO
/****** Object:  StoredProcedure [dbo].[AE_SP002_GetNoStockItem]    Script Date: 6/22/2015 2:54:47 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--EXEC [dbo].[AE_SP002_GetNoStockItem] 'STUTTGART_LIVE'
CREATE PROCEDURE [dbo].[AE_SP002_GetNoStockItem]
@Entity AS NVARCHAR(30)
as
BEGIN
DECLARE @SQL NVARCHAR(max)


DELETE FROM AB_NoStockItem


SET @SQL = '
SELECT T0.*,ISNULL(T1.OnHand,0) AS OnHand
FROM
(
	SELECT T0.ID,T1.ID AS LineID,T0.Outlet,T1.ItemCode,T2.SellItem,T2.InvntItem,T1.Quantity,T1.Quantity * ISNULL(T2.NumInSale,1) AS ReqQty,
	CASE WHEN T2.InvntItem=''Y'' THEN T1.ItemCode ELSE T3.Code END AS ChildItemCode,
	CASE WHEN T2.InvntItem=''Y'' THEN T1.Quantity * ISNULL(T2.NumInSale,1) 
	ELSE ISNULL(((T1.Quantity*ISNULL(T2.NumInSale,1))*T3.Quantity)/NULLIF(T3.FatherQty,0),0) END AS BOMQty
	FROM AB_SalesTransHeader T0
	LEFT JOIN AB_SalesTransDetail T1 ON T0.ID = T1.HeaderID
	LEFT JOIN ' + @Entity + '..OITM T2 ON T1.ItemCode = T2.ItemCode
	LEFT JOIN ' + @Entity + '..[SV_AB_BOMTREE] T3 ON T1.ItemCode = T3.Father
	WHERE T0.[Status] = ''FAIL'' 
) T0
LEFT JOIN ' + @Entity + '..OITW T1 ON T0.Outlet = T1.WhsCode AND T0.ChildItemCode = T1.ItemCode
WHERE T0.BOMQty > T1.OnHand'


INSERT INTO AB_NoStockItem
EXECUTE (@SQL)



UPDATE AB_SalesTransDetail  
SET ErrMsg =  ISNULL(ErrMsg,'') + 'No Stock: ' + T1.NoStockItem 
FROM AB_SalesTransDetail T0
LEFT JOIN
(
	SELECT A.LineID,D.NoStockItem
	FROM AB_NoStockItem A
	CROSS APPLY 
	( 
		SELECT STUFF
		((
			SELECT + ', ' + CAST(B.ChildItemCode AS NVARCHAR(MAX)) 
			FROM AB_NoStockItem B
			WHERE A.LineID = B.LineID 
			ORDER BY A.LineID,b.ChildItemCode
			FOR XML PATH(''),TYPE).value('.','NVARCHAR(MAX)'),1,2,'')
	) D (NoStockItem)
	GROUP BY A.LineID,D.NoStockItem
) T1 ON T0.ID = T1.LineID
WHERE T1.NoStockItem IS NOT NULL





END








 
 
 
 


GO
/****** Object:  StoredProcedure [dbo].[AE_SP003_ItemMasterSync]    Script Date: 6/22/2015 2:54:47 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--[dbo].[AE_SP003_ItemMasterSync]'SBODemoSG',2

CREATE procedure [dbo].[AE_SP003_ItemMasterSync]
@SAPDB as varchar(50),
@Var as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)

SET @SQL = '
INSERT INTO [AB_ItemMaster]  ([ItemCode],[ItemName],[Brand],[Model],[Category],[Department],[Vendor],[Barcode],[Active],[UOM],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.[ItemCode], T0.[ItemName], T1.[FirmName], T0.[SWW], T2.[ItmsGrpNam],T0.[U_AB_SubCategory],T3.[CardName], T0.[CodeBars], T0.[validFor], T0.[IUoMEntry],
DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE() 
FROM '+ @SAPDB +' ..OITM T0  
LEFT OUTER JOIN '+ @SAPDB +' ..OMRC T1 ON T0.[FirmCode] = T1.[FirmCode] 
LEFT OUTER JOIN '+ @SAPDB +' ..OITB T2 ON T0.[ItmsGrpCod] = T2.[ItmsGrpCod] 
LEFT OUTER JOIN '+ @SAPDB +' ..OCRD T3 ON T0.[CardCode] = T3.[CardCode]
WHERE T0.[ItemCode] NOT IN (SELECT [AB_ItemMaster].ItemCode  FROM [AB_ItemMaster] ) ORDER BY  T0.[ItemCode]'

SET @SQL1 ='
UPDATE
     AB_ItemMaster 
SET
     AB_ItemMaster.ItemName = OITM.ItemName,
     AB_ItemMaster.Brand = OMRC.FirmName,
	  AB_ItemMaster.Model = OITM.SWW,
	   AB_ItemMaster.Category = OITB.ItmsGrpNam,
	    AB_ItemMaster.Department = OITM.U_AB_SubCategory,
		 AB_ItemMaster.Vendor = OCRD.CardName,
		  AB_ItemMaster.Barcode = OITM.CodeBars,
		   AB_ItemMaster.Active = OITM.validFor,
		    AB_ItemMaster.UOM = OITM.IUoMEntry,
			AB_ItemMaster.SAPSyncDate = GETDATE(),
			AB_ItemMaster.SAPSyncDateTime = GETDATE()
FROM  AB_ItemMaster 
LEFT OUTER JOIN '+ @SAPDB +' ..OITM ON AB_ItemMaster.ItemCode = OITM.ItemCode 
LEFT OUTER JOIN '+ @SAPDB +' ..OMRC ON OITM.[FirmCode] = OMRC.[FirmCode] 
LEFT OUTER JOIN '+ @SAPDB +' ..OITB ON OITM.[ItmsGrpCod] = OITB.[ItmsGrpCod] 
LEFT OUTER JOIN '+ @SAPDB +' ..OCRD ON OITM.[CardCode] = OCRD.[CardCode]

WHERE
    '+ @SAPDB +' ..OITM.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')'


--PRINT @SQL
--PRINT @SQL1

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP004_PriceListSync]    Script Date: 6/22/2015 2:54:47 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE procedure [dbo].[AE_SP004_PriceListSync]
@SAPDB as varchar(50),
@Var as varchar(10),
@Vat as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)

SET @SQL = '
INSERT INTO [AB_PriceList]  ([ItemCode],[PriceListName],[Currency],[Price],[PriceGST],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.[ItemCode], T2.[ListName], T1.[Currency], T1.[Price], T1.[Price] + (T1.[Price] * ( ' + @Vat + ' / 100)),DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE()
FROM '+ @SAPDB +' ..OITM T0  
LEFT OUTER JOIN '+ @SAPDB +' ..ITM1 T1 ON T0.[ItemCode] = T1.[ItemCode]
LEFT OUTER JOIN '+ @SAPDB +' ..OPLN T2 ON T1.[PriceList] = T2.[ListNum] 
WHERE T0.[ItemCode] NOT IN (SELECT ItemCode  FROM [AB_PriceList] ) ORDER BY  T0.[ItemCode]'

SET @SQL1 ='
UPDATE
     AB_PriceList 
SET
     AB_PriceList.PriceListName = T3.ListName,
	  AB_PriceList.Currency = T2.Currency,
	   AB_PriceList.Price = T2.Price,
	    AB_PriceList.PriceGST =T2.[Price] + ROUND(T1.[Price] * ( ' + @Vat + ' / 100),2),
		  AB_PriceList.SAPSyncDate = DATEADD(day,datediff(day,0,GETDATE()),0),
			AB_PriceList.SAPSyncDateTime = GETDATE()
FROM  AB_PriceList T0
LEFT OUTER JOIN '+ @SAPDB +' ..OITM T1 ON T0.ItemCode = T1.ItemCode 
LEFT OUTER JOIN '+ @SAPDB +' ..ITM1 T2 ON T1.[ItemCode] = T1.[ItemCode] 
LEFT OUTER JOIN '+ @SAPDB +' ..OPLN T3 ON T2.[PriceList] = T3.[ListNum] 

WHERE
   T1.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')'


--PRINT @SQL
--PRINT @SQL1

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP005_PromotionPriceListSync]    Script Date: 6/22/2015 2:54:47 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- [dbo].[AE_SP005_PromotionPriceListSync]'SBODemoSG',2,7
CREATE procedure [dbo].[AE_SP005_PromotionPriceListSync]
@SAPDB as varchar(50),
@Var as varchar(10),
@Vat as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)


SET @SQL = 'DELETE FROM AB_Promotion WHERE CreateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')
		OR UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')'

SET @SQL1 = '
INSERT INTO [AB_Promotion]  ([ItemCode],[PriceListName],[Currency],[Price],[PriceGST],[FromDate],[ToDate],[CreateDate],[UpdateDate],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.ItemCode,T2.ListName,T1.Currency,T1.Price,ROUND(T1.Price * (' + @Vat + ' / 100),2),T1.FromDate,T1.ToDate,T0.CreateDate,T0.UpdateDate,
DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE()
FROM '+ @SAPDB +' ..OSPP T0
LEFT JOIN '+ @SAPDB +' ..SPP1 T1 ON T0.CardCode=T1.CardCode AND T0.ItemCode=T1.ItemCode AND T0.ListNum=T1.ListNum
LEFT JOIN '+ @SAPDB +' ..OPLN T2 ON T0.ListNum=T2.ListNum
WHERE T0.CreateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')
OR T0.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ') '

--PRINT @SQL
--PRINT @SQL1

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP006_WareHouseSync]    Script Date: 6/22/2015 2:54:47 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--[dbo].[AE_SP006_WareHouseSync]'SBODemoSG',2
CREATE procedure [dbo].[AE_SP006_WareHouseSync]
@SAPDB as varchar(50),
@Var as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)

SET @SQL = '
INSERT INTO [AB_Warehouses]  ( [WhsCode],[WhsName],[Active],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.WhsCode , T0.WhsName , T0.Inactive,DATEADD(day,datediff(day,0,GETDATE()),0), GETDATE() FROM '+ @SAPDB +' ..OWHS T0
WHERE T0.WhsCode NOT IN (SELECT WhsCode FROM AB_Warehouses)'

SET @SQL1 ='
UPDATE
     AB_Warehouses 
SET
     AB_Warehouses.WhsName = T1.WhsName,
	  AB_Warehouses.Active = T1.Inactive,
	   	  AB_Warehouses.SAPSyncDate = DATEADD(day,datediff(day,0,GETDATE()),0),
			AB_Warehouses.SAPSyncDateTime = GETDATE()
FROM  AB_Warehouses T0
LEFT OUTER JOIN '+ @SAPDB +' ..OWHS T1 ON T0.WhsCode = T1.WhsCode 
WHERE
   T1.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')'


--PRINT @SQL
--PRINT @SQL1

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP007_CustomerSync]    Script Date: 6/22/2015 2:54:47 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




------[dbo].[AE_SP007_CustomerSync]'SBODemoSG',2

CREATE procedure [dbo].[AE_SP007_CustomerSync]
@SAPDB as varchar(50),
@Var as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)

SET @SQL = '
INSERT INTO [AB_Customers]  ([CardCode],[CardName],[GroupName],[PriceListName],[Phone1],[POSSearch],[Active],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.[CardCode], T0.[CardName], T2.[GroupName], T1.[ListName], T0.[Phone1], T0.[AddID], T0.[validFor],DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE()
FROM '+ @SAPDB +' ..OCRD T0  
LEFT OUTER JOIN '+ @SAPDB +' ..OPLN T1 ON T0.[ListNum] = T1.[ListNum]
LEFT OUTER JOIN '+ @SAPDB +' ..OCRG T2 ON T0.[GroupCode] = T2.[GroupCode] 
WHERE T0.[CardCode] NOT IN (SELECT CardCode  FROM [AB_Customers] ) ORDER BY  T0.[CardCode]'

SET @SQL1 ='
UPDATE
     AB_Customers 
SET
     AB_Customers.CardName = T0.[CardName],
	  AB_Customers.GroupName = T2.[GroupName],
	   AB_Customers.PriceListName = T1.[ListName],
	    AB_Customers.Phone1 = T0.[Phone1],
		AB_Customers.POSSearch = T0.[AddID],
		AB_Customers.Active = T0.[validFor],
		  AB_Customers.SAPSyncDate = DATEADD(day,datediff(day,0,GETDATE()),0),
			AB_Customers.SAPSyncDateTime = GETDATE()
FROM  AB_Customers TT
LEFT OUTER JOIN '+ @SAPDB +' ..OCRD T0 ON T0.CardCode = TT.CardCode 
LEFT OUTER JOIN '+ @SAPDB +' ..OPLN T1 ON T0.[ListNum] = T1.[ListNum]
LEFT OUTER JOIN '+ @SAPDB +' ..OCRG T2 ON T0.[GroupCode] = T2.[GroupCode] 

WHERE
   T0.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')'


--PRINT @SQL
--PRINT @SQL1

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
