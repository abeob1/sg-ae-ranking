

ALTER Procedure [dbo].[AE_SP001_GetINTDBInformation]
@Entity as varchar(30)
as
begin
Declare @SQL varchar(max)

create table #FINAL 
(
    HTransID [int] NOT NULL,
	HOutlet [nvarchar](50) NOT NULL,
	HPOSTxNo [nvarchar](100) NOT NULL,
	HSalesman [Int] NOT NULL,
	HCommEntitle [nvarchar](100)  NULL,
	HCommPercent [nvarchar](100)  NULL,
	HPOSTillId [nvarchar](100) NULL,
	PHOSTxDate [datetime]   NULL,
	HPOSTxDatetime [datetime]   NULL,
	HPOSTxType [nvarchar](5)   NULL,
	HDocTotal [numeric](19, 6) NULL,
	HCardCode [nvarchar](50)   NULL,	
	DTransID [int]   NULL,
	DHeaderID [nvarchar](50)   NULL,
	DOutlet [nvarchar](50)   NULL,
	DItemCode [nvarchar](100)   NULL,
	VatGourpSa [nvarchar] (20) NULL,
	DPriceBefDi [numeric](19, 6) NULL,
	DDiscPrcnt [numeric](19, 6)   NULL,
	DTotalGST [numeric](19, 6)   NULL,
	DQuantity [numeric](19, 6)  NULL,
	DLineTotal [numeric](19, 6) NULL,
	
	PPaymentAmount [numeric](19, 6)   NULL,
	DNetAmount [numeric](19, 6) NULL,
	[Validation2 Msg] [nvarchar](max) NULL,
	[Validation3 Msg] [nvarchar](max) NULL
			)

		
set @SQL  = '
SELECT T0.[ID] [HTransID],T0.[Outlet] [HOutlet],T0.[POSTxNo] [HPOSTxNo],T0.[Salesman] [HSalesman] , T7.[U_AB_CommEntitle] [HCommEntitle],
T7.[U_AB_CommPercent] [HCommPercent] ,T0.[POSTillId] [HPOSTillId],T0.[POSTxDate] [PHOSTxDate],
T0.[POSTxDatetime] [HPOSTxDatetime] ,T0.[POSTxType] [HPOSTxType] , isnull(T0.[DocTotal],0) [HDocTotal], T0.[CardCode] [CardCode], 
T2.[ID] [DTransID],T2.[POSTxNo] [DHeaderID] ,T2.[Outlet] [DOutlet],T2.[ItemCode] [DItemCode], T4.[VatGourpSa], T2.[UnitPrice] [DPriceBefDi],
T2.[DiscAmount] [DDiscPrcnt],T2.[TotalGST] [DTotalGST],T2.[Quantity] [DQuantity], T2.[LineTotal] [DLineTotal],
(select sum(TT1.PaymentAmount)  From [AB_Payment] TT1 where TT1.POSTxNo = T0.POSTxNo) [PPaymentAmount],
(select round(sum(TT.LineTotal + (TT.LineTotal * 0.07)),2)   
from [AB_SalesTransDetail] TT 
where TT.POSTxNo  = T0.POSTxNo 
 ) [DNetAmount],
case 
   when 
      isnull(T1.[U_AB_POSTxNo],'''') <> '''' and T0.[POSTxType] = ''S'' then ''Receipt # '' + T1.[U_AB_POSTxNo] + '' already has an AR Invoice. {''+ cast(T1.DocNum as varchar) +''}'' 
   else '''' end [Validation2 Msg],
case 
  when 
     isnull(T0.[DocTotal],0) <> (select isnull(sum(TT1.PaymentAmount),0)  From [AB_Payment] TT1 where TT1.POSTxNo = T0.POSTxNo) then ''AR Invoice Total not equal to Payment Total.'' 
 else '''' end [Validation3 Msg] 

  FROM [AB_SalesTransHeader] T0 
left outer join ' + @Entity + '.. OINV T1 ON T1.[U_AB_POSTxNo] = T0.[POSTxNo] 
JOIN [AB_SalesTransDetail] T2 ON T2.POSTxNo = T0.POSTxNo 
LEFT OUTER JOIN ' + @Entity + '.. OITM T4 ON T4.ITEMCODE = T2.ITEMCODE
LEFT OUTER JOIN ' + @Entity + '.. OITT T6 ON T6.Code = T2.ItemCode 
LEFT OUTER JOIN ' + @Entity + '.. OSLP T7 ON T7.SlpCode = T0.Salesman 
WHERE (isnull([Status], '''') = '''' OR [Status] <> ''SUCCESS'')
ORDER BY T0.ID , T2.ID '

insert into #FINAL 
 execute(@SQL)

SELECT #FINAL.HTransID , COUNT(#FINAL.[Validation2 Msg] )[Validation2]  into #Validation2 FROM #FINAL WHERE ISNULL(#FINAL.[Validation2 Msg],'') <> ''
GROUP BY #FINAL.HTransID

SELECT #FINAL.HTransID , COUNT(#FINAL.[Validation3 Msg] )[Validation3]  into #Validation3 FROM #FINAL WHERE ISNULL(#FINAL.[Validation3 Msg],'') <> ''
GROUP BY #FINAL.HTransID

SELECT #Final.*, 
CASE WHEN ISNULL(V2.Validation2 ,'') = '' THEN 0 ELSE V2.Validation2 END [Validation2Count],
CASE WHEN ISNULL(V3.Validation3 ,'') = '' THEN 0 ELSE V3.Validation3 END [Validation3Count],
ltrim(#final.[Validation2 Msg] ) + ' ' + ltrim(#final.[Validation3 Msg]) [DetailsErrMsg]


 FROM #FINAL 
LEFT OUTER JOIN #Validation2 V2 ON V2.HTransID = #FINAL.HTransID
LEFT OUTER JOIN #Validation3 V3 ON V3.HTransID = #FINAL.HTransID
order by cast(#Final.HTransID as integer) , cast(#Final.DTransID as integer)

drop table #FINAL
drop table #Validation2
End








 
 
 
 

GO
/****** Object:  StoredProcedure [dbo].[AE_SP001_GetINTDBInformation_OLD]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--[AE_SP001_GetINTDBInformation]'SBODemoSG'


--EXEC [dbo].[AE_SP002_GetNoStockItem] 'STUTTGART_LIVE'
ALTER PROCEDURE [dbo].[AE_SP002_GetNoStockItem]
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
/****** Object:  StoredProcedure [dbo].[AE_SP003_ItemMasterSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--[dbo].[AE_SP003_ItemMasterSync]'SBODemoSG',2

ALTER procedure [dbo].[AE_SP003_ItemMasterSync]
@SAPDB as varchar(50),
@Var as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)
DECLARE @SQL2 VARCHAR(MAX)
DECLARE @SQL3 VARCHAR(MAX)

create table #T1 (Itemcode nvarchar(100) , ItemName nvarchar(200),indicatorname nvarchar(500), indicatorvalue nvarchar(500),ItmsGrpNam nvarchar (500)  )
create table #T2 (Itemcode nvarchar(100) COLLATE SQL_Latin1_General_CP850_CI_AS, Properties nvarchar (max)  )

set @SQL2 = 'Select Itemcode, ItemName,
  indicatorname,
  indicatorvalue,
  ItmsGrpNam 
from '+ @SAPDB +' ..OITM T0
cross apply
(
  values
  (''1'', QryGroup1),  (''2'', QryGroup2),  (''3'', QryGroup3),  (''4'', QryGroup4),  (''5'', QryGroup5),  (''6'', QryGroup6),  (''7'', QryGroup7),  (''8'', QryGroup8),
  (''9'', QryGroup9),  (''10'', QryGroup10),  (''11'', QryGroup11),  (''12'', QryGroup12),  (''13'', QryGroup13),  (''14'', QryGroup14),  (''15'', QryGroup15),  (''16'', QryGroup16),
  (''17'', QryGroup17),  (''18'', QryGroup18),  (''19'', QryGroup19),  (''20'', QryGroup20),  (''21'', QryGroup21),  (''22'', QryGroup22),  (''23'', QryGroup23),  (''24'', QryGroup24),
  (''25'', QryGroup25),  (''26'', QryGroup26),  (''27'', QryGroup27),  (''28'', QryGroup28),  (''29'', QryGroup29),  (''30'', QryGroup30),  (''31'', QryGroup31),  (''32'', QryGroup32),
   (''33'', QryGroup33), (''34'', QryGroup34),  (''35'', QryGroup35),  (''36'', QryGroup36),  (''37'', QryGroup37), (''38'', QryGroup38),  (''39'', QryGroup39), (''40'', QryGroup40),  (''41'', QryGroup41),
  (''42'', QryGroup42),  (''43'', QryGroup43),  (''44'', QryGroup44),  (''45'', QryGroup45),  (''46'', QryGroup46),  (''47'', QryGroup47),  (''48'', QryGroup48),  (''49'', QryGroup49),
  (''50'', QryGroup50),  (''51'', QryGroup51),  (''52'', QryGroup52),  (''53'', QryGroup53),  (''54'', QryGroup54),  (''55'', QryGroup55),  (''56'', QryGroup56),  (''57'', QryGroup57),
  (''58'', QryGroup58),  (''59'', QryGroup59),  (''60'', QryGroup60),  (''61'', QryGroup61),  (''62'', QryGroup62),  (''63'', QryGroup63),  (''64'', QryGroup64)
  ) c (indicatorname, indicatorvalue)
  join '+ @SAPDB +' ..OITG T11 on indicatorname = T11.ItmsTypCod
  where indicatorvalue = ''Y'''

 -- print(@sql2)

  set @SQL3 = 'Select distinct ST2.ItemCode , 
    substring(
        (
            Select '','' +ST1.indicatorname  AS [text()]
            From #T1 ST1
            Where ST1.ItemCode  = ST2.ItemCode 
            ORDER BY ST1.ItemCode 
            For XML PATH ('''')
        ), 2, 1000) [Properties]
From #T1 ST2'

insert into #T1 execute (@SQL2 )
insert into #T2 execute (@SQL3 )

SET @SQL = '
INSERT INTO [AB_ItemMaster]  ([ItemCode],[ItemName],[Brand],[Model],[Category],[Department],[Vendor],[Barcode],[Active],[UOM],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.[ItemCode], left(T0.[ItemName],50), T1.[FirmCode], T0.[SWW], T4.[Properties] ,T2.[ItmsGrpCod] ,T3.[CardCode], T0.[CodeBars], T0.[validFor], T0.[InvntryUom],
DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE() 
FROM '+ @SAPDB +' ..OITM T0 
LEFT OUTER JOIN #T2 T4  ON T4.ItemCode = T0.ItemCode
LEFT OUTER JOIN '+ @SAPDB +' ..OMRC T1 ON T0.[FirmCode] = T1.[FirmCode] 
LEFT OUTER JOIN '+ @SAPDB +' ..OITB T2 ON T0.[ItmsGrpCod] = T2.[ItmsGrpCod] 
LEFT OUTER JOIN '+ @SAPDB +' ..OCRD T3 ON T0.[CardCode] = T3.[CardCode]

WHERE T0.[ItemCode] NOT IN (SELECT [AB_ItemMaster].ItemCode  FROM [AB_ItemMaster] ) ORDER BY  T0.[ItemCode]'

SET @SQL1 ='
UPDATE
     AB_ItemMaster 
SET
     AB_ItemMaster.ItemName = left(OITM.ItemName,50),
     AB_ItemMaster.Brand = OMRC.FirmCode,
	  AB_ItemMaster.Model = OITM.SWW,
	   AB_ItemMaster.Category = #T2.[Properties],
	    AB_ItemMaster.Department = OITB.ItmsGrpCod,
		 AB_ItemMaster.Vendor = OCRD.CardCode,
		  AB_ItemMaster.Barcode = OITM.CodeBars,
		   AB_ItemMaster.Active = OITM.validFor,
		    AB_ItemMaster.UOM = OITM.InvntryUom,
			AB_ItemMaster.SAPSyncDate = DATEADD(day,datediff(day,0,GETDATE()),0),
			AB_ItemMaster.SAPSyncDateTime = GETDATE()
FROM  AB_ItemMaster 
LEFT OUTER JOIN '+ @SAPDB +' ..OITM ON AB_ItemMaster.ItemCode = OITM.ItemCode 
LEFT OUTER JOIN '+ @SAPDB +' ..OMRC ON OITM.[FirmCode] = OMRC.[FirmCode] 
LEFT OUTER JOIN '+ @SAPDB +' ..OITB ON OITM.[ItmsGrpCod] = OITB.[ItmsGrpCod] 
LEFT OUTER JOIN '+ @SAPDB +' ..OCRD ON OITM.[CardCode] = OCRD.[CardCode]
LEFT OUTER JOIN #T2  ON #T2.ItemCode = OITM.ItemCode

WHERE
    '+ @SAPDB +' ..OITM.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')'


--PRINT @SQL
--PRINT @SQL1

EXEC(@SQL)
EXEC(@SQL1)
Drop table #T1
Drop table #T2	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP004_PriceListSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--[AE_SP004_PriceListSync]'SBODemoSG','2','7'

ALTER procedure [dbo].[AE_SP004_PriceListSync]
@SAPDB as varchar(50),
@Var as varchar(10),
@Vat as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)
set @Vat = 1 + (cast(@Vat as decimal(19,2)) /100)

SET @SQL = '
INSERT INTO [AB_PriceList]  ([ItemCode],[PriceList],[PriceListName],[Currency],[Price],[PriceGST],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.[ItemCode],T1.[PriceList] ,T2.[ListName], T1.[Currency], 
CASE WHEN T2.IsGrossPrc = ''N'' THEN T1.Price ELSE T1.Price / ' + @Vat + ' END, 
CASE WHEN T2.IsGrossPrc = ''Y'' THEN T1.Price ELSE T1.Price * ' + @Vat + ' END,
DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE()
FROM '+ @SAPDB +' ..OITM T0  
LEFT OUTER JOIN '+ @SAPDB +' ..ITM1 T1 ON T0.[ItemCode] = T1.[ItemCode]
LEFT OUTER JOIN '+ @SAPDB +' ..OPLN T2 ON T1.[PriceList] = T2.[ListNum]
WHERE T0.[ItemCode] NOT IN (SELECT ItemCode  FROM [AB_PriceList] ) AND T2.GroupCode <> 1 ORDER BY  T0.[ItemCode]'

SET @SQL1 ='
UPDATE
     AB_PriceList 
SET
    AB_PriceList.PriceList = T2.PriceList,
	AB_PriceList.PriceListName = T3.ListName,
	  AB_PriceList.Currency = T2.Currency,
	   AB_PriceList.Price = CASE WHEN T3.IsGrossPrc = ''N'' THEN T2.Price ELSE T2.Price / ' + @Vat  + ' END, 
	    AB_PriceList.PriceGST =CASE WHEN T3.IsGrossPrc = ''Y'' THEN T2.Price ELSE T2.Price * ' + @Vat  + ' END,
		  AB_PriceList.SAPSyncDate = DATEADD(day,datediff(day,0,GETDATE()),0),
			AB_PriceList.SAPSyncDateTime = GETDATE()
FROM  AB_PriceList T0
LEFT OUTER JOIN '+ @SAPDB +' ..OITM T1 ON T0.ItemCode = T1.ItemCode 
LEFT OUTER JOIN '+ @SAPDB +' ..ITM1 T2 ON T0.[ItemCode] = T2.[ItemCode] and T0.[PriceList] = T2.[PriceList]
LEFT OUTER JOIN '+ @SAPDB +' ..OPLN T3 ON T2.[PriceList] = T3.[ListNum] 
WHERE
   T1.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ') AND T3.GroupCode <> 1' 


--PRINT @SQL
--PRINT @SQL1
--T1.[Price] + (T1.[Price] * ( ' + @Vat + ' / 100)),  T2.[Price] + ROUND(T1.[Price] * ( ' + @Vat + ' / 100),2),
--LEFT OUTER JOIN '+ @SAPDB +' ..OCRG T3 ON T2.ListNum = T3.PriceList 
EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP005_PromotionPriceListSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- [dbo].[AE_SP005_PromotionPriceListSync]'SBODemoSG',2,7
ALTER procedure [dbo].[AE_SP005_PromotionPriceListSync]
@SAPDB as varchar(50),
@Var as varchar(10),
@Vat as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)
set @Vat = 1 + (cast(@Vat as decimal(19,2)) /100)

SET @SQL = 'DELETE FROM AB_Promotion WHERE CreateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')
		OR UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')'

SET @SQL1 = '
INSERT INTO [AB_Promotion]  ([ItemCode],[PriceList],[PriceListName],[Currency],[Price],[PriceGST],[FromDate],[ToDate],
[CreateDate],[UpdateDate],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.ItemCode,T0.ListNum ,T2.ListName,T1.Currency,
CASE WHEN T2.IsGrossPrc = ''N'' THEN T1.Price ELSE T1.Price / ' + @Vat + ' END AS Price,
CASE WHEN T2.IsGrossPrc = ''Y'' THEN T1.Price ELSE T1.Price * ' + @Vat + ' END AS PriceGST,
T1.FromDate,T1.ToDate,T0.CreateDate,T0.UpdateDate,
DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE()
FROM '+ @SAPDB +' ..OSPP T0
LEFT JOIN '+ @SAPDB +' ..SPP1 T1 ON T0.CardCode=T1.CardCode AND T0.ItemCode=T1.ItemCode AND T0.ListNum=T1.ListNum
LEFT JOIN '+ @SAPDB +' ..OPLN T2 ON T0.ListNum=T2.ListNum AND T2.GroupCode <> 1
WHERE T0.CreateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')
OR T0.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ') '

--PRINT @SQL
--PRINT @SQL1
--ROUND(T1.Price * (' + @Vat + ' / 100),2)
EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP006_WareHouseSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--[dbo].[AE_SP006_WareHouseSync]'SBODemoSG',2
ALTER procedure [dbo].[AE_SP006_WareHouseSync]
@SAPDB as varchar(50),
@Var as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)

SET @SQL = '
INSERT INTO [AB_Warehouses]  ( [WhsCode],[WhsName],[Active],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.WhsCode , T0.WhsName , case when  T0.Inactive = ''N'' then ''Y'' else ''N'' end,DATEADD(day,datediff(day,0,GETDATE()),0), GETDATE() FROM '+ @SAPDB +' ..OWHS T0
WHERE T0.WhsCode NOT IN (SELECT WhsCode FROM AB_Warehouses)'

SET @SQL1 ='
UPDATE
     AB_Warehouses 
SET
     AB_Warehouses.WhsName = T1.WhsName,
	  AB_Warehouses.Active = case when  T1.Inactive = ''N'' then ''Y'' else ''N'' end,
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
/****** Object:  StoredProcedure [dbo].[AE_SP007_CustomerSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

------[dbo].[AE_SP007_CustomerSync]'SBODemoSG',2

ALTER procedure [dbo].[AE_SP007_CustomerSync]
@SAPDB as varchar(50),
@Var as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)

SET @SQL = '
INSERT INTO [AB_Customers]  ([CardCode],[CardName],[PriceList],[PriceListName],[Phone1],[Mobile],[Email],[Address1],[Address2],[Address3],
 [Country],[Zipcode],[DOB],[JoinDate],[ExpiryDate],[POSSearch],[Active],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.[CardCode], T0.[CardName], T1.[ListNum], T1.[ListName], T0.[Phone1],T0.[Cellular],T0.[E_Mail],T0.[Address],T0.[StreetNo],T0.[Block],T3.[Name] ,T0.[ZipCode], 
T0.[U_AB_DOB], T0.[U_AB_JoinDate],T0.[U_AB_ExpiryDate],T0.[AddID], T0.[validFor],DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE()
FROM '+ @SAPDB +' ..OCRD T0  
LEFT OUTER JOIN '+ @SAPDB +' ..OPLN T1 ON T0.[ListNum] = T1.[ListNum]
LEFT OUTER JOIN '+ @SAPDB +' ..OCRG T2 ON T0.[GroupCode] = T2.[GroupCode]
LEFT OUTER JOIN '+ @SAPDB +' ..OCRY T3 ON T0.[Country] = T3.[Code]
WHERE T0.[CardCode] NOT IN (SELECT CardCode  FROM [AB_Customers] ) AND T0.CardType = ''C''
 ORDER BY  T0.[CardCode]'

SET @SQL1 ='
UPDATE
     AB_Customers 
SET
     AB_Customers.CardName = T0.[CardName],
	  AB_Customers.PriceList = T1.[ListNum],
	   AB_Customers.PriceListName = T1.[ListName],
	    AB_Customers.Phone1 = T0.[Phone1],
		AB_Customers.Mobile = T0.[Cellular],
		AB_Customers.Email = T0.[E_Mail],
		AB_Customers.Address1 = T0.[Address],
		AB_Customers.Address2 = T0.[StreetNo],
		AB_Customers.Address3 = T0.[Block],
		AB_Customers.Country = T3.[Name],
		AB_Customers.Zipcode = T0.[ZipCode],
		AB_Customers.DOB = T0.[U_AB_DOB],
		AB_Customers.JoinDate = T0.[U_AB_JoinDate],
		AB_Customers.ExpiryDate = T0.[U_AB_ExpiryDate],
		AB_Customers.POSSearch = T0.[AddID],
		AB_Customers.Active = T0.[validFor],
		  AB_Customers.SAPSyncDate = DATEADD(day,datediff(day,0,GETDATE()),0),
			AB_Customers.SAPSyncDateTime = GETDATE()
FROM  AB_Customers TT
LEFT OUTER JOIN '+ @SAPDB +' ..OCRD T0 ON T0.CardCode = TT.CardCode 
LEFT OUTER JOIN '+ @SAPDB +' ..OPLN T1 ON T0.[ListNum] = T1.[ListNum]
LEFT OUTER JOIN '+ @SAPDB +' ..OCRG T2 ON T0.[GroupCode] = T2.[GroupCode] 
LEFT OUTER JOIN '+ @SAPDB +' ..OCRY T3 ON T0.[Country] = T3.[Code]

WHERE
   T0.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ') AND T0.CardType = ''C'''


--PRINT @SQL
--PRINT @SQL1
--LEFT OUTER JOIN '+ @SAPDB +' ..OCRG T2 ON T0.[GroupCode] = T2.[GroupCode] 

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP008_CustomerGroupSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--[dbo].[AE_SP008_CustomerGroupSync] 'SBODemoSG',2


ALTER procedure [dbo].[AE_SP008_CustomerGroupSync]
@SAPDB as varchar(50),
@Var as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)

SET @SQL = '
INSERT INTO [AB_CustomerGroup]  ([PriceList], [PriceListName],[SAPSyncDate],[SAPSyncDateTime])
SELECT T1.[ListNum] [PriceList], T1.[ListName] [PriceListName],
DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE() 
FROM '+ @SAPDB +' ..OPLN T1 
WHERE T1.[ListNum] NOT IN (SELECT [AB_CustomerGroup].PriceList  FROM [AB_CustomerGroup] ) 
AND T1.GroupCode <> 1
ORDER BY  T1.[ListNum]'


SET @SQL1 ='
UPDATE
     AB_CustomerGroup 
SET
      AB_CustomerGroup.PriceList = T1.[ListNum],
	   AB_CustomerGroup.PriceListName = T1.ListName,
	   	AB_CustomerGroup.SAPSyncDate = DATEADD(day,datediff(day,0,GETDATE()),0),
		 AB_CustomerGroup.SAPSyncDateTime = GETDATE()
FROM  AB_CustomerGroup T0
LEFT OUTER JOIN '+ @SAPDB +' ..OPLN T1 ON T0.[PriceList] = T1.[ListNum]
WHERE
    T1.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ') AND T1.GroupCode <> 1'


--PRINT @SQL
--PRINT @SQL1

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP009_BrandSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--[AE_SP004_PriceListSync]'SBODemoSG','2','7'

ALTER procedure [dbo].[AE_SP009_BrandSync]
@SAPDB as varchar(50)


as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)


SET @SQL = '
DELETE FROM [AB_BRAND]'

SET @SQL = '
INSERT INTO [AB_BRAND]  ([BrandCode],[BrandName],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.[FirmCode],T0.[FirmName],DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE()
FROM '+ @SAPDB +' ..OMRC T0'

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP010_CategorySync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER procedure [dbo].[AE_SP010_CategorySync]
@SAPDB as varchar(50)


as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)


SET @SQL = '
DELETE FROM [AB_CATEGORY]'

SET @SQL = '
INSERT INTO [AB_CATEGORY]  ([CategoryCode],[CategoryName],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.[ItmsTypCod],T0.[ItmsGrpNam],DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE()
FROM '+ @SAPDB +' ..OITG T0 order by T0.[ItmsTypCod]'

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP011_DepartmentSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER procedure [dbo].[AE_SP011_DepartmentSync]
@SAPDB as varchar(50)


as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)


SET @SQL = '
DELETE FROM [AB_DEPARTMENT]'

SET @SQL = '
INSERT INTO [AB_DEPARTMENT]  ([DepartmentCode],[DepartmentName],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.[ItmsGrpCod],T0.[ItmsGrpNam],DATEADD(day,datediff(day,0,GETDATE()),0),GETDATE()
FROM '+ @SAPDB +' ..OITB T0 order by T0.[ItmsGrpCod]'

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP012_VendorsSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


--[dbo].[AE_SP006_WareHouseSync]'SBODemoSG',2
ALTER procedure [dbo].[AE_SP012_VendorsSync]
@SAPDB as varchar(50),
@Var as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)

SET @SQL = '
INSERT INTO [AB_Vendors]  ( [VendorCode],[VendorName],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.CardCode , T0.CardName , DATEADD(day,datediff(day,0,GETDATE()),0), GETDATE() FROM '+ @SAPDB +' ..OCRD T0
WHERE T0.CardCode NOT IN (SELECT VendorCode FROM AB_Vendors) AND T0.CardType = ''S'''

SET @SQL1 ='
UPDATE
     AB_Vendors
SET
     AB_Vendors.VendorName = T1.CardName,
	   	  AB_Vendors.SAPSyncDate = DATEADD(day,datediff(day,0,GETDATE()),0),
			AB_Vendors.SAPSyncDateTime = GETDATE()
FROM  AB_Vendors T0
LEFT OUTER JOIN '+ @SAPDB +' ..OCRD T1 ON T0.VendorCode = T1.CardCode 
WHERE
   T1.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')  AND T1.CardType = ''S'''


--PRINT @SQL
--PRINT @SQL1

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end
GO
/****** Object:  StoredProcedure [dbo].[AE_SP013_TenderSync]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



--[dbo].[AE_SP012_TenderSync]'SBODemoSG',2
ALTER procedure [dbo].[AE_SP013_TenderSync]
@SAPDB as varchar(50),
@Var as varchar(10)

as
begin

DECLARE @SQL VARCHAR(MAX)
DECLARE @SQL1 VARCHAR(MAX)

SET @SQL = '
INSERT INTO [AB_Tender]  ( [TenderCode],[TenderName],[SAPSyncDate],[SAPSyncDateTime])
SELECT T0.CreditCard , T0.CardName , DATEADD(day,datediff(day,0,GETDATE()),0), GETDATE() FROM '+ @SAPDB +' ..OCRC T0
WHERE T0.CreditCard NOT IN (SELECT TenderCode FROM AB_Tender)'

SET @SQL1 ='
UPDATE
     AB_Tender
SET
     AB_Tender.TenderName = T1.CardName,
	   	  AB_Tender.SAPSyncDate = DATEADD(day,datediff(day,0,GETDATE()),0),
			AB_Tender.SAPSyncDateTime = GETDATE()
FROM  AB_Tender T0
LEFT OUTER JOIN '+ @SAPDB +' ..OCRC T1 ON T0.TenderCode = T1.CreditCard 
WHERE
   T1.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- ' + @Var + ')'


--PRINT @SQL
--PRINT @SQL1

EXEC(@SQL)
EXEC(@SQL1)
	 	 
end

GO
/****** Object:  Table [dbo].[AB_Brand]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

DROP TABLE [dbo].[AB_Brand]
GO

CREATE TABLE [dbo].[AB_Brand](
	[BrandCode] [int] NOT NULL,
	[BrandName] [nvarchar](30) NOT NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_Category]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

DROP TABLE [dbo].[AB_Category]
GO

CREATE TABLE [dbo].[AB_Category](
	[CategoryCode] [int] NOT NULL,
	[CategoryName] [nvarchar](50) NOT NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_CustomerGroup]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

DROP TABLE [dbo].[AB_CustomerGroup]
GO

CREATE TABLE [dbo].[AB_CustomerGroup](
	[PriceList] [int] NULL,
	[PriceListName] [nvarchar](50) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_Customers]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

DROP TABLE [dbo].[AB_Customers]
GO

CREATE TABLE [dbo].[AB_Customers](
	[CardCode] [nvarchar](15) NULL,
	[CardName] [nvarchar](100) NULL,
	[PriceList] [int] NULL,
	[PriceListName] [nvarchar](50) NULL,
	[Phone1] [nvarchar](20) NULL,
	[Mobile] [nvarchar](50) NULL,
	[Email] [nvarchar](100) NULL,
	[Address1] [nvarchar](100) NULL,
	[Address2] [nvarchar](100) NULL,
	[Address3] [nvarchar](100) NULL,
	[Country] [nvarchar](100) NULL,
	[Zipcode] [nvarchar](20) NULL,
	[DOB] [datetime] NULL,
	[JoinDate] [datetime] NULL,
	[ExpiryDate] [datetime] NULL,
	[POSSearch] [nvarchar](64) NULL,
	[Active] [nvarchar](1) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_Department]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_Department]
GO

CREATE TABLE [dbo].[AB_Department](
	[DepartmentCode] [int] NOT NULL,
	[DepartmentName] [nvarchar](50) NOT NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_ItemMaster]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_ItemMaster]
GO
CREATE TABLE [dbo].[AB_ItemMaster](
	[ItemCode] [nvarchar](20) NULL,
	[ItemName] [nvarchar](50) NULL,
	[Brand] [nvarchar](30) NULL,
	[Model] [nvarchar](16) NULL,
	[Category] [nvarchar](20) NULL,
	[Department] [nvarchar](50) NULL,
	[Vendor] [nvarchar](100) NULL,
	[Barcode] [nvarchar](16) NULL,
	[Active] [nvarchar](1) NULL,
	[UOM] [nvarchar](100) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_NoStockItem]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
DROP TABLE [dbo].[AB_NoStockItem]
GO
CREATE TABLE [dbo].[AB_NoStockItem](
	[ID] [int] NOT NULL,
	[LineID] [int] NULL,
	[Outlet] [nvarchar](50) NULL,
	[ItemCode] [nvarchar](100) NULL,
	[SellItem] [char](1) NULL,
	[InvntItem] [char](1) NULL,
	[Quantity] [numeric](19, 6) NULL,
	[ReqQty] [numeric](19, 6) NULL,
	[ChildItemCode] [nvarchar](100) NULL,
	[BOMQty] [numeric](38, 6) NULL,
	[OnHand] [numeric](19, 6) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[AB_Payment]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_Payment]
GO
CREATE TABLE [dbo].[AB_Payment](
	[ID] [int] NOT NULL,
	[POSTxNo] [nvarchar](30) NOT NULL,
	[PaymentCode] [nvarchar](30) NOT NULL,
	[PaymentAmount] [numeric](19, 6) NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_PriceList]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_PriceList]
GO
CREATE TABLE [dbo].[AB_PriceList](
	[ItemCode] [nvarchar](20) NULL,
	[PriceList] [int] NULL,
	[PriceListName] [nvarchar](50) NULL,
	[Currency] [nvarchar](10) NULL,
	[Price] [numeric](19, 6) NULL,
	[PriceGST] [numeric](19, 6) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_Promotion]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_Promotion]
GO
CREATE TABLE [dbo].[AB_Promotion](
	[ItemCode] [nvarchar](20) NULL,
	[PriceList] [int] NULL,
	[PriceListName] [nvarchar](50) NULL,
	[Currency] [nvarchar](10) NULL,
	[Price] [numeric](19, 6) NULL,
	[PriceGST] [numeric](19, 6) NULL,
	[FromDate] [datetime] NULL,
	[ToDate] [datetime] NULL,
	[CreateDate] [datetime] NULL,
	[UpdateDate] [datetime] NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_SalesTransDetail]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_SalesTransDetail]
GO
CREATE TABLE [dbo].[AB_SalesTransDetail](
	[ID] [int] NOT NULL,
	[POSTxNo] [nvarchar](30) NOT NULL,
	[Outlet] [nvarchar](50) NOT NULL,
	[ItemCode] [nvarchar](100) NULL,
	[Quantity] [numeric](19, 6) NOT NULL,
	[UnitPrice] [numeric](19, 6) NULL,
	[DiscAmount] [numeric](19, 6) NULL,
	[LineTotal] [numeric](19, 6) NOT NULL,
	[TotalGST] [numeric](19, 6) NOT NULL,
	[ErrMsg] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_SalesTransHeader]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_SalesTransHeader]
GO
CREATE TABLE [dbo].[AB_SalesTransHeader](
	[ID] [int] NOT NULL,
	[Outlet] [nvarchar](50) NOT NULL,
	[CardCode] [nvarchar](50) NULL,
	[POSTxNo] [nvarchar](100) NOT NULL,
	[Salesman] [int] NULL,
	[POSTillId] [nvarchar](100) NULL,
	[POSTxDate] [datetime] NOT NULL,
	[POSTxDatetime] [datetime] NOT NULL,
	[POSTxType] [nvarchar](5) NOT NULL,
	[DiscAmount] [numeric](19, 6) NULL,
	[TotalGST] [numeric](19, 6) NULL,
	[DocTotal] [numeric](19, 6) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDatetime] [datetime] NULL,
	[Status] [nvarchar](20) NULL,
	[ErrorMsg] [nvarchar](max) NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL,
	[ARDocEntry] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_Tender]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_Tender]
GO
CREATE TABLE [dbo].[AB_Tender](
	[TenderCode] [int] NOT NULL,
	[TenderName] [nvarchar](50) NOT NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_Vendors]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_Vendors]
GO
CREATE TABLE [dbo].[AB_Vendors](
	[VendorCode] [nvarchar](15) NOT NULL,
	[VendorName] [nvarchar](100) NOT NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_Warehouses]    Script Date: 12/21/2015 5:33:49 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
DROP TABLE [dbo].[AB_Warehouses]
GO
CREATE TABLE [dbo].[AB_Warehouses](
	[WhsCode] [nvarchar](8) NULL,
	[WhsName] [nvarchar](50) NULL,
	[Active] [nvarchar](1) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO

