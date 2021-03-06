USE [INTDB_RankingSports]
GO
/****** Object:  Table [dbo].[AB_Customers]    Script Date: 6/22/2015 2:53:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AB_Customers](
	[CardCode] [nvarchar](15) NULL,
	[CardName] [nvarchar](100) NULL,
	[GroupName] [nvarchar](20) NULL,
	[PriceListName] [nvarchar](50) NULL,
	[Phone1] [nvarchar](20) NULL,
	[POSSearch] [nvarchar](64) NULL,
	[Active] [nvarchar](1) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_ItemMaster]    Script Date: 6/22/2015 2:53:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
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
/****** Object:  Table [dbo].[AB_NoStockItem]    Script Date: 6/22/2015 2:53:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
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
/****** Object:  Table [dbo].[AB_Payment]    Script Date: 6/22/2015 2:53:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AB_Payment](
	[ID] [int] NOT NULL,
	[HeaderID] [int] NOT NULL,
	[PaymentCode] [nvarchar](30) NOT NULL,
	[PaymentAmount] [numeric](19, 6) NOT NULL
) ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_PriceList]    Script Date: 6/22/2015 2:53:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AB_PriceList](
	[ItemCode] [nvarchar](20) NULL,
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
/****** Object:  Table [dbo].[AB_Promotion]    Script Date: 6/22/2015 2:53:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AB_Promotion](
	[ItemCode] [nvarchar](20) NULL,
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
/****** Object:  Table [dbo].[AB_SalesTransDetail]    Script Date: 6/22/2015 2:53:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AB_SalesTransDetail](
	[ID] [int] NOT NULL,
	[HeaderID] [int] NOT NULL,
	[Outlet] [nvarchar](50) NOT NULL,
	[ItemCode] [nvarchar](100) NULL,
	[PriceBefDi] [numeric](19, 6) NULL,
	[DiscPrcnt] [numeric](19, 6) NULL,
	[Price] [numeric](19, 6) NOT NULL,
	[Quantity] [numeric](19, 6) NOT NULL,
	[LineTotal] [numeric](19, 6) NOT NULL,
	[ErrMsg] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_SalesTransHeader]    Script Date: 6/22/2015 2:53:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[AB_SalesTransHeader](
	[ID] [int] NOT NULL,
	[Outlet] [nvarchar](50) NOT NULL,
	[POSTxNo] [nvarchar](100) NOT NULL,
	[POSTillId] [nvarchar](100) NULL,
	[POSTxDate] [datetime] NOT NULL,
	[POSTxDatetime] [datetime] NOT NULL,
	[POSTxType] [nvarchar](5) NOT NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDatetime] [datetime] NULL,
	[Status] [nvarchar](20) NULL,
	[ErrorMsg] [nvarchar](max) NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL,
	[ARDocEntry] [int] NULL,
	[CardCode] [nvarchar](50) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO
/****** Object:  Table [dbo].[AB_Warehouses]    Script Date: 6/22/2015 2:53:52 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
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
