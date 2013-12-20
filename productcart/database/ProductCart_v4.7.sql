/****** Object:  StoredProcedure [dbo].[uspInActivatePrds]    Script Date: 1/10/2012 17:12:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspInActivatePrds]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspInActivatePrds]    Script Date: 02/25/2011 22:51:46 ******/
CREATE PROCEDURE [dbo].[uspInActivatePrds]
@SCID nvarchar(10) ,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query=''UPDATE Products SET Products.pcSC_ID=Products.active,Products.active=0 FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID='' + @SCID + '';''
	EXEC(@query)
		
	SET @SMCount=@@ROWCOUNT

END
' 
END
GO
/****** Object:  StoredProcedure [dbo].[uspGetUpdatedPrdCount]    Script Date: 1/10/2012 17:12:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspGetUpdatedPrdCount]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspGetUpdatedPrdCount]    Script Date: 02/25/2011 22:51:04 ******/
CREATE PROCEDURE [dbo].[uspGetUpdatedPrdCount]
@SCID nvarchar(10) ,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query=''SELECT idProduct FROM Products WHERE pcSC_ID='' + @SCID + '';''
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT

END
' 
END
GO
/****** Object:  StoredProcedure [dbo].[uspGetPrdCount]    Script Date: 1/10/2012 17:12:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspGetPrdCount]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspGetPrdCount]    Script Date: 02/25/2011 22:50:27 ******/
CREATE PROCEDURE [dbo].[uspGetPrdCount]
@Param1 nvarchar(1000) ,
@Param2 nvarchar(1000),
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query=''SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '';''
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT

END
' 
END
GO
/****** Object:  StoredProcedure [dbo].[uspChangePrices]    Script Date: 1/10/2012 17:12:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspChangePrices]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspChangePrices]    Script Date: 02/25/2011 22:49:56 ******/
CREATE PROCEDURE [dbo].[uspChangePrices]
@Param1 nvarchar(1000),
@Param2 nvarchar(1000),
@TPrice nvarchar(10),
@CType nvarchar(10),
@Relative nvarchar(10),
@Amount nvarchar(10),
@CRound nvarchar(10),
@SCID nvarchar(10),
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	DECLARE @HasTmp int
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query=''''
	
	IF @CType=''0''
	BEGIN
	
		IF @TPrice=''0''
		BEGIN
			IF @CRound=''0''
				SET @query=''UPDATE Products SET Products.Price=Round(Products.Price*'' + @Amount + '',2)''
			ELSE
				SET @query=''UPDATE Products SET Products.Price=Round(Products.Price*'' + @Amount + '',0)''
			
			SET @query=@query + '' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');''
		END
		
		IF @TPrice=''-1''
		BEGIN
			IF @CRound=''0''
				SET @query=''UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*'' + @Amount + '',2) ELSE Round(Products.bToBPrice*'' + @Amount + '',2) END''
			ELSE
				SET @query=''UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*'' + @Amount + '',0) ELSE Round(Products.bToBPrice*'' + @Amount + '',0) END''
			
			SET @query=@query + '' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');''
		END
		
		IF (@TPrice<>''-1'') AND (@TPrice<>''0'')
		BEGIN
			IF @CRound=''0''
				SET @query=''UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price*'' + @Amount + '',2)''
			ELSE
				SET @query=''UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price*'' + @Amount + '',0)''
			
			SET @query=@query + '' WHERE (pcCC_Pricing.idcustomerCategory='' + @TPrice + '') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + ''));''
		END
		
	END
	
	IF @CType=''1''
	BEGIN
	
		IF @TPrice=''0''
		BEGIN
			IF @CRound=''0''
				SET @query=''UPDATE Products SET Products.Price=Round(Products.Price-'' + @Amount + '',2)''
			ELSE
				SET @query=''UPDATE Products SET Products.Price=Round(Products.Price-'' + @Amount + '',0)''
			
			SET @query=@query + '' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');''
		END
		
		IF @TPrice=''-1''
		BEGIN
			IF @CRound=''0''
				SET @query=''UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price-'' + @Amount + '',2) ELSE Round(Products.bToBPrice-'' + @Amount + '',2) END''
			ELSE
				SET @query=''UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price-'' + @Amount + '',0) ELSE Round(Products.bToBPrice-'' + @Amount + '',0) END''
			
			SET @query=@query + '' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');''
		END
		
		IF (@TPrice<>''-1'') AND (@TPrice<>''0'')
		BEGIN
			IF @CRound=''0''
				SET @query=''UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price-'' + @Amount + '',2)''
			ELSE
				SET @query=''UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price-'' + @Amount + '',0)''
			
			SET @query=@query + '' WHERE (pcCC_Pricing.idcustomerCategory='' + @TPrice + '') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + ''));''
		END
		
	END
		
	IF @query<>''''
	BEGIN
		EXEC(@query)
		SET @SMCount=@@ROWCOUNT
		
		SET @query=''UPDATE Products SET Products.pcSC_ID='' + @SCID + '' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');'' 
		EXEC(@query)
	END
	
	IF @CType=''2''
	BEGIN
		
		SET @HasTmp=0
		
		IF @Relative=''0''
		BEGIN
			SET @HasTmp=1
			IF @CRound=''0''
				SET @query=''SELECT Products.idProduct,NewPrice = Round(Products.Price*'' + @Amount + '',2) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');'' 
			ELSE
				SET @query=''SELECT Products.idProduct,NewPrice = Round(Products.Price*'' + @Amount + '',0) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');''
		END
		
		IF @Relative=''-1''
		BEGIN
			SET @HasTmp=1
			IF @CRound=''0''
				SET @query=''SELECT Products.idProduct,NewPrice = CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*'' + @Amount + '',2) ELSE Round(Products.bToBPrice*'' + @Amount + '',2) END INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');'' 
			ELSE
				SET @query=''SELECT Products.idProduct,NewPrice = CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*'' + @Amount + '',0) ELSE Round(Products.bToBPrice*'' + @Amount + '',0) END INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');''
		END
		
		IF @Relative=''-2''
		BEGIN
			SET @HasTmp=1
			IF @CRound=''0''
				SET @query=''SELECT Products.idProduct,NewPrice = Round(Products.listPrice*'' + @Amount + '',2) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');'' 
			ELSE
				SET @query=''SELECT Products.idProduct,NewPrice = Round(Products.listPrice*'' + @Amount + '',0) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');''
		END
		
		IF @Relative=''-3''
		BEGIN
			SET @HasTmp=1
			IF @CRound=''0''
				SET @query=''SELECT Products.idProduct,NewPrice = Round(Products.cost*'' + @Amount + '',2) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');'' 
			ELSE
				SET @query=''SELECT Products.idProduct,NewPrice = Round(Products.cost*'' + @Amount + '',0) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');''
		END
		
		IF @HasTmp=0
		BEGIN
			IF @CRound=''0''
				SET @query=''SELECT pcCC_Pricing.idProduct,NewPrice = Round(pcCC_Pricing.pcCC_Price*'' + @Amount + '',2) INTO tmpSale1 FROM pcCC_Pricing WHERE (pcCC_Pricing.idcustomerCategory='' + @Relative + '') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + ''));'' 
			ELSE
				SET @query=''SELECT pcCC_Pricing.idProduct,NewPrice = Round(pcCC_Pricing.pcCC_Price*'' + @Amount + '',0) INTO tmpSale1 FROM pcCC_Pricing WHERE (pcCC_Pricing.idcustomerCategory='' + @Relative + '') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + ''));''
		END
		
		EXEC(@query)
		
		IF @TPrice=''0''
			SET @query=''UPDATE Products SET Products.Price=tmpSale1.NewPrice FROM Products,tmpSale1 WHERE Products.IdProduct=tmpSale1.IDProduct;''

		
		IF @TPrice=''-1''
			SET @query=''UPDATE Products SET Products.bToBPrice=tmpSale1.NewPrice FROM Products,tmpSale1 WHERE Products.IdProduct=tmpSale1.IDProduct;''
		
		IF (@TPrice<>''-1'') AND (@TPrice<>''0'')
			SET @query=''UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=tmpSale1.NewPrice FROM pcCC_Pricing,tmpSale1 WHERE (pcCC_Pricing.idcustomerCategory='' + @TPrice + '') AND (pcCC_Pricing.IdProduct=tmpSale1.IDProduct);''
		
		EXEC(@query)
		SET @SMCount=@@ROWCOUNT
		
		DROP TABLE tmpSale1
		
		SET @query=''UPDATE Products SET Products.pcSC_ID='' + @SCID + '' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + '');'' 
		EXEC(@query)
		
	END
	
END
' 
END
GO
/****** Object:  StoredProcedure [dbo].[uspBackUpPrices]    Script Date: 1/10/2012 17:12:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspBackUpPrices]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspBackUpPrices]    Script Date: 02/25/2011 22:49:18 ******/
CREATE PROCEDURE [dbo].[uspBackUpPrices]
@Param1 nvarchar(1000),
@Param2 nvarchar(1000),
@SCID nvarchar(7),
@SalesID nvarchar(7),
@TPrice nvarchar(10),
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	IF @TPrice=''0''
		SET @query=''INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT '' + @SCID + '','' + @SalesID + '','' + @TPrice + '',Products.idProduct,Products.Price FROM '' + @Param1 + '' WHERE '' + @Param2 + '';''
	ELSE
		BEGIN
		IF @TPrice=''-1''
			SET @query=''INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT '' + @SCID + '','' + @SalesID + '','' + @TPrice + '',Products.idProduct,Products.bToBPrice FROM '' + @Param1 + '' WHERE '' + @Param2 + '';''
		ELSE
			SET @query=''INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT '' + @SCID + '','' + @SalesID + '','' + @TPrice + '',pcCC_Pricing.idProduct,pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory='' + @TPrice + '' AND (pcCC_Pricing.IdProduct IN (SELECT Products.IdProduct FROM '' + @Param1 + '' WHERE '' + @Param2 + ''));''
		END
		
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT

END
' 
END
GO
/****** Object:  StoredProcedure [dbo].[uspAddCatPrices]    Script Date: 1/10/2012 17:12:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspAddCatPrices]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspAddCatPrices]    Script Date: 02/25/2011 22:48:41 ******/
CREATE PROCEDURE [dbo].[uspAddCatPrices]
@Param1 nvarchar(1000),
@Param2 nvarchar(1000),
@IDCat nvarchar(7),
@CAmount nvarchar(10),
@CType int,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	IF @CType=0
		SET @query=''INSERT INTO pcCC_Pricing (idcustomerCategory,IDProduct,pcCC_Price) SELECT '' + @IDCat + '',Products.idProduct,Round(Products.Price*'' + @CAmount + '',2) FROM '' + @Param1 + '' WHERE (Products.idProduct NOT IN (SELECT idProduct FROM pcCC_Pricing WHERE idcustomerCategory='' + @IDCat + '')) AND '' + @Param2 + '';''
	ELSE
		SET @query=''INSERT INTO pcCC_Pricing(idcustomerCategory,IDProduct,pcCC_Price) SELECT '' + @IDCat + '',Products.idProduct,WPrice = CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*'' + @CAmount + '',2) ELSE Round(Products.bToBPrice*'' + @CAmount + '',2) END FROM '' + @Param1 + '' WHERE (Products.idProduct NOT IN (SELECT idProduct FROM pcCC_Pricing WHERE idcustomerCategory='' + @IDCat + '')) AND '' + @Param2 + '';''
		
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT

END
' 
END
GO
/****** Object:  StoredProcedure [dbo].[uspActivatePrds]    Script Date: 1/10/2012 17:12:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspActivatePrds]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspActivatePrds]    Script Date: 02/25/2011 22:48:15 ******/
CREATE PROCEDURE [dbo].[uspActivatePrds]
@SCID nvarchar(10) ,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query=''UPDATE Products SET Products.active=Products.pcSC_ID,Products.pcSC_ID=0 FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID='' + @SCID + '';''
	EXEC(@query)	
	
	SET @SMCount=@@ROWCOUNT

END
' 
END
GO
/****** Object:  Table [dbo].[used_discounts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[used_discounts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[used_discounts](
	[iddiscount] [int] NULL,
	[idcustomer] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[ups_license]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[ups_license]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[ups_license](
	[idUPS] [int] NULL,
	[ups_UserId] [nvarchar](100) NULL,
	[ups_Password] [nvarchar](100) NULL,
	[ups_AccessLicense] [nvarchar](100) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[twoCheckout]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[twoCheckout]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[twoCheckout](
	[id_twocheckout] [int] IDENTITY(1,1) NOT NULL,
	[store_id] [nvarchar](50) NULL,
	[v2co] [int] NULL,
	[v2co_TestMode] [int] NULL,
 CONSTRAINT [aaaaatwoCheckout_PK] PRIMARY KEY NONCLUSTERED 
(
	[id_twocheckout] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[tclink]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[tclink]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[tclink](
	[idTCLink] [int] NULL,
	[TCLinkid] [nvarchar](255) NULL,
	[TCLinkPassword] [nvarchar](255) NULL,
	[CVV] [int] NULL,
	[TCTestmode] [int] NULL,
	[TCLinkCheck] [int] NULL,
	[TCLinkCheckPending] [int] NULL,
	[TCLinkecheck] [nvarchar](50) NULL,
	[TCCurcode] [nvarchar](50) NULL,
	[TranType] [nvarchar](255) NULL,
	[avs] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[taxPrd]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[taxPrd]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[taxPrd](
	[idTaxPerProduct] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[countryCode] [nvarchar](4) NULL,
	[countryCodeEq] [int] NULL,
	[stateCode] [nvarchar](4) NULL,
	[stateCodeEq] [int] NULL,
	[zip] [nvarchar](12) NULL,
	[zipEq] [int] NULL,
	[taxPerProduct] [float] NULL,
 CONSTRAINT [aaaaataxPrd_PK] PRIMARY KEY NONCLUSTERED 
(
	[idTaxPerProduct] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[taxLoc]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[taxLoc]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[taxLoc](
	[idTaxPerPlace] [int] IDENTITY(1,1) NOT NULL,
	[countryCode] [nvarchar](4) NULL,
	[countryCodeEq] [int] NULL,
	[stateCode] [nvarchar](4) NULL,
	[stateCodeEq] [int] NULL,
	[zip] [nvarchar](12) NULL,
	[zipEq] [int] NULL,
	[taxLoc] [float] NULL,
	[taxDesc] [nvarchar](50) NULL,
 CONSTRAINT [aaaaataxLoc_PK] PRIMARY KEY NONCLUSTERED 
(
	[idTaxPerPlace] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[suppliers]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[suppliers]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[suppliers](
	[idSupplier] [int] NOT NULL,
	[supplierName] [nvarchar](50) NULL,
	[supplierEmail] [nvarchar](150) NULL,
	[supplierAddress] [nvarchar](150) NULL,
	[supplierZip] [nvarchar](50) NULL,
	[supplierCity] [nvarchar](50) NULL,
	[supplierStateCode] [nvarchar](4) NULL,
	[supplierAnotherState] [nvarchar](50) NULL,
	[supplierCountryCode] [nvarchar](4) NULL,
	[supplierPhone] [nvarchar](50) NULL,
	[receiveSellEmail] [int] NULL,
	[receiveUnderStockAlert] [int] NULL,
 CONSTRAINT [aaaaasuppliers_PK] PRIMARY KEY NONCLUSTERED 
(
	[idSupplier] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[states]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[states]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[states](
	[stateCode] [nvarchar](4) NOT NULL,
	[stateName] [nvarchar](250) NULL,
	[pcCountryCode] [nvarchar](20) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[shipService]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[shipService]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[shipService](
	[idshipservice] [int] IDENTITY(1,1) NOT NULL,
	[serviceActive] [int] NOT NULL,
	[serviceCode] [nvarchar](50) NOT NULL,
	[servicePriority] [int] NOT NULL,
	[serviceDescription] [nvarchar](75) NULL,
	[serviceFree] [int] NULL,
	[serviceFreeOverAmt] [money] NULL,
	[serviceHandlingFee] [money] NULL,
	[serviceHandlingIntFee] [money] NULL,
	[serviceShowHandlingFee] [int] NOT NULL,
	[serviceLimitation] [int] NULL,
	[serviceDefaultRate] [money] NULL,
 CONSTRAINT [aaaaashipService_PK] PRIMARY KEY NONCLUSTERED 
(
	[idshipservice] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[ShipmentTypes]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[ShipmentTypes]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[ShipmentTypes](
	[idShipment] [int] NOT NULL,
	[shipmentDesc] [nvarchar](70) NULL,
	[priceToAdd] [float] NULL,
	[active] [int] NULL,
	[international] [bit] NOT NULL,
	[ipriceToAdd] [float] NULL,
	[shipserver] [nvarchar](250) NULL,
	[userID] [nvarchar](50) NULL,
	[password] [nvarchar](50) NULL,
	[AccessLicense] [nvarchar](250) NULL,
	[FedExKey] [nvarchar](250) NULL,
	[FedExPwd] [nvarchar](250) NULL,
 CONSTRAINT [aaaaaShipmentTypes_PK] PRIMARY KEY NONCLUSTERED 
(
	[idShipment] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[shipAlert]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[shipAlert]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[shipAlert](
	[idShipAlert] [int] IDENTITY(1,1) NOT NULL,
	[shipExists] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[SB_Settings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[SB_Settings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[SB_Settings](
	[Setting_ID] [int] IDENTITY(1,1) NOT NULL,
	[Setting_APIUser] [nvarchar](50) NULL,
	[Setting_APIPassword] [nvarchar](50) NULL,
	[Setting_APIKey] [nvarchar](250) NULL,
	[Setting_AutoReg] [int] NULL,
	[Setting_RegSuccess] [int] NULL,
	[Setting_LastCustomerID] [nvarchar](50) NULL,
 CONSTRAINT [PK__SB_Settings__697C9932] PRIMARY KEY CLUSTERED 
(
	[Setting_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[SB_Packages]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[SB_Packages]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[SB_Packages](
	[SB_PackageID] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[SB_LinkID] [nvarchar](50) NULL,
	[SB_IsLinked] [int] NULL,
	[SB_RefName] [nvarchar](50) NULL,
	[SB_Amount] [float] NULL,
	[SB_BillingPeriod] [nvarchar](50) NULL,
	[SB_BillingFrequency] [nvarchar](50) NULL,
	[SB_BillingCycles] [int] NULL,
	[SB_CurrencyCode] [nvarchar](4) NULL,
	[SB_IsTrial] [int] NULL,
	[SB_TrialAmount] [float] NULL,
	[SB_TrialBillingPeriod] [nvarchar](50) NULL,
	[SB_TrialBillingFrequency] [nvarchar](50) NULL,
	[SB_TrialBillingCycles] [int] NULL,
	[SB_StartsImmediately] [int] NULL,
	[SB_StartDate] [nvarchar](50) NULL,
	[SB_Agree] [int] NULL,
	[SB_AgreeText] [nvarchar](4000) NULL,
	[SB_Type] [int] NULL,
	[SB_ShowTrialPrice] [int] NULL,
	[SB_TrialDesc] [nvarchar](4000) NULL,
	[SB_ShowFreeTrial] [int] NULL,
	[SB_ShowStartDate] [int] NULL,
	[SB_StartDateDesc] [nvarchar](4000) NULL,
	[SB_ShowReoccurenceDate] [int] NULL,
	[SB_ReoccurenceDesc] [nvarchar](4000) NULL,
	[SB_ShowEOSDate] [int] NULL,
	[SB_EOSDesc] [nvarchar](4000) NULL,
	[SB_ShowTrialDate] [int] NULL,
	[SB_TrialDate] [nvarchar](50) NULL,
	[SB_FreeTrialDesc] [nvarchar](50) NULL,
 CONSTRAINT [PK__SB_Packages__5575A085] PRIMARY KEY CLUSTERED 
(
	[SB_PackageID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[SB_Orders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[SB_Orders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[SB_Orders](
	[SB_OrderID] [int] IDENTITY(1,1) NOT NULL,
	[idOrder] [int] NULL,
	[SB_GUID] [nvarchar](50) NULL,
	[SB_Terms] [nvarchar](500) NULL,
 CONSTRAINT [PK__SB_Orders__66A02C87] PRIMARY KEY CLUSTERED 
(
	[SB_OrderID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[Referrer]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[Referrer]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[Referrer](
	[IdRefer] [int] IDENTITY(1,1) NOT NULL,
	[Name] [nvarchar](150) NOT NULL,
	[sortOrder] [int] NULL,
	[Removed] [int] NULL,
 CONSTRAINT [aaaaaReferrer_PK] PRIMARY KEY NONCLUSTERED 
(
	[IdRefer] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[recipients]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[recipients]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[recipients](
	[idRecipient] [int] IDENTITY(1,1) NOT NULL,
	[idCustomer] [int] NULL,
	[recipient_FullName] [nvarchar](50) NULL,
	[recipient_Address] [nvarchar](150) NULL,
	[recipient_City] [nvarchar](50) NULL,
	[recipient_StateCode] [nvarchar](50) NULL,
	[recipient_State] [nvarchar](50) NULL,
	[recipient_Zip] [nvarchar](50) NULL,
	[recipient_CountryCode] [nvarchar](50) NULL,
	[recipient_Company] [nvarchar](150) NULL,
	[recipient_Address2] [nvarchar](150) NULL,
	[recipient_NickName] [nvarchar](50) NULL,
	[recipient_FirstName] [nvarchar](100) NULL,
	[recipient_LastName] [nvarchar](100) NULL,
	[recipient_Phone] [nvarchar](50) NULL,
	[recipient_Fax] [nvarchar](50) NULL,
	[recipient_Email] [nvarchar](250) NULL,
	[Recipient_Residential] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[PSIGate]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[PSIGate]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[PSIGate](
	[Config_File_Name] [nvarchar](150) NULL,
	[Config_File_Name_Full] [nvarchar](150) NULL,
	[Host] [nvarchar](150) NULL,
	[Port] [nvarchar](20) NULL,
	[Userid] [nvarchar](50) NULL,
	[Mode] [nvarchar](50) NULL,
	[id] [int] NULL,
	[psi_post] [nvarchar](10) NULL,
	[psi_testmode] [nvarchar](10) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[protx]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[protx]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[protx](
	[idProtx] [int] NOT NULL,
	[Protxid] [nvarchar](255) NULL,
	[ProtxPassword] [nvarchar](255) NULL,
	[CVV] [int] NOT NULL,
	[ProtxTestmode] [int] NOT NULL,
	[ProtxCurcode] [nvarchar](50) NULL,
	[TxType] [nvarchar](255) NULL,
	[avs] [int] NOT NULL,
	[ProtxMethod] [nvarchar](250) NULL,
	[ProtxCardTypes] [nvarchar](250) NULL,
	[ProtxApply3DSecure] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[ProductsOrdered]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[ProductsOrdered]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[ProductsOrdered](
	[idProductOrdered] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[idOrder] [int] NULL,
	[idProduct] [int] NULL,
	[service] [bit] NOT NULL,
	[quantity] [int] NULL,
	[idOptionA] [int] NULL,
	[idOptionB] [int] NULL,
	[unitPrice] [float] NULL,
	[unitCost] [float] NULL,
	[xfdetails] [ntext] NULL,
	[idconfigSession] [int] NULL,
	[rmaSubmitted] [int] NULL,
	[QDiscounts] [float] NULL,
	[ItemsDiscounts] [float] NULL,
	[pcPackageInfo_ID] [int] NULL,
	[pcDropShipper_ID] [int] NULL,
	[pcPrdOrd_BackOrder] [int] NULL,
	[pcPrdOrd_SentNotice] [int] NULL,
	[pcPrdOrd_SelectedOptions] [ntext] NULL,
	[pcPrdOrd_OptionsPriceArray] [ntext] NULL,
	[pcPrdOrd_OptionsArray] [ntext] NULL,
	[pcPO_EPID] [int] NULL,
	[pcPO_GWOpt] [int] NULL,
	[pcPO_GWNote] [nvarchar](250) NULL,
	[pcPO_GWPrice] [money] NULL,
	[pcPrdOrd_Shipped] [int] NULL,
	[pcPrdOrd_BundledDisc] [money] NULL,
	[pcPO_LinkID] [nvarchar](250) NULL,
	[pcPO_SubFrequency] [int] NOT NULL,
	[pcPO_SubPeriod] [nvarchar](20) NULL,
	[pcPO_SubCycles] [int] NOT NULL,
	[pcPO_SubTrialFrequency] [int] NOT NULL,
	[pcPO_SubTrialPeriod] [nvarchar](20) NULL,
	[pcPO_SubTrialCycles] [int] NOT NULL,
	[pcPO_IsTrial] [int] NOT NULL,
	[pcPO_SubAmount] [float] NOT NULL,
	[pcPO_SubTrialAmount] [int] NOT NULL,
	[pcPO_SubAgree] [int] NOT NULL,
	[pcPO_SubType] [nvarchar](20) NULL,
	[pcPO_NoShipping] [int] NOT NULL,
	[pcPO_SubStartDate] [nvarchar](50) NULL,
	[pcPO_SubDetails] [nvarchar](250) NULL,
	[pcPO_SubUPDStartDate] [nvarchar](50) NULL,
	[pcPO_SubActive] [int] NOT NULL,
	[pcSubscription_ID] [int] NOT NULL,
	[pcSC_ID] [int] NULL,
 CONSTRAINT [aaaaaProductsOrdered_PK] PRIMARY KEY NONCLUSTERED 
(
	[idProductOrdered] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[products]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[products]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[products](
	[idProduct] [int] IDENTITY(1,1) NOT NULL,
	[idSupplier] [int] NULL,
	[description] [nvarchar](255) NULL,
	[details] [ntext] NULL,
	[configOnly] [bit] NOT NULL,
	[serviceSpec] [bit] NOT NULL,
	[price] [float] NULL,
	[listPrice] [float] NULL,
	[bToBPrice] [float] NULL,
	[imageUrl] [nvarchar](150) NULL,
	[smallImageUrl] [nvarchar](150) NULL,
	[largeImageURL] [nvarchar](150) NULL,
	[sku] [nvarchar](100) NULL,
	[stock] [int] NULL,
	[listHidden] [int] NULL,
	[weight] [int] NULL,
	[deliveringTime] [int] NULL,
	[active] [int] NULL,
	[IdOptionGroupA] [int] NULL,
	[Arequired] [bit] NULL,
	[IdOptionGroupB] [int] NULL,
	[Brequired] [bit] NULL,
	[hotDeal] [int] NULL,
	[cost] [float] NULL,
	[visits] [int] NULL,
	[sales] [int] NULL,
	[emailText] [nvarchar](250) NULL,
	[stockLevelAlert] [int] NULL,
	[formQuantity] [int] NULL,
	[showInHome] [int] NULL,
	[rndNum] [int] NULL,
	[priority] [int] NULL,
	[notax] [int] NULL,
	[noshipping] [int] NULL,
	[removed] [int] NULL,
	[custom1] [int] NULL,
	[content1] [nvarchar](150) NULL,
	[custom2] [int] NULL,
	[content2] [nvarchar](150) NULL,
	[custom3] [int] NULL,
	[content3] [nvarchar](150) NULL,
	[xfield1] [int] NULL,
	[x1req] [int] NULL,
	[xfield2] [int] NULL,
	[x2req] [int] NULL,
	[xfield3] [int] NULL,
	[x3req] [int] NULL,
	[iRewardPoints] [int] NULL,
	[NoPrices] [int] NULL,
	[IDBrand] [int] NULL,
	[OverSizeSpec] [nvarchar](50) NULL,
	[Downloadable] [int] NULL,
	[sDesc] [ntext] NULL,
	[noStock] [int] NULL,
	[noshippingtext] [int] NULL,
	[pcprod_HideBTOPrice] [int] NULL,
	[pcprod_QtyValidate] [int] NULL,
	[pcprod_MinimumQty] [int] NULL,
	[pcprod_QtyToPound] [int] NULL,
	[pcprod_EnteredOn] [datetime] NULL,
	[pcprod_OrdInHome] [int] NULL,
	[pcprod_HideDefConfig] [int] NULL,
	[pcProdImage_Columns] [int] NULL,
	[pcProd_NotifyStock] [int] NULL,
	[pcProd_ReorderLevel] [int] NULL,
	[pcProd_SentNotice] [int] NULL,
	[pcSupplier_ID] [int] NULL,
	[pcProd_IsDropShipped] [int] NULL,
	[pcDropShipper_ID] [int] NULL,
	[pcProd_BackOrder] [int] NULL,
	[pcProd_ShipNDays] [int] NULL,
	[pcprod_GC] [int] NULL,
	[pcProd_SkipDetailsPage] [int] NULL,
	[pcProd_DisplayLayout] [nvarchar](150) NULL,
	[pcProd_MetaDesc] [ntext] NULL,
	[pcProd_MetaTitle] [nvarchar](250) NULL,
	[pcProd_MetaKeywords] [ntext] NULL,
	[pcProd_BTODefaultPrice] [money] NULL,
	[pcProd_BTODefaultWPrice] [money] NULL,
	[pcProd_HideSKU] [int] NULL,
	[pcProd_PrdNotes] [varchar](4000) NULL,
	[pcProd_EditedDate] [datetime] NULL,
	[pcProd_SavedTimes] [int] NOT NULL,
	[pcProd_Surcharge1] [money] NOT NULL,
	[pcProd_Surcharge2] [money] NOT NULL,
	[pcProd_multiQty] [int] NOT NULL,
	[pcProd_MaxSelect] [int] NOT NULL,
	[pcPrd_MojoZoom] [int] NULL,
	[pcProd_AvgRating] [float] NULL,
	[pcSC_ID] [int] NULL,
	[pcProd_GoogleCat] [nvarchar] (250) NULL,
	[pcProd_GoogleGender] [nvarchar] (250) NULL,
	[pcProd_GoogleAge] [nvarchar] (250) NULL,
	[pcProd_GoogleSize] [nvarchar] (250) NULL,
	[pcProd_GoogleColor] [nvarchar] (250) NULL,
	[pcProd_GooglePattern] [nvarchar] (250) NULL,
	[pcProd_GoogleMaterial] [nvarchar] (250) NULL,
	[pcProd_GoogleGroup] [nvarchar] (250) NULL,
 CONSTRAINT [aaaaaproducts_PK] PRIMARY KEY NONCLUSTERED 
(
	[idProduct] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[pfporders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pfporders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pfporders](
	[idpfporder] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[idOrder] [int] NULL,
	[amt] [money] NULL,
	[tender] [nvarchar](250) NULL,
	[trxtype] [nvarchar](250) NULL,
	[origid] [nvarchar](250) NULL,
	[acct] [nvarchar](250) NULL,
	[expdate] [nvarchar](250) NULL,
	[idCustomer] [int] NULL,
	[fullname] [nvarchar](250) NULL,
	[street] [nvarchar](250) NULL,
	[state] [nvarchar](250) NULL,
	[email] [nvarchar](250) NULL,
	[zip] [nvarchar](250) NULL,
	[captured] [int] NULL,
	[pcSecurityKeyID] [int] NOT NULL,
 CONSTRAINT [aaaaapfporders_PK] PRIMARY KEY NONCLUSTERED 
(
	[idpfporder] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[Permissions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[Permissions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[Permissions](
	[IdPm] [int] NOT NULL,
	[PmName] [nvarchar](150) NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcXMLSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcXMLSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcXMLSettings](
	[pcXMLSet_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcXMLSet_Log] [int] NULL,
	[pcXMLSet_LogErrors] [int] NULL,
	[pcXMLSet_CaptureRequest] [int] NULL,
	[pcXMLSet_CaptureResponse] [int] NULL,
	[pcXMLSet_EnforceHTTPs] [int] NULL,
 CONSTRAINT [Index_B7432DED_298C_46C3] PRIMARY KEY CLUSTERED 
(
	[pcXMLSet_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcXMLSettings]') AND name = N'Index_EFA3839E_D3EA_490A')
CREATE UNIQUE NONCLUSTERED INDEX [Index_EFA3839E_D3EA_490A] ON [dbo].[pcXMLSettings] 
(
	[pcXMLSet_ID] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcXMLPartners]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcXMLPartners]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcXMLPartners](
	[pcXP_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcXP_PartnerID] [nvarchar](50) NULL,
	[pcXP_Password] [nvarchar](50) NULL,
	[pcXP_Key] [nvarchar](30) NULL,
	[pcXP_Name] [nvarchar](150) NULL,
	[pcXP_Email] [nvarchar](150) NULL,
	[pcXP_Company] [nvarchar](150) NULL,
	[pcXP_Address] [nvarchar](250) NULL,
	[pcXP_Address2] [nvarchar](250) NULL,
	[pcXP_City] [nvarchar](150) NULL,
	[pcXP_StateCode] [nvarchar](50) NULL,
	[pcXP_Province] [nvarchar](50) NULL,
	[pcXP_Zip] [nvarchar](10) NULL,
	[pcXP_CountryCode] [nvarchar](50) NULL,
	[pcXP_Phone] [nvarchar](50) NULL,
	[pcXP_Fax] [nvarchar](50) NULL,
	[pcXP_Status] [int] NULL,
	[pcXP_Removed] [int] NULL,
	[pcXP_ExportAdmin] [int] NULL,
	[pcXP_FTPHost] [nvarchar](250) NULL,
	[pcXP_FTPDirectory] [nvarchar](250) NULL,
	[pcXP_FTPUsername] [nvarchar](50) NULL,
	[pcXP_FTPPassword] [nvarchar](50) NULL,
 CONSTRAINT [Index_4A0663DE_6810_4986] PRIMARY KEY CLUSTERED 
(
	[pcXP_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcXMLPartners]') AND name = N'Index_A4C9FAD5_10DE_40D6')
CREATE UNIQUE NONCLUSTERED INDEX [Index_A4C9FAD5_10DE_40D6] ON [dbo].[pcXMLPartners] 
(
	[pcXP_ID] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcXMLLogs]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcXMLLogs]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcXMLLogs](
	[pcXL_id] [int] IDENTITY(1,1) NOT NULL,
	[pcXP_id] [int] NULL,
	[pcXL_RequestKey] [nvarchar](30) NULL,
	[pcXL_RequestType] [int] NULL,
	[pcXL_UpdatedID] [int] NULL,
	[pcXL_BackupFile] [nvarchar](100) NULL,
	[pcXL_Undo] [int] NULL,
	[pcXL_ResultCount] [int] NULL,
	[pcXL_Date] [datetime] NULL,
	[pcXL_LastID] [int] NULL,
	[pcXL_UndoID] [int] NULL,
	[pcXL_Status] [int] NULL,
	[pcXL_RequestXML] [ntext] NULL,
	[pcXL_ResponseXML] [ntext] NULL,
 CONSTRAINT [Index_051BEF0D_ADA3_49BC] PRIMARY KEY CLUSTERED 
(
	[pcXL_id] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcXMLLogs]') AND name = N'Index_16CFF963_2E6C_4E62')
CREATE UNIQUE NONCLUSTERED INDEX [Index_16CFF963_2E6C_4E62] ON [dbo].[pcXMLLogs] 
(
	[pcXL_id] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcXMLIPs]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcXMLIPs]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcXMLIPs](
	[pcXIP_id] [int] IDENTITY(1,1) NOT NULL,
	[pcXIP_IPAddr] [nvarchar](20) NULL,
	[pcXIP_TurnOn] [int] NULL,
 CONSTRAINT [Index_4D740058_86CA_40AD] PRIMARY KEY CLUSTERED 
(
	[pcXIP_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcXMLIPs]') AND name = N'Index_63F73B3C_051A_48C2')
CREATE UNIQUE NONCLUSTERED INDEX [Index_63F73B3C_051A_48C2] ON [dbo].[pcXMLIPs] 
(
	[pcXIP_id] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcXMLExportLogs]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcXMLExportLogs]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcXMLExportLogs](
	[pcXEL_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcXP_ID] [int] NULL,
	[pcXEL_ExportedID] [int] NULL,
	[pcXEL_IDType] [int] NULL,
 CONSTRAINT [Index_4F4CBCA5_AE00_48C4] PRIMARY KEY CLUSTERED 
(
	[pcXEL_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcVATRates]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcVATRates]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcVATRates](
	[pcVATRate_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcVATRate_Category] [nvarchar](250) NULL,
	[pcVATRate_Rate] [money] NULL,
	[pcVATCountry_Code] [nvarchar](4) NULL,
 CONSTRAINT [PK__pcVATRates__1B9317B3] PRIMARY KEY CLUSTERED 
(
	[pcVATRate_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcVATCountries]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcVATCountries]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcVATCountries](
	[pcVATCountry_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcVATCountry_Code] [nvarchar](4) NULL,
	[pcVATCountry_State] [nvarchar](250) NULL,
 CONSTRAINT [Index_5C8B8D8C_FA7B_4A6B] PRIMARY KEY CLUSTERED 
(
	[pcVATCountry_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcVATCountries]') AND name = N'Index_C5C6380D_6691_44DE')
CREATE UNIQUE NONCLUSTERED INDEX [Index_C5C6380D_6691_44DE] ON [dbo].[pcVATCountries] 
(
	[pcVATCountry_ID] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcUPSPreferences]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcUPSPreferences]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcUPSPreferences](
	[pcUPSPref_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcUPSPref_Service] [nvarchar](20) NULL,
	[pcUPSPref_PackageType] [nvarchar](20) NULL,
	[pcUPSPref_PaymentMethod] [nvarchar](150) NULL,
	[pcUPSPref_AccountNumber] [nvarchar](50) NULL,
	[pcUPSPref_ReadyHours] [nvarchar](10) NULL,
	[pcUPSPref_ReadyMinutes] [nvarchar](10) NULL,
	[pcUPSPref_ReadyAMPM] [nvarchar](10) NULL,
	[pcUPSPref_PUHours] [nvarchar](10) NULL,
	[pcUPSPref_PUMinutes] [nvarchar](10) NULL,
	[pcUPSPref_RefNumber1] [nvarchar](10) NULL,
	[pcUPSPref_RefNumber2] [nvarchar](10) NULL,
	[pcUPSPref_RefData1] [nvarchar](150) NULL,
	[pcUPSPref_RefData2] [nvarchar](150) NULL,
	[pcUPSPref_CODPackage] [nvarchar](150) NULL,
	[pcUPSPref_CODAmount] [nvarchar](150) NULL,
	[pcUPSPref_CODCurrency] [nvarchar](150) NULL,
	[pcUPSPref_CODFunds] [nvarchar](150) NULL,
	[pcUPSPref_ShipmentNotification] [nvarchar](10) NULL,
	[pcUPSPref_NotifiCode1] [nvarchar](10) NULL,
	[pcUPSPref_NotifiCode2] [nvarchar](10) NULL,
	[pcUPSPref_NotifiCode3] [nvarchar](10) NULL,
	[pcUPSPref_NotifiCode4] [nvarchar](10) NULL,
	[pcUPSPref_NotifiCode5] [nvarchar](10) NULL,
	[pcUPSPref_NotifiEmail1] [nvarchar](250) NULL,
	[pcUPSPref_NotifiEmail2] [nvarchar](250) NULL,
	[pcUPSPref_NotifiEmail3] [nvarchar](250) NULL,
	[pcUPSPref_NotifiEmail4] [nvarchar](250) NULL,
	[pcUPSPref_NotifiEmail5] [nvarchar](250) NULL,
	[pcUPSPref_SaturdayDelivery] [nvarchar](10) NULL,
	[pcUPSPref_InsuredValue] [nvarchar](10) NULL,
	[pcUPSPref_VerbalConfirmation] [nvarchar](10) NULL,
 CONSTRAINT [PK__pcUPSPreferences__0B5CAFEA] PRIMARY KEY CLUSTERED 
(
	[pcUPSPref_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcUploadFiles]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcUploadFiles]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcUploadFiles](
	[pcUpld_IDFile] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcUpld_IDFeedback] [int] NULL,
	[pcUpld_FileName] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcTaxZonesGroups]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcTaxZonesGroups]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcTaxZonesGroups](
	[pcTaxZonesGroup_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcTaxZoneRate_ID] [int] NULL,
	[pcTaxZoneDesc_ID] [int] NULL
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcTaxZonesGroups]') AND name = N'pcTaxZone_ID')
CREATE NONCLUSTERED INDEX [pcTaxZone_ID] ON [dbo].[pcTaxZonesGroups] 
(
	[pcTaxZoneDesc_ID] ASC
) ON [PRIMARY]
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcTaxZonesGroups]') AND name = N'pcTaxZoneRate_ID')
CREATE NONCLUSTERED INDEX [pcTaxZoneRate_ID] ON [dbo].[pcTaxZonesGroups] 
(
	[pcTaxZoneRate_ID] ASC
) ON [PRIMARY]
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcTaxZonesGroups]') AND name = N'pcTaxZonesRates_ID')
CREATE NONCLUSTERED INDEX [pcTaxZonesRates_ID] ON [dbo].[pcTaxZonesGroups] 
(
	[pcTaxZonesGroup_ID] ASC
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[pcPay_PayPalAdvanced]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_PayPalAdvanced]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_PayPalAdvanced](
	[pcPay_PayPalAd_ID] [int] NOT NULL,
	[pcPay_PayPalAd_Partner] [nvarchar] (250) NULL,
	[pcPay_PayPalAd_MerchantLogin] [nvarchar] (250) NULL,
	[pcPay_PayPalAd_Vendor] [nvarchar] (255) NULL,
	[pcPay_PayPalAd_User] [nvarchar] (255) NULL,
	[pcPay_PayPalAd_Password] [nvarchar] (255) NULL,
	[pcPay_PayPalAd_TransType] [nvarchar] (10) NULL,
	[pcPay_PayPalAd_CSC] [nvarchar] (10) NULL,
	[pcPay_PayPalAd_Sandbox] [nvarchar] (10) NULL 
) ON [PRIMARY]
END
GO

/****** Object:  Table [dbo].[pcPay_PFL_Authorize]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_PFL_Authorize]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_PFL_Authorize](
	[idPFL_Authorize] [int] IDENTITY(1,1) NOT NULL,
	[idOrder] [int] NULL,
	[orderDate] [datetime] NULL,
	[paySource] [nvarchar] (250) NULL,
	[amount] [money] NULL,
	[paymentmethod] [nvarchar] (250) NULL,
	[transtype] [nvarchar] (250) NULL,
	[authcode] [nvarchar] (250) NULL,
	[captured] [int] NULL
) ON [PRIMARY]
END
GO

/****** Object:  Table [dbo].[pcTaxZones]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcTaxZones]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcTaxZones](
	[pcTaxZone_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcTaxZone_CountryCode] [nvarchar](50) NULL,
	[pcTaxZone_Province] [nvarchar](50) NULL,
	[pcTaxZone_PostalCode] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcTaxZoneRates]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcTaxZoneRates]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcTaxZoneRates](
	[pcTaxZoneRate_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcTaxZone_ID] [int] NULL,
	[pcTaxZoneRate_Name] [nvarchar](50) NULL,
	[pcTaxZoneRate_Type] [nvarchar](50) NULL,
	[pcTaxZoneRate_Order] [int] NULL,
	[pcTaxZoneRate_Rate] [float] NULL,
	[pcTaxZoneRate_ApplyToSH] [int] NULL,
	[pcTaxZoneRate_Taxable] [int] NULL,
	[pcTaxZoneRate_LocalZone] [int] NULL,
 CONSTRAINT [PrimaryKey] PRIMARY KEY CLUSTERED 
(
	[pcTaxZoneRate_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcTaxZoneRates]') AND name = N'pcTaxZone_ID')
CREATE NONCLUSTERED INDEX [pcTaxZone_ID] ON [dbo].[pcTaxZoneRates] 
(
	[pcTaxZone_ID] ASC
) ON [PRIMARY]
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcTaxZoneRates]') AND name = N'pcTaxZoneRate_ID')
CREATE NONCLUSTERED INDEX [pcTaxZoneRate_ID] ON [dbo].[pcTaxZoneRates] 
(
	[pcTaxZoneRate_ID] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcTaxZoneDescriptions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcTaxZoneDescriptions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcTaxZoneDescriptions](
	[pcTaxZoneDesc_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcTaxZoneDesc] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcTaxGroups]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcTaxGroups]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcTaxGroups](
	[pcTaxGroup_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcTaxZoneDesc_ID] [int] NULL,
	[pcTaxZone_ID] [int] NULL
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcTaxGroups]') AND name = N'pcTaxGroup_ID')
CREATE NONCLUSTERED INDEX [pcTaxGroup_ID] ON [dbo].[pcTaxGroups] 
(
	[pcTaxGroup_ID] ASC
) ON [PRIMARY]
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcTaxGroups]') AND name = N'pcTaxZone_ID')
CREATE NONCLUSTERED INDEX [pcTaxZone_ID] ON [dbo].[pcTaxGroups] 
(
	[pcTaxZone_ID] ASC
) ON [PRIMARY]
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcTaxGroups]') AND name = N'pcTaxZoneDesc_ID')
CREATE NONCLUSTERED INDEX [pcTaxZoneDesc_ID] ON [dbo].[pcTaxGroups] 
(
	[pcTaxZoneDesc_ID] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcTaxEptCust]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcTaxEptCust]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcTaxEptCust](
	[pcTaxEptCust_ID] [int] IDENTITY(1,1) NOT NULL,
	[idCustomer] [int] NULL,
	[pcTaxZoneRate_ID] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcTaxEpt]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcTaxEpt]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcTaxEpt](
	[pcTEpt_StateCode] [nvarchar](50) NULL,
	[pcTEpt_ProductList] [ntext] NULL,
	[pcTEpt_CategoryList] [ntext] NULL,
	[pcTEpt_EptAll] [int] NULL,
	[pcTaxZoneRate_ID] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcTaskManager]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcTaskManager]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcTaskManager](
	[pcTaskManager_id] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcTaskVersion] [nvarchar](100) NULL,
	[pcTaskNum] [int] NULL,
	[pcTaskComplete] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSuppliers]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSuppliers]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSuppliers](
	[pcSupplier_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcSupplier_Username] [nvarchar](50) NULL,
	[pcSupplier_Password] [nvarchar](100) NULL,
	[pcSupplier_FirstName] [nvarchar](100) NULL,
	[pcSupplier_LastName] [nvarchar](100) NULL,
	[pcSupplier_Company] [nvarchar](100) NULL,
	[pcSupplier_Phone] [nvarchar](50) NULL,
	[pcSupplier_Email] [nvarchar](50) NULL,
	[pcSupplier_URL] [nvarchar](250) NULL,
	[pcSupplier_FromAddress] [nvarchar](250) NULL,
	[pcSupplier_FromAddress2] [nvarchar](250) NULL,
	[pcSupplier_FromCity] [nvarchar](50) NULL,
	[pcSupplier_FromStateProvinceCode] [nvarchar](50) NULL,
	[pcSupplier_FromZip] [nvarchar](20) NULL,
	[pcSupplier_FromCountryCode] [nvarchar](50) NULL,
	[pcSupplier_BillingAddress] [nvarchar](250) NULL,
	[pcSupplier_BillingAddress2] [nvarchar](250) NULL,
	[pcSupplier_Billingcity] [nvarchar](50) NULL,
	[pcSupplier_BillingStateProvinceCode] [nvarchar](50) NULL,
	[pcSupplier_BillingZip] [nvarchar](20) NULL,
	[pcSupplier_BillingCountryCode] [nvarchar](50) NULL,
	[pcSupplier_IsDropShipper] [int] NULL,
	[pcSupplier_NoticeEmail] [nvarchar](50) NULL,
	[pcSupplier_NoticeType] [int] NULL,
	[pcSupplier_NoticeMsg] [ntext] NULL,
	[pcSupplier_NotifyManually] [int] NULL,
	[pcSupplier_CustNotifyUpdates] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcStoreVersions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcStoreVersions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcStoreVersions](
	[pcStoreVersion_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcStoreVersion_Num] [nvarchar](50) NULL,
	[pcStoreVersion_Sub] [nvarchar](50) NULL,
	[pcStoreVersion_SP] [int] NULL,
 CONSTRAINT [aaaaapcStoreVersions_PK] PRIMARY KEY CLUSTERED 
(
	[pcStoreVersion_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcStoreSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcStoreSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcStoreSettings](
	[pcStoreSettings_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcStoreSettings_CompanyName] [nvarchar](150) NULL,
	[pcStoreSettings_CompanyAddress] [nvarchar](250) NULL,
	[pcStoreSettings_CompanyZip] [nvarchar](20) NULL,
	[pcStoreSettings_CompanyCity] [nvarchar](50) NULL,
	[pcStoreSettings_CompanyState] [nvarchar](50) NULL,
	[pcStoreSettings_CompanyCountry] [nvarchar](50) NULL,
	[pcStoreSettings_CompanyLogo] [nvarchar](250) NULL,
	[pcStoreSettings_QtyLimit] [int] NOT NULL,
	[pcStoreSettings_AddLimit] [int] NOT NULL,
	[pcStoreSettings_Pre] [int] NOT NULL,
	[pcStoreSettings_CustPre] [int] NOT NULL,
	[pcStoreSettings_CatImages] [int] NOT NULL,
	[pcStoreSettings_ShowStockLmt] [int] NOT NULL,
	[pcStoreSettings_OutOfStockPurchase] [int] NOT NULL,
	[pcStoreSettings_Cursign] [nvarchar](10) NULL,
	[pcStoreSettings_DecSign] [nvarchar](4) NULL,
	[pcStoreSettings_DivSign] [nvarchar](4) NULL,
	[pcStoreSettings_DateFrmt] [nvarchar](10) NULL,
	[pcStoreSettings_MinPurchase] [int] NOT NULL,
	[pcStoreSettings_WholesaleMinPurchase] [int] NOT NULL,
	[pcStoreSettings_URLredirect] [nvarchar](250) NULL,
	[pcStoreSettings_SSL] [nvarchar](4) NULL,
	[pcStoreSettings_SSLUrl] [nvarchar](250) NULL,
	[pcStoreSettings_IntSSLPage] [nvarchar](4) NULL,
	[pcStoreSettings_PrdRow] [int] NOT NULL,
	[pcStoreSettings_PrdRowsPerPage] [int] NOT NULL,
	[pcStoreSettings_CatRow] [int] NOT NULL,
	[pcStoreSettings_CatRowsPerPage] [int] NOT NULL,
	[pcStoreSettings_BType] [nvarchar](4) NULL,
	[pcStoreSettings_StoreOff] [nvarchar](4) NULL,
	[pcStoreSettings_StoreMsg] [ntext] NULL,
	[pcStoreSettings_WL] [int] NOT NULL,
	[pcStoreSettings_TF] [int] NOT NULL,
	[pcStoreSettings_orderLevel] [nvarchar](4) NULL,
	[pcStoreSettings_DisplayStock] [int] NOT NULL,
	[pcStoreSettings_HideCategory] [int] NOT NULL,
	[pcStoreSettings_AllowNews] [int] NOT NULL,
	[pcStoreSettings_NewsCheckOut] [int] NOT NULL,
	[pcStoreSettings_NewsReg] [int] NOT NULL,
	[pcStoreSettings_NewsLabel] [nvarchar](150) NULL,
	[pcStoreSettings_PCOrd] [int] NOT NULL,
	[pcStoreSettings_HideSortPro] [int] NOT NULL,
	[pcStoreSettings_DFLabel] [nvarchar](50) NULL,
	[pcStoreSettings_DFShow] [nvarchar](4) NULL,
	[pcStoreSettings_DFReq] [nvarchar](4) NULL,
	[pcStoreSettings_TFLabel] [nvarchar](50) NULL,
	[pcStoreSettings_TFShow] [nvarchar](4) NULL,
	[pcStoreSettings_TFReq] [nvarchar](4) NULL,
	[pcStoreSettings_DTCheck] [nvarchar](4) NULL,
	[pcStoreSettings_DeliveryZip] [nvarchar](4) NULL,
	[pcStoreSettings_OrderName] [nvarchar](4) NULL,
	[pcStoreSettings_HideDiscField] [nvarchar](4) NULL,
	[pcStoreSettings_AllowSeparate] [nvarchar](4) NULL,
	[pcStoreSettings_ReferLabel] [nvarchar](250) NULL,
	[pcStoreSettings_ViewRefer] [int] NOT NULL,
	[pcStoreSettings_RefNewCheckout] [int] NOT NULL,
	[pcStoreSettings_RefNewReg] [int] NOT NULL,
	[pcStoreSettings_BrandLogo] [int] NOT NULL,
	[pcStoreSettings_BrandPro] [int] NOT NULL,
	[pcStoreSettings_RewardsActive] [int] NOT NULL,
	[pcStoreSettings_RewardsIncludeWholesale] [int] NOT NULL,
	[pcStoreSettings_RewardsPercent] [int] NOT NULL,
	[pcStoreSettings_RewardsLabel] [nvarchar](50) NULL,
	[pcStoreSettings_RewardsReferral] [int] NOT NULL,
	[pcStoreSettings_RewardsFlat] [int] NOT NULL,
	[pcStoreSettings_RewardsFlatValue] [int] NOT NULL,
	[pcStoreSettings_RewardsPerc] [int] NOT NULL,
	[pcStoreSettings_RewardsPercValue] [int] NOT NULL,
	[pcStoreSettings_XML] [nvarchar](6) NULL,
	[pcStoreSettings_QDiscountType] [int] NOT NULL,
	[pcStoreSettings_BTOdisplayType] [int] NOT NULL,
	[pcStoreSettings_BTOOutofStockPurchase] [int] NOT NULL,
	[pcStoreSettings_BTOShowImage] [int] NOT NULL,
	[pcStoreSettings_BTOQuote] [int] NOT NULL,
	[pcStoreSettings_BTOQuoteSubmit] [int] NOT NULL,
	[pcStoreSettings_BTOQuoteSubmitOnly] [int] NOT NULL,
	[pcStoreSettings_BTODetLinkType] [int] NOT NULL,
	[pcStoreSettings_BTODetTxt] [nvarchar](50) NULL,
	[pcStoreSettings_BTOPopWidth] [int] NOT NULL,
	[pcStoreSettings_BTOPopHeight] [int] NOT NULL,
	[pcStoreSettings_BTOPopImage] [int] NOT NULL,
	[pcStoreSettings_ConfigPurchaseOnly] [int] NOT NULL,
	[pcStoreSettings_ShowSKU] [int] NOT NULL,
	[pcStoreSettings_ShowSmallImg] [int] NOT NULL,
	[pcStoreSettings_Terms] [int] NOT NULL,
	[pcStoreSettings_TermsLabel] [nvarchar](255) NULL,
	[pcStoreSettings_TermsCopy] [ntext] NULL,
	[pcStoreSettings_HideRMA] [int] NOT NULL,
	[pcStoreSettings_ShowHD] [int] NOT NULL,
	[pcStoreSettings_StoreUseToolTip] [int] NOT NULL,
	[pcStoreSettings_ErrorHandler] [int] NOT NULL,
	[pcStoreSettings_AllowCheckoutWR] [int] NULL,
	[pcStoreSettings_ViewPrdStyle] [nvarchar](4) NULL,
	[pcStoreSettings_CustomerIPAlert] [nvarchar](4) NULL,
	[pcStoreSettings_CompanyPhoneNumber] [nvarchar](20) NULL,
	[pcStoreSettings_CompanyFaxNumber] [nvarchar](20) NULL,
	[pcStoreSettings_TermsShown] [int] NULL,
	[pcStoreSettings_DisableGiftRegistry] [int] NOT NULL,
	[pcStoreSettings_seoURLs] [int] NOT NULL,
	[pcStoreSettings_seoURLs404] [nvarchar](50) NULL,
	[pcStoreSettings_QuickBuy] [int] NULL,
	[pcStoreSettings_ATCEnabled] [int] NULL,
	[pcStoreSettings_GuestCheckoutOpt] [int] NULL,
	[pcStoreSettings_RestoreCart] [int] NULL,
	[pcStoreSettings_AddThisDisplay] [int] NULL,
	[pcStoreSettings_AddThisCode] [nvarchar](4000) NULL,
	[pcStoreSettings_GoogleAnalytics] [nvarchar](50) NULL,
	[pcStoreSettings_MetaTitle] [nvarchar](255) NULL,
	[pcStoreSettings_MetaDescription] [nvarchar](255) NULL,
	[pcStoreSettings_MetaKeywords] [nvarchar](255) NULL,
	[pcStoreSettings_DisableDiscountCodes] [int] NULL,
	[pcStoreSettings_PinterestDisplay] [int] NULL,
	[pcStoreSettings_PinterestCounter] [nvarchar](15) NULL,
 CONSTRAINT [aaaaapcStoreSettings_PK] PRIMARY KEY NONCLUSTERED 
(
	[pcStoreSettings_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSecurityKeys]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSecurityKeys]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSecurityKeys](
	[pcSecurityKeyID] [int] IDENTITY(1,1) NOT NULL,
	[pcSecurityKey] [nvarchar](250) NULL,
	[pcActiveKey] [int] NOT NULL,
	[pcDateUpdated] [datetime] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSearchFields_Products]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSearchFields_Products]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSearchFields_Products](
	[idSearchFieldProduct] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[idSearchData] [int] NULL,
 CONSTRAINT [Index_6126E254_43E7_4311] PRIMARY KEY CLUSTERED 
(
	[idSearchFieldProduct] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSearchFields_Mappings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSearchFields_Mappings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSearchFields_Mappings](
	[idSearchFieldMapping] [int] IDENTITY(1,1) NOT NULL,
	[idSearchField] [int] NULL,
	[pcSearchFieldsColumn] [nvarchar](250) NULL,
	[pcSearchFieldsFileID] [nvarchar](50) NULL,
 CONSTRAINT [Index_C2C0F006_B8BD_43FF] PRIMARY KEY CLUSTERED 
(
	[idSearchFieldMapping] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSearchFields_Categories]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSearchFields_Categories]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSearchFields_Categories](
	[idSearchFieldCategory] [int] IDENTITY(1,1) NOT NULL,
	[idCategory] [int] NULL,
	[idSearchData] [int] NULL,
 CONSTRAINT [Index_5661C123_6E41_49AC] PRIMARY KEY CLUSTERED 
(
	[idSearchFieldCategory] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSearchFields]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSearchFields]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSearchFields](
	[idSearchField] [int] IDENTITY(1,1) NOT NULL,
	[pcSearchFieldName] [nvarchar](250) NULL,
	[pcSearchFieldShow] [int] NULL,
	[pcSearchFieldOrder] [int] NULL,
	[pcSearchFieldCPShow] [int] NULL,
	[pcSearchFieldSearch] [int] NULL,
	[pcSearchFieldCPSearch] [int] NULL,
 CONSTRAINT [PK_pcSearchFields] PRIMARY KEY CLUSTERED 
(
	[idSearchField] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSearchData]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSearchData]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSearchData](
	[idSearchData] [int] IDENTITY(1,1) NOT NULL,
	[idSearchField] [int] NULL,
	[pcSearchDataName] [nvarchar](4000) NULL,
	[pcSearchDataOrder] [int] NULL,
 CONSTRAINT [PK_pcSearchData] PRIMARY KEY CLUSTERED 
(
	[idSearchData] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSavedPrdStats]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSavedPrdStats]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSavedPrdStats](
	[pcSPS_ID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[idProduct] [int] NULL,
	[pcSPS_SavedTimes] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSavedCartStatistics]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSavedCartStatistics]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSavedCartStatistics](
	[pcSCStatID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcSCMonth] [int] NOT NULL,
	[pcSCYear] [int] NOT NULL,
	[pcSCTotals] [int] NOT NULL,
	[pcSCAnonymous] [int] NOT NULL,
	[pcSCTopPrds] [nvarchar](800) NULL,
 CONSTRAINT [PK_pcSavedCartStatistics] PRIMARY KEY CLUSTERED 
(
	[pcSCStatID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSavedCarts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSavedCarts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSavedCarts](
	[SavedCartID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[SavedCartGUID] [nvarchar](250) NULL,
	[SavedCartDate] [datetime] NULL,
	[SavedCartName] [nvarchar](250) NULL,
	[idcustomer] [int] NOT NULL,
 CONSTRAINT [PK_pcSavedCarts] PRIMARY KEY CLUSTERED 
(
	[SavedCartID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSavedCartArray]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSavedCartArray]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSavedCartArray](
	[SCArrayID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[SavedCartID] [int] NOT NULL,
	[SCArray0] [nvarchar](10) NULL,
	[SCArray1] [nvarchar](200) NULL,
	[SCArray2] [nvarchar](10) NULL,
	[SCArray3] [nvarchar](25) NULL,
	[SCArray4] [nvarchar](400) NULL,
	[SCArray5] [nvarchar](25) NULL,
	[SCArray6] [nvarchar](25) NULL,
	[SCArray7] [nvarchar](200) NULL,
	[SCArray8] [nvarchar](10) NULL,
	[SCArray9] [nvarchar](200) NULL,
	[SCArray10] [nvarchar](10) NULL,
	[SCArray11] [nvarchar](200) NULL,
	[SCArray12] [nvarchar](10) NULL,
	[SCArray13] [nvarchar](10) NULL,
	[SCArray14] [nvarchar](25) NULL,
	[SCArray15] [nvarchar](25) NULL,
	[SCArray16] [nvarchar](200) NULL,
	[SCArray17] [nvarchar](25) NULL,
	[SCArray18] [nvarchar](10) NULL,
	[SCArray19] [nvarchar](10) NULL,
	[SCArray20] [nvarchar](10) NULL,
	[SCArray21] [nvarchar](400) NULL,
	[SCArray22] [nvarchar](10) NULL,
	[SCArray23] [nvarchar](200) NULL,
	[SCArray24] [nvarchar](200) NULL,
	[SCArray25] [nvarchar](200) NULL,
	[SCArray26] [nvarchar](200) NULL,
	[SCArray27] [nvarchar](10) NULL,
	[SCArray28] [nvarchar](25) NULL,
	[SCArray29] [nvarchar](200) NULL,
	[SCArray30] [nvarchar](200) NULL,
	[SCArray31] [nvarchar](200) NULL,
	[SCArray32] [nvarchar](200) NULL,
	[SCArray33] [nvarchar](200) NULL,
	[SCArray34] [nvarchar](200) NULL,
	[SCArray35] [nvarchar](200) NULL,
	[SCArray36] [nvarchar](200) NULL,
	[SCArray37] [nvarchar](200) NULL,
	[SCArray38] [nvarchar](200) NULL,
	[SCArray39] [nvarchar](200) NULL,
	[SCArray40] [nvarchar](200) NULL,
	[SCArray41] [nvarchar](200) NULL,
	[SCArray42] [nvarchar](200) NULL,
	[SCArray43] [nvarchar](200) NULL,
	[SCArray44] [nvarchar](200) NULL,
	[SCArray45] [nvarchar](200) NULL,

 CONSTRAINT [PK_pcSavedCartArray] PRIMARY KEY CLUSTERED 
(
	[SCArrayID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSales_Pending]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSales_Pending]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSales_Pending](
	[pcSP_ID] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[pcSales_ID] [int] NULL,
 CONSTRAINT [PK_pcSales_Pending] PRIMARY KEY CLUSTERED 
(
	[pcSP_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSales_Completed]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSales_Completed]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSales_Completed](
	[pcSC_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcSales_ID] [int] NULL,
	[pcSC_Status] [int] NULL,
	[pcSC_StartedDate] [datetime] NULL,
	[pcSC_BUStartedDate] [datetime] NULL,
	[pcSC_BUComDate] [datetime] NULL,
	[pcSC_BUTotal] [int] NULL,
	[pcSC_StoppedDate] [datetime] NULL,
	[pcSC_REStartedDate] [datetime] NULL,
	[pcSC_REComDate] [datetime] NULL,
	[pcSC_RETotal] [int] NULL,
	[pcSC_ComDate] [datetime] NULL,
	[pcSC_SaveName] [nvarchar](250) NULL,
	[pcSC_SaveDesc] [nvarchar](1000) NULL,
	[pcSC_SaveTech] [nvarchar](2000) NULL,
	[pcSC_SaveIcon] [nvarchar](250) NULL,
	[pcSC_Archived] [int] NULL,
 CONSTRAINT [PK_pcSales_Completed] PRIMARY KEY CLUSTERED 
(
	[pcSC_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSales_BackUp]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSales_BackUp]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSales_BackUp](
	[pcSB_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcSC_ID] [int] NULL,
	[pcSales_ID] [int] NULL,
	[IDProduct] [int] NULL,
	[pcSales_TargetPrice] [int] NULL,
	[pcSB_Price] [float] NULL,
 CONSTRAINT [PK_pcSales_BackUp] PRIMARY KEY CLUSTERED 
(
	[pcSB_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcSales]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcSales]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcSales](
	[pcSales_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcSales_TargetPrice] [int] NULL,
	[pcSales_Type] [int] NULL,
	[pcSales_Relative] [int] NULL,
	[pcSales_Amount] [float] NULL,
	[pcSales_Round] [int] NULL,
	[pcSales_Name] [nvarchar](250) NULL,
	[pcSales_ImgURL] [nvarchar](250) NULL,
	[pcSales_Desc] [nvarchar](1000) NULL,
	[pcSales_CreatedDate] [datetime] NULL,
	[pcSales_EditedDate] [datetime] NULL,
	[pcSales_Param1] [nvarchar](1000) NULL,
	[pcSales_Param2] [nvarchar](1000) NULL,
	[pcSales_Tech] [nvarchar](2000) NULL,
	[pcSales_Removed] [int] NULL,
 CONSTRAINT [PK_pcSales] PRIMARY KEY CLUSTERED 
(
	[pcSales_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcRevSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcRevSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcRevSettings](
	[pcRS_RatingType] [int] NULL,
	[pcRS_MainRateTxt1] [nvarchar](255) NULL,
	[pcRS_MainRateTxt2] [nvarchar](255) NULL,
	[pcRS_MainRateTxt3] [nvarchar](255) NULL,
	[pcRS_SubRateTxt1] [nvarchar](255) NULL,
	[pcRS_SubRateTxt2] [nvarchar](255) NULL,
	[pcRS_MaxRating] [int] NULL,
	[pcRS_Img1] [nvarchar](255) NULL,
	[pcRS_Img2] [nvarchar](255) NULL,
	[pcRS_Img3] [nvarchar](255) NULL,
	[pcRS_Img4] [nvarchar](255) NULL,
	[pcRS_Img5] [nvarchar](255) NULL,
	[pcRS_Active] [int] NULL,
	[pcRS_ShowRatSum] [int] NULL,
	[pcRS_RevCount] [int] NULL,
	[pcRS_NeedCheck] [int] NULL,
	[pcRS_LockPost] [int] NULL,
	[pcRS_PostCount] [int] NULL,
	[pcRS_CalMain] [int] NULL,
	[pcRS_sendReviewReminderTemplate] [nvarchar](255) NULL,
	[pcRS_RewardForReviewURL] [nvarchar](255) NULL,
	[pcRS_LastRunDate] [datetime] NULL,
	[pcRS_SendReviewReminder] [int] NULL,
	[pcRS_sendReviewReminderDays] [int] NULL,
	[pcRS_sendReviewReminderType] [int] NULL,
	[pcRS_sendReviewReminderFormat] [int] NULL,
	[pcRS_RewardForReview] [int] NULL,
	[pcRS_RewardForReviewFirstPts] [int] NULL,
	[pcRS_RewardForReviewAdditionalPts] [int] NULL,
	[pcRS_RewardForReviewMinLength] [int] NULL,
	[pcRS_RewardForReviewMaxPts] [int] NULL,
	[pcRS_DisplayRatings] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcRevLists]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcRevLists]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcRevLists](
	[pcRL_IDField] [int] NULL,
	[pcRL_Name] [nvarchar](255) NULL,
	[pcRL_Value] [nvarchar](255) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcReviewSpecials]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcReviewSpecials]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcReviewSpecials](
	[pcRS_IDProduct] [int] NULL,
	[pcRS_FieldList] [ntext] NULL,
	[pcRS_FieldOrder] [ntext] NULL,
	[pcRS_Required] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcReviewsData]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcReviewsData]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcReviewsData](
	[pcRD_IDReview] [int] NULL,
	[pcRD_IDField] [int] NULL,
	[pcRD_Feel] [int] NULL,
	[pcRD_Rate] [int] NULL,
	[pcRD_Comment] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcReviews]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcReviews]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcReviews](
	[pcRev_IDReview] [int] IDENTITY(1,1) NOT NULL,
	[pcRev_IDProduct] [int] NULL,
	[pcRev_Active] [int] NULL,
	[pcRev_IP] [nvarchar](50) NULL,
	[pcRev_Date] [datetime] NULL,
	[pcRev_MainRate] [int] NULL,
	[pcRev_MainDRate] [int] NULL,
	[pcRev_IDOrder] [int] NULL,
	[pcRev_IDCustomer] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcReviewPoints]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcReviewPoints]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcReviewPoints](
	[pcRP_ID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcRP_IDReview] [int] NULL,
	[pcRP_IDCustomer] [int] NULL,
	[pcRP_PointsAwarded] [int] NULL,
	[pcRP_DateAwarded] [datetime] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcReviewNotifications]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcReviewNotifications]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcReviewNotifications](
	[pcRN_idCustomer] [int] NULL,
	[pcRN_idOrder] [int] NULL,
	[pcRN_UniqueID] [nvarchar](36) NULL,
	[pcRN_DateSent] [datetime] NULL,
	[pcRN_DateLastViewed] [datetime] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcRevFields]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcRevFields]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcRevFields](
	[pcRF_IDField] [int] IDENTITY(1,1) NOT NULL,
	[pcRF_Name] [nvarchar](255) NULL,
	[pcRF_Type] [int] NULL,
	[pcRF_Active] [int] NULL,
	[pcRF_Required] [int] NULL,
	[pcRF_Order] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcRevExc]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcRevExc]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcRevExc](
	[pcRE_IDProduct] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcRevBadWords]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcRevBadWords]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcRevBadWords](
	[pcRBW_word] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[PCReturns]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[PCReturns]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[PCReturns](
	[idRMA] [int] IDENTITY(1,1) NOT NULL,
	[rmaNumber] [nvarchar](250) NULL,
	[rmaReturnReason] [ntext] NULL,
	[rmaDateTime] [datetime] NULL,
	[rmaReturnStatus] [ntext] NULL,
	[idOrder] [int] NULL,
	[rmaIdProducts] [ntext] NULL,
	[rmaApproved] [int] NULL,
 CONSTRAINT [aaaaaPCReturns_PK] PRIMARY KEY NONCLUSTERED 
(
	[idRMA] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcRecentRevSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcRecentRevSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcRecentRevSettings](
	[pcRR_ID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcRR_RecentRevCount] [int] NOT NULL,
	[pcRR_Style] [nvarchar](4) NULL,
	[pcRR_PageDesc] [ntext] NULL,
	[pcRR_RevDays] [int] NOT NULL,
	[pcRR_NotForSale] [int] NOT NULL,
	[pcRR_OutOfStock] [int] NOT NULL,
	[pcRR_SKU] [int] NOT NULL,
	[pcRR_ShowImg] [int] NOT NULL,
	[pcRR_ReviewsPerProduct] [int] NULL,
 CONSTRAINT [PK_pcRecentRevSettings] PRIMARY KEY CLUSTERED 
(
	[pcRR_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcProductsVATRates]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcProductsVATRates]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcProductsVATRates](
	[pcProductsVATRates_ID] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[pcVATRate_ID] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcProductsOrderedOptions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcProductsOrderedOptions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcProductsOrderedOptions](
	[ProdOrdOpt_ID] [int] IDENTITY(1,1) NOT NULL,
	[idProductOrdered] [int] NULL,
	[idoptoptgrp] [int] NULL,
 CONSTRAINT [aaaaapcProductsOrderedOptions_PK] PRIMARY KEY NONCLUSTERED 
(
	[ProdOrdOpt_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcProductsOptions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcProductsOptions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcProductsOptions](
	[pcProdOpt_ID] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[idOptionGroup] [int] NULL,
	[pcProdOpt_Required] [int] NULL,
	[pcProdOpt_Order] [int] NULL,
 CONSTRAINT [aaaaapcProductsOptions_PK] PRIMARY KEY NONCLUSTERED 
(
	[pcProdOpt_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcProductsImages]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcProductsImages]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcProductsImages](
	[pcProdImage_ID] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[pcProdImage_Url] [nvarchar](150) NULL,
	[pcProdImage_LargeUrl] [nvarchar](150) NULL,
	[pcProdImage_Order] [int] NULL,
 CONSTRAINT [aaaaapcProductsImages_PK] PRIMARY KEY NONCLUSTERED 
(
	[pcProdImage_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcProductsExc]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcProductsExc]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcProductsExc](
	[pcPE_IDProduct] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPriority]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPriority]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPriority](
	[pcPri_IDPri] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcPri_Name] [nvarchar](100) NULL,
	[pcPri_Img] [nvarchar](100) NULL,
	[pcPri_ShowImg] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPrdPromotions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPrdPromotions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPrdPromotions](
	[pcPrdPro_id] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[idproduct] [int] NOT NULL,
	[pcPrdPro_QtyTrigger] [int] NOT NULL,
	[pcPrdPro_DiscountType] [int] NOT NULL,
	[pcPrdPro_DiscountValue] [money] NULL,
	[pcPrdPro_ApplyUnits] [int] NOT NULL,
	[pcPrdPro_PromoMsg] [nvarchar](255) NULL,
	[pcPrdPro_ConfirmMsg] [nvarchar](255) NULL,
	[pcPrdPro_Sdesc] [nvarchar](255) NULL,
	[pcPrdPro_Inactive] [int] NULL,
	[pcPrdPro_IncExcCust] [int] NULL,
	[pcPrdPro_IncExcCPrice] [int] NULL,
	[pcPrdPro_RetailFlag] [int] NULL,
	[pcPrdPro_WholesaleFlag] [int] NULL,
 CONSTRAINT [PK_pcPrdPromotions] PRIMARY KEY CLUSTERED 
(
	[pcPrdPro_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPPFProducts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPPFProducts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPPFProducts](
	[pcPPFProds_id] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcPrdPro_id] [int] NOT NULL,
	[idproduct] [int] NOT NULL,
 CONSTRAINT [PK_pcPPFProducts] PRIMARY KEY CLUSTERED 
(
	[pcPPFProds_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPPFCusts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPPFCusts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPPFCusts](
	[pcPPFCusts_id] [int] IDENTITY(1,1) NOT NULL,
	[pcPrdPro_id] [int] NULL,
	[idCustomer] [int] NULL,
 CONSTRAINT [PK__pcPPFCusts__1960B67E] PRIMARY KEY CLUSTERED 
(
	[pcPPFCusts_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPPFCustPriceCats]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPPFCustPriceCats]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPPFCustPriceCats](
	[pcPPFCustPriceCats_id] [int] IDENTITY(1,1) NOT NULL,
	[pcPrdPro_id] [int] NULL,
	[idCustomerCategory] [int] NULL,
 CONSTRAINT [PK__pcPPFCustPriceCa__2D67AF2B] PRIMARY KEY CLUSTERED 
(
	[pcPPFCustPriceCats_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPPFCategories]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPPFCategories]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPPFCategories](
	[pcPPFCats_id] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcPrdPro_id] [int] NOT NULL,
	[idcategory] [int] NOT NULL,
	[pcPPFCats_IncSubCats] [int] NOT NULL,
 CONSTRAINT [PK_pcPPFCategories] PRIMARY KEY CLUSTERED 
(
	[pcPPFCats_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_USAePay_Orders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_USAePay_Orders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_USAePay_Orders](
	[idePayOrder] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[idOrder] [int] NULL,
	[Amount] [money] NULL,
	[paymentmethod] [nvarchar](250) NULL,
	[transtype] [nvarchar](250) NULL,
	[RefNum] [nvarchar](250) NULL,
	[ccCard] [nvarchar](250) NULL,
	[ccExp] [nvarchar](250) NULL,
	[idCustomer] [int] NULL,
	[fname] [nvarchar](250) NULL,
	[lname] [nvarchar](250) NULL,
	[address] [nvarchar](250) NULL,
	[zip] [nvarchar](250) NULL,
	[captured] [int] NULL,
	[pcSecurityKeyID] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_USAePay]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_USAePay]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_USAePay](
	[pcPay_Uep_Id] [int] NULL,
	[pcPay_Uep_SourceKey] [nvarchar](255) NULL,
	[pcPay_Uep_TransType] [int] NULL,
	[pcPay_Uep_TestMode] [int] NULL,
	[pcPay_Uep_Checking] [int] NULL,
	[pcPay_Uep_CheckPending] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_TripleDeal]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_TripleDeal]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_TripleDeal](
	[pcPay_TD_ID] [int] NULL,
	[pcPay_TD_MerchantName] [nvarchar](50) NULL,
	[pcPay_TD_MerchantPassword] [nvarchar](250) NULL,
	[pcPay_TD_Profile] [nvarchar](50) NULL,
	[pcPay_TD_ClientLang] [nvarchar](50) NULL,
	[pcPay_TD_PayPeriod] [int] NULL,
	[pcPay_TD_TestMode] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_SkipJack]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_SkipJack]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_SkipJack](
	[pcPay_SkipJack_ID] [int] NULL,
	[pcPay_SkipJack_SerialNumber] [nvarchar](150) NULL,
	[pcPay_SkipJack_TestMode] [int] NULL,
	[pcPay_SkipJack_Cvc2] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_SecPay]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_SecPay]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_SecPay](
	[pcPay_SecPay_ID] [int] NULL,
	[pcPay_SecPay_TransType] [nvarchar](50) NULL,
	[pcPay_SecPay_Username] [nvarchar](250) NULL,
	[pcPay_SecPay_Password] [nvarchar](250) NULL,
	[pcPay_SecPay_TestMode] [int] NULL,
	[pcPay_SecPay_Cvc2] [int] NULL,
	[pcPay_SecPay_AVS] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_PayPal_Authorize]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_PayPal_Authorize]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_PayPal_Authorize](
	[idPayPal_Authorize] [int] IDENTITY(1,1) NOT NULL,
	[idOrder] [int] NULL,
	[orderDate] [datetime] NULL,
	[orderStatus] [int] NULL,
	[gwTransId] [nvarchar](250) NULL,
	[amount] [money] NULL,
	[paymentmethod] [nvarchar](25) NULL,
	[transtype] [nvarchar](250) NULL,
	[authcode] [nvarchar](250) NULL,
	[idCustomer] [int] NULL,
	[comments] [nvarchar](250) NULL,
	[AuthorizedDate] [datetime] NULL,
	[captured] [int] NULL,
	[CurrencyCode] [nvarchar](4) NULL,
	[gwCode] [int] NULL,
 CONSTRAINT [PK__pcPay_PayPal_Aut__5D4BCC77] PRIMARY KEY CLUSTERED 
(
	[idPayPal_Authorize] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_PayPal]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_PayPal]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_PayPal](
	[pcPay_PayPal_ID] [int] NOT NULL,
	[pcPay_PayPal_TransType] [int] NOT NULL,
	[pcPay_PayPal_Username] [nvarchar](255) NULL,
	[pcPay_PayPal_Password] [nvarchar](255) NULL,
	[pcPay_PayPal_AVS] [int] NOT NULL,
	[pcPay_PayPal_CVC] [int] NOT NULL,
	[pcPay_PayPal_Sandbox] [int] NOT NULL,
	[pcPay_PayPal_CertStore] [nvarchar](255) NULL,
	[pcPay_PayPal_Signature] [nvarchar](250) NULL,
	[pcPay_PayPal_Currency] [nvarchar](5) NULL,
	[pcPay_PayPal_Vendor] [nvarchar](250) NULL,
	[pcPay_PayPal_Partner] [nvarchar](250) NULL,
	[pcPay_PayPal_Subject] [nvarchar](250) NULL,
	[pcPay_PayPal_CardTypes] [nvarchar](250) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_PaymentExpress]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_PaymentExpress]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_PaymentExpress](
	[pcPay_PaymentExpress_ID] [int] NULL,
	[pcPay_PaymentExpress_TransType] [nvarchar](50) NULL,
	[pcPay_PaymentExpress_Username] [nvarchar](250) NULL,
	[pcPay_PaymentExpress_Password] [nvarchar](250) NULL,
	[pcPay_PaymentExpress_TestMode] [int] NULL,
	[pcPay_PaymentExpress_Cvc2] [int] NULL,
	[pcPay_PaymentExpress_ReceiptEmail] [nvarchar](250) NULL,
	[pcPay_PaymentExpress_TestUsername] [nvarchar](250) NULL,
	[pcPay_PaymentExpress_AVS] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_Paymentech]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_Paymentech]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_Paymentech](
	[pcPay_PT_Id] [int] NOT NULL,
	[pcPay_PT_MerchantId] [nvarchar](255) NULL,
	[pcPay_PT_BIN] [nvarchar](255) NULL,
	[pcPay_PT_Testing] [nvarchar](255) NULL,
	[pcPay_PT_TransType] [nvarchar](50) NULL,
	[pcPay_PT_CVC] [int] NOT NULL,
	[pcPay_PT_CurrencyCode] [nvarchar](50) NULL,
	[pcPay_PT_APIType] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_ParaData]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_ParaData]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_ParaData](
	[pcPay_ParaData_ID] [int] NOT NULL,
	[pcPay_ParaData_TransType] [nvarchar](255) NULL,
	[pcPay_ParaData_Key] [nvarchar](255) NULL,
	[pcPay_ParaData_CVC] [int] NOT NULL,
	[pcPay_ParaData_AVS] [int] NOT NULL,
	[pcPay_ParaData_TestMode] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_OrdersMoneris]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_OrdersMoneris]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_OrdersMoneris](
	[pcPay_MOrder_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcPay_MOrder_OrderID] [int] NOT NULL,
	[pcPay_MOrder_TransKey] [nvarchar](250) NULL,
	[pcPay_MOrder_Result] [nvarchar](250) NULL,
	[pcPay_MOrder_responseId] [nvarchar](100) NULL,
	[pcPay_MOrder_responseCode] [nvarchar](50) NULL,
	[pcPay_MOrder_DateStamp] [nvarchar](50) NULL,
	[pcPay_MOrder_TimeStamp] [nvarchar](50) NULL,
	[pcPay_MOrder_Bankcode] [nvarchar](100) NULL,
	[pcPay_MOrder_Transname] [nvarchar](100) NULL,
	[pcPay_MOrder_cardholder] [nvarchar](250) NULL,
	[pcPay_MOrder_total] [nvarchar](50) NULL,
	[pcPay_MOrder_card] [nvarchar](50) NULL,
	[pcPay_MOrder_f4l4] [nvarchar](50) NULL,
	[pcPay_MOrder_expDate] [nvarchar](50) NULL,
	[pcPay_MOrder_message] [nvarchar](250) NULL,
	[pcPay_MOrder_ISOcode] [nvarchar](50) NULL,
	[pcPay_MOrder_TransId] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_NETOne]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_NETOne]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_NETOne](
	[pcPay_NETOne_ID] [int] NOT NULL,
	[pcPay_NETOne_MID] [nvarchar](255) NULL,
	[pcPay_NETOne_Mkey] [nvarchar](255) NULL,
	[pcPay_NETOne_Tcode] [nvarchar](255) NULL,
	[pcPay_NETOne_CVV] [int] NOT NULL,
	[pcPay_NETOne_CardTypes] [nvarchar](255) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_Moneris]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_Moneris]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_Moneris](
	[pcPay_Moneris_ID] [int] NOT NULL,
	[pcPay_Moneris_StoreId] [nvarchar](255) NULL,
	[pcPay_Moneris_Key] [nvarchar](255) NULL,
	[pcPay_Moneris_TransType] [nvarchar](255) NULL,
	[pcPay_Moneris_Lang] [nvarchar](255) NULL,
	[pcPay_Moneris_TestMode] [int] NOT NULL,
	[pcPay_Moneris_CVVEnabled] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_LinkPointAPI]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_LinkPointAPI]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_LinkPointAPI](
	[pcPay_LPAPI_ID] [int] IDENTITY(1,1) NOT NULL,
	[idOrder] [int] NULL,
	[pcPay_LPAPI_OrderDate] [datetime] NULL,
	[pcPay_LPAPI_OrderStatus] [int] NOT NULL,
	[pcPay_LPAPI_CCNum] [nvarchar](20) NULL,
	[pcPay_LPAPI_CCExpmonth] [nvarchar](2) NULL,
	[pcPay_LPAPI_CCExpyear] [nvarchar](4) NULL,
	[pcPay_LPAPI_Amount] [money] NULL,
	[pcPay_LPAPI_Paymentmethod] [nvarchar](25) NULL,
	[pcPay_LPAPI_Transtype] [nvarchar](250) NULL,
	[pcPay_LPAPI_Authcode] [nvarchar](20) NULL,
	[idCustomer] [int] NOT NULL,
	[pcPay_LPAPI_Comments] [nvarchar](250) NULL,
	[pcPay_LPAPI_AuthorizedDate] [datetime] NULL,
	[pcPay_LPAPI_Captured] [int] NOT NULL,
	[pcPay_LPAPI_RTDate] [nvarchar](50) NULL,
	[pcPay_LPAPI_Fname] [nvarchar](70) NULL,
	[pcPay_LPAPI_Lname] [nvarchar](50) NULL,
	[pcSecurityKeyID] [int] NOT NULL,
 CONSTRAINT [Index_4FDE565B_86B4_4C25] PRIMARY KEY CLUSTERED 
(
	[pcPay_LPAPI_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_HSBC]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_HSBC]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_HSBC](
	[pcPay_HSBC_ID] [int] NOT NULL,
	[pcPay_HSBC_UserId] [nvarchar](255) NULL,
	[pcPay_HSBC_Password] [nvarchar](255) NULL,
	[pcPay_HSBC_ClientId] [nvarchar](255) NULL,
	[pcPay_HSBC_TransType] [nvarchar](255) NULL,
	[pcPay_HSBC_CVV] [int] NOT NULL,
	[pcPay_HSBC_Currency] [nvarchar](255) NULL,
	[pcPay_HSBC_TestMode] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_GestPay_Response]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_GestPay_Response]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_GestPay_Response](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[CUSTOMINFO] [ntext] NULL,
	[PAY1_ALERTCODE] [int] NULL,
	[PAY1_ALERTDESCRIPTION] [nvarchar](255) NULL,
	[PAY1_AMOUNT] [money] NULL,
	[PAY1_AUTHORIZATIONCODE] [nvarchar](6) NULL,
	[PAY1_BANKTRANSACTIONID] [nvarchar](9) NULL,
	[PAY1_COUNTRY] [nvarchar](30) NULL,
	[PAY1_CHEMAIL] [nvarchar](50) NULL,
	[PAY1_CHNAME] [nvarchar](50) NULL,
	[PAY1_ERRORCODE] [int] NULL,
	[PAY1_ERRORDESCRIPTION] [nvarchar](255) NULL,
	[PAY1_OTP] [nvarchar](32) NULL,
	[PAY1_SHOPTRANSACTIONID] [nvarchar](50) NULL,
	[PAY1_TRANSACTIONRESULT] [nvarchar](2) NULL,
	[PAY1_UICCODE] [int] NULL,
	[PAY1_VBV] [nvarchar](50) NULL,
	[SHOPLOGIN] [nvarchar](30) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_GestPay_OTP]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_GestPay_OTP]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_GestPay_OTP](
	[pcPay_GestPay_OTP_id] [int] IDENTITY(1,1) NOT NULL,
	[pcPay_GestPay_OTP] [nvarchar](20) NULL,
	[pcPay_GestPay_OTP_Used] [int] NOT NULL,
	[pcPay_GestPay_OTP_Type] [char](3) NULL
) ON [PRIMARY]
END
GO
SET ANSI_PADDING OFF
GO
/****** Object:  Table [dbo].[pcPay_GestPay]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_GestPay]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_GestPay](
	[pcPay_GestPay_Id] [int] NOT NULL,
	[pcPay_GestPay_ShopLogin] [nvarchar](255) NULL,
	[pcPay_GestPay_idLanguage] [int] NOT NULL,
	[pcPay_GestPay_idCurrency] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_FastCharge]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_FastCharge]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_FastCharge](
	[pcPay_FAC_ID] [int] NOT NULL,
	[pcPay_FAC_ATSID] [nvarchar](50) NULL,
	[pcPay_FAC_TransType] [int] NOT NULL,
	[pcPay_FAC_CVV] [int] NOT NULL,
	[pcPay_FAC_Checking] [int] NOT NULL,
	[pcPay_FAC_CheckPending] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_EPN]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_EPN]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_EPN](
	[pcPay_EPN_ID] [int] NULL,
	[pcPay_EPN_Account] [nvarchar](255) NULL,
	[pcPay_EPN_CVV] [int] NULL,
	[pcPay_EPN_TestMode] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_eMerchant]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_eMerchant]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_eMerchant](
	[pcPay_eMerch_ID] [int] NOT NULL,
	[pcPay_eMerch_MerchantID] [nvarchar](250) NULL,
	[pcPay_eMerch_PaymentKey] [nvarchar](250) NULL,
	[pcPay_eMerch_CVD] [int] NOT NULL,
	[pcPay_eMerch_TransType] [nvarchar](2) NULL,
	[pcPay_eMerch_CardType] [nvarchar](4) NULL,
	[pcPay_eMerch_TestMode] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_eMerch_Orders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_eMerch_Orders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_eMerch_Orders](
	[pcPay_eMerch_Ord_ID] [int] IDENTITY(1,1) NOT NULL,
	[idOrder] [int] NOT NULL,
	[idCustomer] [int] NOT NULL,
	[pcPay_eMerch_Ord_Amount] [nvarchar](20) NULL,
	[pcPay_eMerch_Ord_CardType] [nvarchar](4) NULL,
	[pcPay_eMerch_Ord_CardNumber] [nvarchar](100) NULL,
	[pcPay_eMerch_Ord_CardExp] [nvarchar](10) NULL,
	[pcPay_eMerch_Ord_TxnNumber] [nvarchar](250) NULL,
	[pcPay_eMerch_Ord_fname] [nvarchar](250) NULL,
	[pcPay_eMerch_Ord_lname] [nvarchar](250) NULL,
	[pcPay_eMerch_Ord_streetAddr] [nvarchar](250) NULL,
	[pcPay_eMerch_Ord_Country] [nvarchar](10) NULL,
	[pcPay_eMerch_Ord_Zip] [nvarchar](20) NULL,
	[pcPay_eMerch_Ord_Captured] [int] NOT NULL,
 CONSTRAINT [Index_EC6D5FA6_3F92_4F04] PRIMARY KEY CLUSTERED 
(
	[pcPay_eMerch_Ord_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_EIG_Vault]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_EIG_Vault]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_EIG_Vault](
	[pcPay_EIG_Vault_ID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[idOrder] [int] NULL,
	[idCustomer] [int] NULL,
	[IsSaved] [int] NULL,
	[pcPay_EIG_Vault_CardNum] [nvarchar](25) NULL,
	[pcPay_EIG_Vault_CardType] [nvarchar](10) NULL,
	[pcPay_EIG_Vault_CardExp] [nvarchar](10) NULL,
	[pcPay_EIG_Vault_Token] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_EIG_Authorize]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_EIG_Authorize]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_EIG_Authorize](
	[idauthorder] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[idOrder] [int] NULL,
	[idCustomer] [int] NULL,
	[captured] [int] NULL,
	[pcSecurityKeyID] [int] NULL,
	[vaultToken] [nvarchar](50) NULL,
	[amount] [money] NULL,
	[paymentmethod] [nvarchar](250) NULL,
	[transtype] [nvarchar](250) NULL,
	[authcode] [nvarchar](250) NULL,
	[ccnum] [nvarchar](250) NULL,
	[ccexp] [nvarchar](10) NULL,
	[cctype] [nvarchar](25) NULL,
	[fname] [nvarchar](250) NULL,
	[lname] [nvarchar](250) NULL,
	[address] [nvarchar](250) NULL,
	[zip] [nvarchar](25) NULL,
	[trans_id] [nvarchar](250) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_EIG]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_EIG]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_EIG](
	[pcPay_EIG_ID] [int] NULL,
	[pcPay_EIG_Username] [nvarchar](100) NULL,
	[pcPay_EIG_Password] [nvarchar](100) NULL,
	[pcPay_EIG_Key] [nvarchar](100) NULL,
	[pcPay_EIG_Type] [nvarchar](50) NULL,
	[pcPay_EIG_Version] [nvarchar](4) NULL,
	[pcPay_EIG_Curcode] [nvarchar](4) NULL,
	[pcPay_EIG_CVV] [int] NULL,
	[pcPay_EIG_SaveCards] [int] NULL,
	[pcPay_EIG_UseVault] [int] NULL,
	[pcPay_EIG_TestMode] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_CyberSource]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON

GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_CyberSource]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_CyberSource](
	[pcPay_Cys_Id] [int] NULL,
	[pcPay_Cys_MerchantID] [nvarchar](255) NULL,
	[pcPay_Cys_TransType] [int] NULL,
	[pcPay_Cys_CardType] [nvarchar](255) NULL,
	[pcPay_Cys_CVV] [int] NULL,
	[pcPay_Cys_TestMode] [int] NULL,
	[pcPay_Cys_eCheck] [int] NULL,
	[pcPay_Cys_eCheckPending] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_Chronopay]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_Chronopay]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_Chronopay](
	[CP_Id] [int] NOT NULL,
	[CP_ProdID] [nvarchar](50) NULL,
	[CP_Currency] [nvarchar](50) NULL,
	[CP_testmode] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcPay_Chronopay]') AND name = N'CP_Id')
CREATE NONCLUSTERED INDEX [CP_Id] ON [dbo].[pcPay_Chronopay] 
(
	[CP_Id] ASC
) ON [PRIMARY]
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcPay_Chronopay]') AND name = N'CP_ProdID')
CREATE NONCLUSTERED INDEX [CP_ProdID] ON [dbo].[pcPay_Chronopay] 
(
	[CP_ProdID] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcPay_Centinel_Orders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_Centinel_Orders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_Centinel_Orders](
	[pcPay_CentOrd_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcPay_CentOrd_OrderID] [int] NULL,
	[pcPay_CentOrd_Enrolled] [nvarchar](255) NULL,
	[pcPay_CentOrd_ErrorNo] [nvarchar](255) NULL,
	[pcPay_CentOrd_ErrorDesc] [nvarchar](255) NULL,
	[pcPay_CentOrd_PAResStatus] [nvarchar](255) NULL,
	[pcPay_CentOrd_SignatureVerification] [nvarchar](255) NULL,
	[pcPay_CentOrd_EciFlag] [nvarchar](255) NULL,
	[pcPay_CentOrd_Xid] [nvarchar](255) NULL,
	[pcPay_CentOrd_Cavv] [nvarchar](255) NULL,
	[pcPay_CentOrd_rErrorNo] [nvarchar](255) NULL,
	[pcPay_CentOrd_rErrorDesc] [nvarchar](255) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_Centinel]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_Centinel]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_Centinel](
	[pcPay_Cent_ID] [int] NULL,
	[pcPay_Cent_TransactionURL] [nvarchar](255) NULL,
	[pcPay_Cent_ProcessorId] [nvarchar](255) NULL,
	[pcPay_Cent_MerchantID] [nvarchar](255) NULL,
	[pcPay_Cent_Active] [int] NULL,
	[pcPay_Cent_Password] [nvarchar](250) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_CBN]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_CBN]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_CBN](
	[pcPay_CBN_id] [int] NULL,
	[pcPay_CBN_merchant] [nvarchar](255) NULL,
	[pcPay_CBN_test] [int] NULL,
	[pcPay_CBN_status] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPay_ACHDirect]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPay_ACHDirect]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPay_ACHDirect](
	[pcPay_ACH_ID] [int] NOT NULL,
	[pcPay_ACH_MerchantID] [nvarchar](255) NULL,
	[pcPay_ACH_PWD] [nvarchar](255) NULL,
	[pcPay_ACH_TransType] [nvarchar](255) NULL,
	[pcPay_ACH_TestMode] [int] NOT NULL,
	[pcPay_ACH_CVV] [int] NOT NULL,
	[pcPay_ACH_CardTypes] [nvarchar](255) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcPackageInfo]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcPackageInfo]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcPackageInfo](
	[pcPackageInfo_ID] [int] IDENTITY(1,1) NOT NULL,
	[idOrder] [int] NULL,
	[pcPackageInfo_PackageNumber] [int] NULL,
	[pcPackageInfo_PackageWeight] [int] NULL,
	[pcPackageInfo_ShipToName] [nvarchar](100) NULL,
	[pcPackageInfo_ShipToAddress1] [nvarchar](150) NULL,
	[pcPackageInfo_ShipToAddress2] [nvarchar](100) NULL,
	[pcPackageInfo_ShipToAddress3] [nvarchar](100) NULL,
	[pcPackageInfo_ShipToResidential] [int] NULL,
	[pcPackageInfo_ShipToCity] [nvarchar](100) NULL,
	[pcPackageInfo_ShipToStateCode] [nvarchar](50) NULL,
	[pcPackageInfo_ShipToZip] [nvarchar](20) NULL,
	[pcPackageInfo_ShipToCountry] [nvarchar](100) NULL,
	[pcPackageInfo_ShipToPhone] [nvarchar](30) NULL,
	[pcPackageInfo_ShipToEmail] [nvarchar](150) NULL,
	[pcPackageInfo_PackageDescription] [nvarchar](250) NULL,
	[pcPackageInfo_ShipFromCompanyName] [nvarchar](100) NULL,
	[pcPackageInfo_ShipFromAttentionName] [nvarchar](100) NULL,
	[pcPackageInfo_ShipFromPhoneNumber] [nvarchar](30) NULL,
	[pcPackageInfo_ShipFromAddress1] [nvarchar](150) NULL,
	[pcPackageInfo_ShipFromAddress2] [nvarchar](100) NULL,
	[pcPackageInfo_ShipFromAddress3] [nvarchar](100) NULL,
	[pcPackageInfo_ShipFromCity] [nvarchar](100) NULL,
	[pcPackageInfo_ShipFromStateProvinceCode] [nvarchar](50) NULL,
	[pcPackageInfo_ShipFromPostalCode] [nvarchar](20) NULL,
	[pcPackageInfo_ShipFromCountryCode] [nvarchar](50) NULL,
	[pcPackageInfo_UPSServiceCode] [nvarchar](35) NULL,
	[pcPackageInfo_UPSPackageType] [nvarchar](20) NULL,
	[pcPackageInfo_PackageInsuredValue] [nvarchar](100) NULL,
	[pcPackageInfo_PackageLength] [nvarchar](20) NULL,
	[pcPackageInfo_PackageWidth] [nvarchar](20) NULL,
	[pcPackageInfo_PackageHeight] [nvarchar](20) NULL,
	[pcPackageInfo_AddSaturdayDelivery] [int] NULL,
	[pcPackageInfo_AddVerbalConfirmation] [int] NULL,
	[pcPackageInfo_VerbalConfirmationCN] [nvarchar](250) NULL,
	[pcPackageInfo_AddAdditionalHandling] [int] NULL,
	[pcPackageInfo_OverSizedIndicator] [int] NULL,
	[pcPackageInfo_UPSNotifyEmail1] [nvarchar](150) NULL,
	[pcPackageInfo_UPSNotifyEmail2] [nvarchar](150) NULL,
	[pcPackageInfo_UPSNotifyEmailMsg] [ntext] NULL,
	[pcPackageInfo_UPSCODAmount] [money] NULL,
	[pcPackageInfo_UPSCODFunds] [int] NULL,
	[pcPackageInfo_Status] [int] NULL,
	[pcPackageInfo_TrackingNumber] [nvarchar](150) NULL,
	[pcPackageInfo_ShipMethod] [nvarchar](150) NULL,
	[pcPackageInfo_ShippedDate] [datetime] NULL,
	[pcPackageInfo_Comments] [ntext] NULL,
	[pcPackageInfo_ShipToContactName] [nvarchar](50) NULL,
	[pcPackageInfo_FDXSPODFlag] [int] NULL,
	[pcPackageInfo_FDXCarrierCode] [nvarchar](50) NULL,
	[pcPackageInfo_UPSLabelFormat] [nvarchar](4) NULL,
	[pcPackageInfo_FDXRate] [money] NULL,
	[pcPackageInfo_MethodFlag] [int] NULL,
	[pcPackageInfo_EndiciaLabelFile] [nvarchar](250) NULL,
	[pcPackageInfo_Endicia] [int] NULL,
	[pcPackageInfo_EndiciaIsPIC] [int] NULL,
	[pcPackageInfo_EndiciaExp] [datetime] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcNewArrivalsSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcNewArrivalsSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcNewArrivalsSettings](
	[pcNAS_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcNAS_NewArrCount] [int] NULL,
	[pcNAS_Style] [nvarchar](4) NULL,
	[pcNAS_PageDesc] [ntext] NULL,
	[pcNAS_NDays] [int] NULL,
	[pcNAS_NotForSale] [int] NULL,
	[pcNAS_OutOfStock] [int] NULL,
	[pcNAS_SKU] [int] NULL,
	[pcNAS_ShowImg] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcMailUpSubs]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcMailUpSubs]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcMailUpSubs](
	[pcMailUpSubs_ID] [int] IDENTITY(1,1) NOT NULL,
	[idCustomer] [int] NULL,
	[pcMailUpLists_ID] [int] NULL,
	[pcMailUpSubs_LastSave] [datetime] NULL,
	[pcMailUpSubs_SyncNeeded] [int] NULL,
	[pcMailUpSubs_Optout] [int] NULL,
 CONSTRAINT [PK__pcMailUp__026B82D77E42ABEE] PRIMARY KEY CLUSTERED 
(
	[pcMailUpSubs_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcMailUpSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcMailUpSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcMailUpSettings](
	[pcMailUpSett_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcMailUpSett_APIUser] [nvarchar](50) NULL,
	[pcMailUpSett_APIPassword] [nvarchar](50) NULL,
	[pcMailUpSett_URL] [nvarchar](250) NULL,
	[pcMailUpSett_AutoReg] [int] NULL,
	[pcMailUpSett_RegSuccess] [int] NULL,
	[pcMailUpSett_LastCustomerID] [nvarchar](50) NULL,
	[pcMailUpSett_BulkRegister] [int] NULL,
	[pcMailUpSett_LastIDList] [nvarchar](2000) NULL,
	[pcMailUpSett_LastIDProcess] [nvarchar](2000) NULL,
	[pcMailUpSett_TurnOff] [int] NULL,
 CONSTRAINT [PK__pcMailUp__CD1982486B2FD77A] PRIMARY KEY CLUSTERED 
(
	[pcMailUpSett_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcMailUpSavedGroups]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcMailUpSavedGroups]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcMailUpSavedGroups](
	[pcMailUpSavedGroups_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcMailUpSavedGroups_Name] [nvarchar](250) NULL,
	[pcMailUpSavedGroups_Data] [ntext] NULL,
 CONSTRAINT [PK__pcMailUp__2E6089FB05E3CDB6] PRIMARY KEY CLUSTERED 
(
	[pcMailUpSavedGroups_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcMailUpLists]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcMailUpLists]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcMailUpLists](
	[pcMailUpLists_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcMailUpLists_ListID] [int] NULL,
	[pcMailUpLists_ListGuid] [nvarchar](250) NULL,
	[pcMailUpLists_ListName] [nvarchar](250) NULL,
	[pcMailUpLists_ListDesc] [ntext] NULL,
	[pcMailUpLists_Active] [int] NULL,
	[pcMailUpLists_Removed] [int] NULL,
 CONSTRAINT [PK__pcMailUp__967199A571DCD509] PRIMARY KEY CLUSTERED 
(
	[pcMailUpLists_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcMailUpGroups]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcMailUpGroups]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcMailUpGroups](
	[pcMailUpGroups_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcMailUpLists_ID] [int] NULL,
	[pcMailUpGroups_GroupID] [int] NULL,
	[pcMailUpGroups_GroupName] [nvarchar](250) NULL,
 CONSTRAINT [PK__pcMailUp__1A2AD3207889D298] PRIMARY KEY CLUSTERED 
(
	[pcMailUpGroups_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcImageDirectory]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcImageDirectory]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcImageDirectory](
	[pcImgDir_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcImgDir_Name] [nvarchar](250) NULL,
	[pcImgDir_Type] [nvarchar](50) NULL,
	[pcImgDir_Size] [int] NOT NULL,
	[pcImgDir_DateUploaded] [datetime] NULL,
	[pcImgDir_DateIndexed] [datetime] NULL,
 CONSTRAINT [Index_AF4ABAC7_F1DA_499B] PRIMARY KEY CLUSTERED 
(
	[pcImgDir_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcHomePageSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcHomePageSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcHomePageSettings](
	[pcHPS_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcHPS_FeaturedCount] [int] NOT NULL,
	[pcHPS_Style] [nvarchar](4) NULL,
	[pcHPS_PageDesc] [ntext] NULL,
	[pcHPS_First] [int] NOT NULL,
	[pcHPS_ShowSKU] [int] NOT NULL,
	[pcHPS_ShowImg] [int] NOT NULL,
	[pcHPS_SpcCount] [int] NOT NULL,
	[pcHPS_SpcOrder] [int] NOT NULL,
	[pcHPS_NewCount] [int] NOT NULL,
	[pcHPS_NewOrder] [int] NOT NULL,
	[pcHPS_BestCount] [int] NOT NULL,
	[pcHPS_BestOrder] [int] NOT NULL,
 CONSTRAINT [aaaaapcHomePageSettings_PK] PRIMARY KEY NONCLUSTERED 
(
	[pcHPS_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcGWSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcGWSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcGWSettings](
	[pcGWSet_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcGWSet_Show] [int] NULL,
	[pcGWSet_Overview] [int] NULL,
	[pcGWSet_HTML] [ntext] NULL,
	[pcGWSet_HTMLCart] [ntext] NULL,
	[pcGWSet_OverviewCart] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcGWOptions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcGWOptions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcGWOptions](
	[pcGW_IDOpt] [int] IDENTITY(1,1) NOT NULL,
	[pcGW_OptName] [nvarchar](250) NULL,
	[pcGW_OptImg] [nvarchar](150) NULL,
	[pcGW_OptPrice] [float] NULL,
	[pcGW_Removed] [int] NOT NULL,
	[pcGW_OptActive] [int] NOT NULL,
	[pcGW_OptOrder] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcGCOrdered]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcGCOrdered]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcGCOrdered](
	[pcGO_IDProduct] [int] NULL,
	[pcGO_IDOrder] [int] NULL,
	[pcGO_GcCode] [nvarchar](50) NULL,
	[pcGO_ExpDate] [datetime] NULL,
	[pcGO_Amount] [float] NULL,
	[pcGO_Status] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcGC]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcGC]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcGC](
	[pcGC_IDProduct] [int] NULL,
	[pcGC_Exp] [int] NULL,
	[pcGC_ExpDate] [datetime] NULL,
	[pcGC_ExpDays] [int] NULL,
	[pcGC_EOnly] [int] NULL,
	[pcGC_CodeGen] [int] NULL,
	[pcGC_GenFile] [nvarchar](150) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcFTypes]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcFTypes]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcFTypes](
	[pcFType_IDType] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcFType_Name] [nvarchar](100) NULL,
	[pcFType_Img] [nvarchar](100) NULL,
	[pcFType_ShowImg] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcFStatus]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcFStatus]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcFStatus](
	[pcFStat_IDStatus] [int] IDENTITY(1,1) NOT NULL,
	[pcFStat_Name] [nvarchar](100) NULL,
	[pcFStat_Img] [nvarchar](100) NULL,
	[pcFStat_BGColor] [nvarchar](10) NULL,
	[pcFStat_ShowImg] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcExportGoogle]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcExportGoogle]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcExportGoogle](
	[pcExpG_ID] [int] IDENTITY(1,1) NOT NULL,
	[idproduct] [int] NULL,
	[pcExpG_MPN] [nvarchar](50) NULL,
	[pcExpG_UPC] [nvarchar](50) NULL,
	[pcExpG_ISBN] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcExportCashback]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcExportCashback]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcExportCashback](
	[pcExpCB_ID] [int] IDENTITY(1,1) NOT NULL,
	[idproduct] [int] NULL,
	[pcExpCB_MPN] [nvarchar](250) NULL,
	[pcExpCB_UPC] [nvarchar](250) NULL,
	[pcExpCB_ISBN] [nvarchar](250) NULL,
	[pcExpCB_SHIPPING] [nvarchar](250) NULL,
	[pcExpCB_COMMISSION] [nvarchar](250) NULL,
 CONSTRAINT [Index_1F6B3502_08EE_4F13] PRIMARY KEY CLUSTERED 
(
	[pcExpCB_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcEvProducts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcEvProducts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcEvProducts](
	[pcEP_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcEP_IDEvent] [int] NULL,
	[pcEP_IDProduct] [int] NULL,
	[pcEP_Qty] [int] NULL,
	[pcEP_HQty] [int] NULL,
	[pcEP_IDOptionA] [int] NULL,
	[pcEP_IDOptionB] [int] NULL,
	[pcEP_xdetails] [ntext] NULL,
	[pcEP_GC] [int] NULL,
	[pcEP_IDConfig] [int] NULL,
	[pcEP_Price] [float] NULL,
	[pcEP_OptionsArray] [nvarchar](250) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcEvents]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcEvents]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcEvents](
	[pcEv_IDEvent] [int] IDENTITY(1,1) NOT NULL,
	[pcEv_IDCustomer] [int] NULL,
	[pcEv_Type] [nvarchar](50) NULL,
	[pcEv_Name] [nvarchar](250) NULL,
	[pcEv_Date] [datetime] NULL,
	[pcEv_Delivery] [int] NULL,
	[pcEv_MyAddr] [int] NULL,
	[pcEv_Hide] [int] NULL,
	[pcEv_Notify] [int] NULL,
	[pcEv_IncGcs] [int] NULL,
	[pcEv_Active] [int] NULL,
	[pcEv_Code] [nvarchar](50) NULL,
	[pcEv_HideAddress] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcErrorHandler]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcErrorHandler]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcErrorHandler](
	[pcErrorHandler_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcErrorHandler_SessionID] [nvarchar](250) NULL,
	[pcErrorHandler_RequestMethod] [nvarchar](250) NULL,
	[pcErrorHandler_ServerPort] [nvarchar](50) NULL,
	[pcErrorHandler_HTTPS] [nvarchar](50) NULL,
	[pcErrorHandler_LocalAddr] [nvarchar](250) NULL,
	[pcErrorHandler_RemoteAddr] [nvarchar](250) NULL,
	[pcErrorHandler_UserAgent] [nvarchar](250) NULL,
	[pcErrorHandler_URL] [nvarchar](250) NULL,
	[pcErrorHandler_HttpHost] [nvarchar](250) NULL,
	[pcErrorHandler_HttpLang] [nvarchar](50) NULL,
	[pcErrorHandler_ErrNumber] [nvarchar](50) NULL,
	[pcErrorHandler_ErrSource] [nvarchar](250) NULL,
	[pcErrorHandler_ErrDescription] [ntext] NULL,
	[pcErrorHandler_InsertDate] [datetime] NULL,
	[pcErrorHandler_customerRefID] [nvarchar](100) NULL,
 CONSTRAINT [aaaaapcErrorHandler_PK] PRIMARY KEY NONCLUSTERED 
(
	[pcErrorHandler_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcEDCTrans]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcEDCTrans]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcEDCTrans](
	[pcET_ID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[IDOrder] [int] NULL,
	[pcPackageInfo_ID] [int] NULL,
	[pcET_LabelFile] [nvarchar](250) NULL,
	[pcET_TransID] [nvarchar](100) NULL,
	[pcET_TransDate] [datetime] NULL,
	[pcET_Postage] [money] NULL,
	[pcET_RefundID] [int] NULL,
	[pcET_Method] [int] NULL,
	[pcET_Success] [int] NULL,
	[pcET_ErrMsg] [nvarchar](4000) NULL,
	[pcET_PicNum] [nvarchar](100) NULL,
	[pcET_CustomsNum] [nvarchar](100) NULL,
	[pcET_Fees] [money] NULL,
	[pcET_FeesDetails] [nvarchar](4000) NULL,
	[pcET_subPostage] [money] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcEDCSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcEDCSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcEDCSettings](
	[pcES_ID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcES_UserID] [int] NULL,
	[pcES_WebPass] [nvarchar](100) NULL,
	[pcES_PassP] [nvarchar](100) NULL,
	[pcES_AutoRefill] [int] NULL,
	[pcES_TriggerAmount] [money] NULL,
	[pcES_FillAmount] [money] NULL,
	[pcES_LogTrans] [int] NULL,
	[pcES_Reg] [int] NULL,
	[pcES_TestMode] [int] NULL,
	[pcES_AutoRmvLogs] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcEDCLogs]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcEDCLogs]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcEDCLogs](
	[pcELog_ID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcET_ID] [int] NULL,
	[pcELog_Request] [ntext] NULL,
	[pcELog_Response] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcDropShippersSuppliers]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcDropShippersSuppliers]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcDropShippersSuppliers](
	[pcDS_ID] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[pcDS_IsDropShipper] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcDropShippersOrders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcDropShippersOrders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcDropShippersOrders](
	[pcDropShipO_ID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcDropShipO_DropShipper_ID] [int] NULL,
	[pcDropShipO_idOrder] [int] NULL,
	[pcDropShipO_OrderStatus] [int] NULL,
	[pcDropShipO_Custom1] [nvarchar](250) NULL,
	[pcDropShipO_Custom2] [nvarchar](250) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcDropshippers]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcDropshippers]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcDropshippers](
	[pcDropShipper_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcDropShipper_Username] [nvarchar](50) NULL,
	[pcDropShipper_Password] [nvarchar](100) NULL,
	[pcDropShipper_FirstName] [nvarchar](100) NULL,
	[pcDropShipper_LastName] [nvarchar](100) NULL,
	[pcDropShipper_Company] [nvarchar](100) NULL,
	[pcDropShipper_Phone] [nvarchar](50) NULL,
	[pcDropShipper_Email] [nvarchar](50) NULL,
	[pcDropShipper_URL] [nvarchar](250) NULL,
	[pcDropShipper_FromAddress] [nvarchar](250) NULL,
	[pcDropShipper_FromAddress2] [nvarchar](250) NULL,
	[pcDropShipper_FromCity] [nvarchar](50) NULL,
	[pcDropShipper_FromStateProvinceCode] [nvarchar](50) NULL,
	[pcDropShipper_FromZip] [nvarchar](20) NULL,
	[pcDropShipper_FromCountryCode] [nvarchar](50) NULL,
	[pcDropShipper_BillingAddress] [nvarchar](250) NULL,
	[pcDropShipper_BillingAddress2] [nvarchar](250) NULL,
	[pcDropShipper_Billingcity] [nvarchar](50) NULL,
	[pcDropShipper_BillingStateProvinceCode] [nvarchar](50) NULL,
	[pcDropShipper_BillingZip] [nvarchar](20) NULL,
	[pcDropShipper_BillingCountryCode] [nvarchar](50) NULL,
	[pcDropShipper_NoticeEmail] [nvarchar](50) NULL,
	[pcDropShipper_NoticeType] [int] NULL,
	[pcDropShipper_NoticeMsg] [ntext] NULL,
	[pcDropShipper_NotifyManually] [int] NULL,
	[pcDropShipper_CustNotifyUpdates] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcDFShip]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcDFShip]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcDFShip](
	[pcFShip_IDDiscount] [int] NULL,
	[pcFShip_IDShipOpt] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcDFProds]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcDFProds]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcDFProds](
	[pcFPro_IDDiscount] [int] NULL,
	[pcFPro_IDProduct] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcDFCusts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcDFCusts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcDFCusts](
	[pcFCust_IDDiscount] [int] NULL,
	[pcFCust_IDCustomer] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcDFCustPriceCats]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcDFCustPriceCats]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcDFCustPriceCats](
	[pcFCPCat_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcFCPCat_IDDiscount] [int] NULL,
	[pcFCPCat_IDCategory] [int] NULL,
 CONSTRAINT [Index_8C37218D_9D5B_48D0] PRIMARY KEY CLUSTERED 
(
	[pcFCPCat_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcDFCats]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcDFCats]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcDFCats](
	[pcFCat_IDDiscount] [int] NULL,
	[pcFCat_IDCategory] [int] NULL,
	[pcFCat_SubCats] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCustomerTermsAgreed]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCustomerTermsAgreed]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCustomerTermsAgreed](
	[pcCustomerTermsAgreed_ID] [int] IDENTITY(1,1) NOT NULL,
	[idCustomer] [int] NOT NULL,
	[idOrder] [int] NOT NULL,
	[pcCustomerTermsAgreed_InsertDate] [datetime] NULL,
 CONSTRAINT [Index_FF0FB7B7_7CEF_480B] PRIMARY KEY CLUSTERED 
(
	[pcCustomerTermsAgreed_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCustomerSessions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCustomerSessions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCustomerSessions](
	[idDbSession] [int] IDENTITY(1,1) NOT NULL,
	[randomKey] [int] NULL,
	[idCustomer] [int] NULL,
	[pcCustSession_Date] [datetime] NULL,
	[pcCustSession_CustomerEmail] [nvarchar](255) NULL,
	[pcCustSession_CustomerPassword] [nvarchar](100) NULL,
	[pcCustSession_ShippingFirstName] [nvarchar](100) NULL,
	[pcCustSession_ShippingLastName] [nvarchar](100) NULL,
	[pcCustSession_ShippingCompany] [nvarchar](255) NULL,
	[pcCustSession_ShippingAddress] [nvarchar](255) NULL,
	[pcCustSession_ShippingAddress2] [nvarchar](150) NULL,
	[pcCustSession_ShippingCity] [nvarchar](255) NULL,
	[pcCustSession_ShippingStateCode] [nvarchar](4) NULL,
	[pcCustSession_ShippingProvince] [nvarchar](255) NULL,
	[pcCustSession_ShippingPostalCode] [nvarchar](20) NULL,
	[pcCustSession_ShippingCountryCode] [nvarchar](4) NULL,
	[pcCustSession_ShippingPhone] [nvarchar](30) NULL,
	[pcCustSession_ShippingFax] [nvarchar](30) NULL,
	[pcCustSession_ShippingResidential] [nvarchar](4) NULL,
	[pcCustSession_ShippingNickName] [nvarchar](50) NULL,
	[pcCustSession_ShippingReferenceId] [nvarchar](10) NULL,
	[pcCustSession_TaxShippingAlone] [nvarchar](4) NULL,
	[pcCustSession_TaxShippingAndHandlingTogether] [nvarchar](4) NULL,
	[pcCustSession_TaxCountyCode] [nvarchar](10) NULL,
	[pcCustSession_TaxLocation] [nvarchar](50) NULL,
	[pcCustSession_TaxProductAmount] [nvarchar](50) NULL,
	[pcCustSession_TF1] [nvarchar](50) NULL,
	[pcCustSession_DF1] [nvarchar](50) NULL,
	[pcCustSession_IdRefer] [int] NULL,
	[pcCustSession_UseRewards] [int] NULL,
	[pcCustSession_RewardsBalance] [int] NULL,
	[pcCustSession_IdPayment] [int] NULL,
	[pcCustSession_OrdPackageNumber] [int] NULL,
	[pcCustSession_ShippingArray] [ntext] NULL,
	[pcCustSession_Comment] [ntext] NULL,
	[pcCustSession_ShippingEmail] [nvarchar](255) NULL,
	[pcCustSession_BillingFirstName] [nvarchar](100) NULL,
	[pcCustSession_BillingLastName] [nvarchar](100) NULL,
	[pcCustSession_BillingCompany] [nvarchar](255) NULL,
	[pcCustSession_BillingAddress] [nvarchar](255) NULL,
	[pcCustSession_BillingAddress2] [nvarchar](150) NULL,
	[pcCustSession_BillingCity] [nvarchar](50) NULL,
	[pcCustSession_BillingStateCode] [nvarchar](4) NULL,
	[pcCustSession_BillingProvince] [nvarchar](50) NULL,
	[pcCustSession_BillingPostalCode] [nvarchar](20) NULL,
	[pcCustSession_BillingCountryCode] [nvarchar](4) NULL,
	[pcCustSession_BillingPhone] [nvarchar](30) NULL,
	[pcCustSession_BillingFax] [nvarchar](30) NULL,
	[pcCustSession_discountcode] [nvarchar](400) NULL,
	[pcCustSession_CartRewards] [money] NOT NULL,
	[pcCustSession_NullShipper] [nvarchar](5) NULL,
	[pcCustSession_OrderName] [nvarchar](50) NULL,
	[pcCustSession_GcReName] [nvarchar](100) NULL,
	[pcCustSession_GcReEmail] [nvarchar](75) NULL,
	[pcCustSession_GcReMsg] [nvarchar](250) NULL,
	[pcCustSession_NullShipRates] [nvarchar](5) NULL,
	[pcCustSession_taxDetailsString] [nvarchar](400) NULL,
	[pcCustSession_VATTotal] [money] NOT NULL,
	[pcCustSession_intCodeCnt] [int] NOT NULL,
	[pcCustSession_discountAmount] [nvarchar](400) NULL,
	[pcCustSession_total] [money] NOT NULL,
	[pcCustSession_taxAmount] [money] NOT NULL,
	[pcCustSession_GWTotal] [money] NOT NULL,
	[pcCustSession_ShowShipAddr] [int] NOT NULL,
	[pcCustSession_RewardsDollarValue] [money] NOT NULL,
	[pcCustSession_chkPayment] [nvarchar](5) NULL,
	[pcCustSession_pSubTotal] [money] NOT NULL,
	[pcCustSession_DiscountCodeTotal] [money] NOT NULL,
	[pcCustSession_CatDiscTotal] [money] NOT NULL,
	[pcCustSession_strBundleArray] [nvarchar](400) NULL,
	[pcCustSession_GCDetails] [nvarchar](400) NULL,
	[pcCustSession_GCTotal] [money] NOT NULL,
	[pcCustSession_SB_taxAmount] [float] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCustomerFieldsValues]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCustomerFieldsValues]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCustomerFieldsValues](
	[pcCFV_ID] [int] IDENTITY(1,1) NOT NULL,
	[idCustomer] [int] NOT NULL,
	[pcCField_ID] [int] NOT NULL,
	[pcCFV_Value] [ntext] NULL,
 CONSTRAINT [aaaaapcCustomerFieldsValues_PK] PRIMARY KEY NONCLUSTERED 
(
	[pcCFV_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCustomerFields]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCustomerFields]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCustomerFields](
	[pcCField_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcCField_Name] [nvarchar](150) NULL,
	[pcCField_Description] [ntext] NULL,
	[pcCField_FieldType] [int] NOT NULL,
	[pcCField_Value] [nvarchar](250) NULL,
	[pcCField_Length] [int] NOT NULL,
	[pcCField_Maximum] [int] NOT NULL,
	[pcCField_Required] [int] NOT NULL,
	[pcCField_ShowOnReg] [int] NOT NULL,
	[pcCField_ShowOnCheckout] [int] NOT NULL,
	[pcCField_PricingCategories] [int] NOT NULL,
	[pcCField_Order] [int] NULL,
 CONSTRAINT [aaaaapcCustomerFields_PK] PRIMARY KEY NONCLUSTERED 
(
	[pcCField_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCustomerCategories]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCustomerCategories]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCustomerCategories](
	[idCustomerCategory] [int] IDENTITY(1,1) NOT NULL,
	[pcCC_Name] [nvarchar](250) NULL,
	[pcCC_Description] [nvarchar](250) NULL,
	[pcCC_WholesalePriv] [int] NULL,
	[pcCC_CategoryType] [nvarchar](10) NULL,
	[pcCC_ATB_Percentage] [int] NULL,
	[pcCC_ATB_Off] [nvarchar](10) NULL,
	[pcCC_NFSoverride] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCustFieldsPricingCats]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCustFieldsPricingCats]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCustFieldsPricingCats](
	[pcCFPC_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcCField_ID] [int] NOT NULL,
	[idCustomerCategory] [int] NOT NULL,
 CONSTRAINT [aaaaapcCustFieldsPricingCats_PK] PRIMARY KEY NONCLUSTERED 
(
	[pcCFPC_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCPFProducts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCPFProducts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCPFProducts](
	[pcCPFProds_id] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcCatPro_id] [int] NOT NULL,
	[idproduct] [int] NOT NULL,
 CONSTRAINT [PK_pcCPFProducts] PRIMARY KEY CLUSTERED 
(
	[pcCPFProds_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCPFCategories]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCPFCategories]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCPFCategories](
	[pcCPFCats_id] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcCatPro_id] [int] NOT NULL,
	[idcategory] [int] NOT NULL,
	[pcCPFCats_IncSubCats] [int] NOT NULL,
 CONSTRAINT [PK_pcCPFCategories] PRIMARY KEY CLUSTERED 
(
	[pcCPFCats_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcContents]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcContents]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcContents](
	[pcCont_IDPage] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcCont_PageName] [nvarchar](255) NULL,
	[pcCont_Description] [ntext] NULL,
	[pcCont_IncHeader] [int] NULL,
	[pcCont_InActive] [int] NULL,
	[pcCont_MetaTitle] [nvarchar](250) NULL,
	[pcCont_MetaDesc] [nvarchar](500) NULL,
	[pcCont_MetaKeywords] [nvarchar](500) NULL,
	[pcCont_Order] [int] NOT NULL,
	[pcCont_Parent] [int] NOT NULL,
	[pcCont_Published] [int] NOT NULL,
	[pcCont_Thumbnail] [nvarchar](255) NULL,
	[pcCont_PageTitle] [nvarchar](255) NULL,
	[pcCont_Comments] [nvarchar](4000) NULL,
	[pcCont_MenuExclude] [int] NOT NULL,
	[pcCont_CustomerType] [nvarchar](50) NULL,
	[pcCont_Draft] [ntext] NULL,
	[pcCont_HideBackButton] [int] NULL,
	[pcCont_DraftStatus] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcComments]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcComments]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcComments](
	[pcComm_IdFeedback] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcComm_IDOrder] [int] NULL,
	[pcComm_IDParent] [int] NULL,
	[pcComm_IDUser] [int] NULL,
	[pcComm_CreatedDate] [datetime] NULL,
	[pcComm_EditedDate] [datetime] NULL,
	[pcComm_FType] [int] NULL,
	[pcComm_FStatus] [int] NULL,
	[pcComm_Priority] [int] NULL,
	[pcComm_Description] [ntext] NULL,
	[pcComm_Details] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCC_Pricing]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCC_Pricing]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCC_Pricing](
	[idCC_Price] [int] IDENTITY(1,1) NOT NULL,
	[idcustomerCategory] [int] NOT NULL,
	[idProduct] [int] NOT NULL,
	[pcCC_Price] [money] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCC_BTO_Pricing]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCC_BTO_Pricing]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCC_BTO_Pricing](
	[idCC_BTO_Price] [int] IDENTITY(1,1) NOT NULL,
	[idcustomerCategory] [int] NOT NULL,
	[idBTOProduct] [int] NOT NULL,
	[idBTOItem] [int] NULL,
	[pcCC_BTO_Price] [money] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCatPromotions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCatPromotions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCatPromotions](
	[pcCatPro_id] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[idcategory] [int] NOT NULL,
	[pcCatPro_QtyTrigger] [int] NOT NULL,
	[pcCatPro_DiscountType] [int] NOT NULL,
	[pcCatPro_DiscountValue] [money] NOT NULL,
	[pcCatPro_ApplyUnits] [int] NOT NULL,
	[pcCatPro_PromoMsg] [nvarchar](255) NULL,
	[pcCatPro_ConfirmMsg] [nvarchar](255) NULL,
	[pcCatPro_SDesc] [nvarchar](255) NULL,
 CONSTRAINT [PK_pcCatPromotions] PRIMARY KEY CLUSTERED 
(
	[pcCatPro_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCatDiscounts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCatDiscounts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCatDiscounts](
	[pcCD_IDDiscount] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcCD_IDCategory] [int] NOT NULL,
	[pcCD_quantityFrom] [int] NULL,
	[pcCD_quantityUntil] [int] NULL,
	[pcCD_discountPerUnit] [float] NULL,
	[pcCD_num] [int] NULL,
	[pcCD_percentage] [int] NULL,
	[pcCD_discountPerWUnit] [float] NULL,
	[pcCD_baseproductonly] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcCartArray]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcCartArray]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcCartArray](
	[pcCartArray_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcCartArray_Key] [int] NULL,
	[pcCartArray_Date] [datetime] NULL,
	[pcCartArray_0] [int] NULL,
	[pcCartArray_1] [nvarchar](200) NULL,
	[pcCartArray_2] [int] NULL,
	[pcCartArray_3] [money] NULL,
	[pcCartArray_4] [nvarchar](450) NULL,
	[pcCartArray_5] [money] NULL,
	[pcCartArray_6] [float] NOT NULL,
	[pcCartArray_7] [nvarchar](200) NULL,
	[pcCartArray_8] [int] NULL,
	[pcCartArray_9] [nvarchar](200) NULL,
	[pcCartArray_10] [int] NULL,
	[pcCartArray_11] [nvarchar](200) NULL,
	[pcCartArray_12] [int] NULL,
	[pcCartArray_13] [int] NULL,
	[pcCartArray_14] [money] NULL,
	[pcCartArray_15] [money] NULL,
	[pcCartArray_16] [nvarchar](200) NULL,
	[pcCartArray_17] [money] NULL,
	[pcCartArray_18] [int] NULL,
	[pcCartArray_19] [int] NULL,
	[pcCartArray_20] [int] NULL,
	[pcCartArray_21] [nvarchar](450) NULL,
	[pcCartArray_22] [int] NULL,
	[pcCartArray_23] [nvarchar](200) NULL,
	[pcCartArray_24] [nvarchar](200) NULL,
	[pcCartArray_25] [nvarchar](200) NULL,
	[pcCartArray_26] [nvarchar](200) NULL,
	[pcCartArray_27] [int] NULL,
	[pcCartArray_28] [money] NULL,
	[pcCartArray_29] [nvarchar](200) NULL,
	[pcCartArray_30] [nvarchar](200) NULL,
	[pcCartArray_31] [nvarchar](200) NULL,
	[pcCartArray_32] [nvarchar](200) NULL,
	[pcCartArray_33] [nvarchar](200) NULL,
	[pcCartArray_34] [nvarchar](200) NULL,
	[pcCartArray_35] [nvarchar](200) NULL,
	[pcCartArray_36] [nvarchar](250) NULL,
	[pcCartArray_37] [nvarchar](250) NULL,
	[pcCartArray_38] [nvarchar](250) NULL,
	[pcCartArray_39] [nvarchar](250) NULL,
	[pcCartArray_40] [nvarchar](250) NULL,
	[pcCartArray_41] [nvarchar](250) NULL,
	[pcCartArray_42] [nvarchar](250) NULL,
	[pcCartArray_43] [nvarchar](250) NULL,
	[pcCartArray_44] [nvarchar](250) NULL,
	[pcCartArray_45] [nvarchar](250) NULL,
 CONSTRAINT [Index_2788F14D_764A_4FE4] PRIMARY KEY CLUSTERED 
(
	[pcCartArray_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcCartArray]') AND name = N'Index_3A9EA7DA_E1B1_48D8')
CREATE UNIQUE NONCLUSTERED INDEX [Index_3A9EA7DA_E1B1_48D8] ON [dbo].[pcCartArray] 
(
	[pcCartArray_ID] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcBTODefaultPriceCats]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcBTODefaultPriceCats]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcBTODefaultPriceCats](
	[pcBDPC_id] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NOT NULL,
	[idCustomerCategory] [int] NOT NULL,
	[pcBDPC_Price] [money] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcBestSellerSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcBestSellerSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcBestSellerSettings](
	[pcBSS_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcBSS_BestSellCount] [int] NULL,
	[pcBSS_Style] [nvarchar](4) NULL,
	[pcBSS_PageDesc] [ntext] NULL,
	[pcBSS_NSold] [int] NULL,
	[pcBSS_NotForSale] [int] NULL,
	[pcBSS_OutOfStock] [int] NULL,
	[pcBSS_SKU] [int] NULL,
	[pcBSS_ShowImg] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcAmazonSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcAmazonSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcAmazonSettings](
	[pcAmzSet_id] [int] IDENTITY(1,1) NOT NULL,
	[pcAmzSet_prdIDType] [int] NULL,
	[pcAmzSet_icondition] [int] NULL,
	[pcAmzSet_price] [int] NULL,
	[pcAmzSet_willshipout] [int] NULL,
	[pcAmzSet_expship] [nvarchar](10) NULL,
	[pcAmzSet_marketplace] [nvarchar](10) NULL,
 CONSTRAINT [Index_3D252FD2_1B8B_4D26] PRIMARY KEY CLUSTERED 
(
	[pcAmzSet_id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[pcAmazonSettings]') AND name = N'Index_74E91CEF_EDDC_481F')
CREATE UNIQUE NONCLUSTERED INDEX [Index_74E91CEF_EDDC_481F] ON [dbo].[pcAmazonSettings] 
(
	[pcAmzSet_id] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[pcAmazon]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcAmazon]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcAmazon](
	[pcAmz_id] [int] IDENTITY(1,1) NOT NULL,
	[idproduct] [int] NULL,
	[pcAmz_productID] [nvarchar](50) NULL,
	[pcAmz_prdIDType] [int] NULL,
	[pcAmz_icondition] [int] NULL,
	[pcAmz_price] [int] NULL,
	[pcAmz_sku] [nvarchar](50) NULL,
	[pcAmz_quantity] [int] NULL,
	[pcAmz_willshipout] [int] NULL,
	[pcAmz_expship] [nvarchar](10) NULL,
	[pcAmz_marketplace] [nvarchar](10) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcAffiliatesPayments]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcAffiliatesPayments]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcAffiliatesPayments](
	[pcAffpay_idpayment] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[pcAffpay_idAffiliate] [int] NULL,
	[pcAffpay_Amount] [float] NULL,
	[pcAffpay_PayDate] [datetime] NULL,
	[pcAffpay_Status] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcAdminComments]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcAdminComments]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[pcAdminComments](
	[pcACOM_ID] [int] IDENTITY(1,1) NOT NULL,
	[idOrder] [int] NULL,
	[pcACOM_ComType] [int] NULL,
	[pcACOM_Comments] [ntext] NULL,
	[pcDropShipper_ID] [int] NULL,
	[pcACom_IsSupplier] [int] NULL,
	[pcPackageInfo_ID] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[pcAdminAuditLog]    Script Date: 10/19/2011 08:37:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[pcAdminAuditLog]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE  [dbo].[pcAdminAuditLog](
	[pcAdminAuditLogID] [int] IDENTITY(1,1) NOT NULL,
	[idAdmin] [int] NOT NULL,
	[idOrder] [int] NOT NULL,
	[pcAdminAuditDate] [datetime] NOT NULL,
	[pcAdminAuditPage] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[payTypes]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[payTypes]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[payTypes](
	[idPayment] [int] IDENTITY(1,1) NOT NULL,
	[gwCode] [int] NULL,
	[paymentDesc] [nvarchar](70) NULL,
	[priceToAdd] [float] NULL,
	[percentageToAdd] [float] NULL,
	[ssl] [int] NULL,
	[sslUrl] [nvarchar](150) NULL,
	[emailText] [ntext] NULL,
	[quantityFrom] [int] NULL,
	[quantityUntil] [int] NULL,
	[weightFrom] [int] NULL,
	[weightUntil] [int] NULL,
	[priceFrom] [float] NULL,
	[priceUntil] [float] NULL,
	[active] [int] NOT NULL,
	[Cbtob] [int] NULL,
	[CReq] [int] NULL,
	[Cprompt] [nvarchar](50) NULL,
	[Type] [nvarchar](50) NULL,
	[terms] [ntext] NULL,
	[cvv] [int] NULL,
	[paymentPriority] [int] NULL,
	[paymentNickName] [nvarchar](250) NULL,
	[pcPayTypes_processOrder] [int] NULL,
	[pcPayTypes_setPayStatus] [int] NULL,
	[pcPayTypes_ppab] [int] NOT NULL,
	[pcPayTypes_Subscription] [int] NOT NULL,
 CONSTRAINT [aaaaapayTypes_PK] PRIMARY KEY NONCLUSTERED 
(
	[idPayment] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[paypal]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[paypal]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[paypal](
	[id] [int] NULL,
	[Pay_To] [nvarchar](150) NULL,
	[URL] [nvarchar](250) NULL,
	[PP_Currency] [nvarchar](50) NULL,
	[PP_Sandbox] [int] NULL,
	[PP_Country] [nvarchar] (4)  NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[orders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[orders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[orders](
	[idOrder] [int] IDENTITY(1,1) NOT NULL,
	[orderDate] [datetime] NULL,
	[idCustomer] [int] NULL,
	[details] [ntext] NULL,
	[total] [float] NULL,
	[address] [nvarchar](150) NULL,
	[zip] [nvarchar](10) NULL,
	[stateCode] [nvarchar](4) NULL,
	[state] [nvarchar](50) NULL,
	[city] [nvarchar](50) NULL,
	[countryCode] [nvarchar](4) NULL,
	[comments] [ntext] NULL,
	[taxAmount] [float] NULL,
	[shipmentDetails] [nvarchar](200) NULL,
	[paymentDetails] [nvarchar](200) NULL,
	[discountDetails] [nvarchar](200) NULL,
	[randomNumber] [int] NULL,
	[shippingAddress] [nvarchar](150) NULL,
	[shippingStateCode] [nvarchar](4) NULL,
	[shippingState] [nvarchar](50) NULL,
	[shippingCity] [nvarchar](50) NULL,
	[shippingCountryCode] [nvarchar](4) NULL,
	[shippingZip] [nvarchar](10) NULL,
	[orderStatus] [int] NULL,
	[viewed] [int] NULL,
	[idAffiliate] [int] NULL,
	[processDate] [datetime] NULL,
	[shipDate] [datetime] NULL,
	[shipVia] [nvarchar](155) NULL,
	[trackingNum] [nvarchar](50) NULL,
	[affiliatePay] [float] NULL,
	[returnDate] [datetime] NULL,
	[returnReason] [nvarchar](150) NULL,
	[iRewardPoints] [int] NOT NULL,
	[ShippingFullName] [nvarchar](50) NULL,
	[iRewardValue] [float] NOT NULL,
	[iRewardRefId] [int] NOT NULL,
	[iRewardPointsRef] [int] NOT NULL,
	[iRewardPointsCustAccrued] [int] NOT NULL,
	[IDRefer] [int] NULL,
	[address2] [nvarchar](150) NULL,
	[shippingCompany] [nvarchar](150) NULL,
	[shippingAddress2] [nvarchar](150) NULL,
	[taxDetails] [ntext] NULL,
	[adminComments] [ntext] NULL,
	[rmaCredit] [money] NULL,
	[DPs] [int] NULL,
	[gwAuthCode] [nvarchar](100) NULL,
	[gwTransID] [nvarchar](70) NULL,
	[paymentCode] [nvarchar](100) NULL,
	[SRF] [int] NULL,
	[ordShiptype] [int] NULL,
	[ordPackageNum] [int] NULL,
	[ord_DeliveryDate] [datetime] NULL,
	[ord_OrderName] [nvarchar](100) NULL,
	[ord_VAT] [float] NULL,
	[pcOrd_CatDiscounts] [float] NULL,
	[pcOrd_DiscountsUsed] [nvarchar](250) NULL,
	[pcOrd_Payer] [nvarchar](150) NULL,
	[pcOrd_PaymentStatus] [int] NULL,
	[pcOrd_CustAllowSeparate] [int] NULL,
	[pcOrd_CustRequestStr] [nvarchar](150) NULL,
	[pcPay_PayPal_Signature] [nvarchar](250) NULL,
	[pcOrd_GcCode] [nvarchar](50) NULL,
	[pcOrd_GcUsed] [float] NULL,
	[pcOrd_GCs] [int] NULL,
	[pcOrd_IDEvent] [int] NULL,
	[pcOrd_GWTotal] [float] NULL,
	[pcOrd_GcReName] [nvarchar](150) NULL,
	[pcOrd_GcReEmail] [nvarchar](100) NULL,
	[pcOrd_GcReMsg] [ntext] NULL,
	[pcOrd_ShippingEmail] [nvarchar](70) NULL,
	[pcOrd_ShippingPhone] [nvarchar](30) NULL,
	[pcOrd_ShippingFax] [nvarchar](20) NULL,
	[pcOrd_CustomerIP] [nvarchar](20) NULL,
	[pcOrd_Time] [datetime] NULL,
	[pcOrd_ShipWeight] [int] NULL,
	[pcOrd_GoogleIDOrder] [nvarchar](150) NULL,
	[pcOrd_BuyerAccountAge] [nvarchar](6) NULL,
	[pcOrd_AVSRespond] [nvarchar](10) NULL,
	[pcOrd_PartialCCNumber] [nvarchar](4) NULL,
	[pcOrd_EligibleForProtection] [nvarchar](5) NULL,
	[pcOrd_CVNResponse] [nvarchar](10) NULL,
	[gwTransParentId] [nvarchar](50) NULL,
	[pcOrd_Archived] [int] NOT NULL,
	[pcOrd_OrderKey] [nvarchar](50) NULL,
	[pcOrd_ShowShipAddr] [int] NOT NULL,
	[pcOrd_GCDetails] [nvarchar](800) NULL,
	[pcOrd_GCAmount] [money] NULL,
	[pcOrd_SubTax] [float] NOT NULL,
	[pcOrd_SubTrialTax] [float] NOT NULL,
	[pcOrd_SubShipping] [float] NOT NULL,
	[pcOrd_SubTrialShipping] [float] NOT NULL,
	[pcOrd_MobileSF] [int] NULL,
 CONSTRAINT [aaaaaorders_PK] PRIMARY KEY NONCLUSTERED 
(
	[idOrder] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[optionsGroups]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[optionsGroups]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[optionsGroups](
	[idOptionGroup] [int] IDENTITY(1,1) NOT NULL,
	[OptionGroupDesc] [nvarchar](250) NULL,
 CONSTRAINT [aaaaaoptionsGroups_PK] PRIMARY KEY NONCLUSTERED 
(
	[idOptionGroup] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[options_optionsGroups]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[options_optionsGroups]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[options_optionsGroups](
	[idoptoptgrp] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[idOptionGroup] [int] NULL,
	[idOption] [int] NULL,
	[price] [float] NULL,
	[Wprice] [float] NULL,
	[sortOrder] [int] NULL,
	[InActive] [int] NULL,
 CONSTRAINT [aaaaaoptions_optionsGroups_PK] PRIMARY KEY NONCLUSTERED 
(
	[idoptoptgrp] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[options]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[options]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[options](
	[idOption] [int] IDENTITY(1,1) NOT NULL,
	[optionDescrip] [nvarchar](250) NULL,
 CONSTRAINT [aaaaaoptions_PK] PRIMARY KEY NONCLUSTERED 
(
	[idOption] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[optGrps]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[optGrps]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[optGrps](
	[idOptionGroup] [int] NULL,
	[idoption] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[offlinepayments]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[offlinepayments]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[offlinepayments](
	[idOrder] [int] NULL,
	[idpayment] [int] NULL,
	[AccNum] [nvarchar](100) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[News]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[News]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[News](
	[idnews] [int] IDENTITY(1,1) NOT NULL,
	[fromdate] [datetime] NULL,
	[fromemail] [nvarchar](100) NULL,
	[fromname] [nvarchar](100) NULL,
	[title] [nvarchar](255) NULL,
	[msgbody] [ntext] NULL,
	[msgtype] [int] NULL,
	[custfile] [nvarchar](255) NULL,
	[custtotal] [int] NULL,
 CONSTRAINT [aaaaaNews_PK] PRIMARY KEY NONCLUSTERED 
(
	[idnews] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[netbillorders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[netbillorders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[netbillorders](
	[idnetbillorder] [int] IDENTITY(1,1) NOT NULL,
	[idOrder] [int] NULL,
	[amount] [money] NULL,
	[paymentmethod] [nvarchar](250) NULL,
	[transtype] [nvarchar](250) NULL,
	[authcode] [nvarchar](250) NULL,
	[ccnum] [nvarchar](250) NULL,
	[ccexp] [nvarchar](250) NULL,
	[idCustomer] [int] NULL,
	[fname] [nvarchar](250) NULL,
	[lname] [nvarchar](250) NULL,
	[address] [nvarchar](250) NULL,
	[zip] [nvarchar](250) NULL,
	[trans_id] [nvarchar](50) NULL,
	[captured] [int] NULL,
	[pcSecurityKeyID] [int] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[netbill]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[netbill]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[netbill](
	[idNetbill] [int] NULL,
	[NBAccountID] [nvarchar](255) NULL,
	[NBCVVEnabled] [int] NULL,
	[NBAVS] [int] NULL,
	[NBTranType] [nvarchar](50) NULL,
	[NetbillCheck] [int] NULL,
	[NBSiteTag] [nvarchar](255) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[moneris]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[moneris]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[moneris](
	[moneris_id] [int] NULL,
	[store_id] [nvarchar](255) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[linkpoint]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[linkpoint]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[linkpoint](
	[id] [int] NULL,
	[storeName] [nvarchar](50) NULL,
	[lp_testmode] [nvarchar](50) NULL,
	[lp_cards] [nvarchar](50) NULL,
	[transType] [nvarchar](50) NULL,
	[CVM] [int] NULL,
	[lp_yourpay] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[layout]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[layout]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[layout](
	[ID] [int] NOT NULL,
	[headerid] [nvarchar](4) NULL,
	[recalculate] [nvarchar](100) NULL,
	[continueshop] [nvarchar](100) NULL,
	[checkout] [nvarchar](100) NULL,
	[submit] [nvarchar](100) NULL,
	[morebtn] [nvarchar](100) NULL,
	[viewcartbtn] [nvarchar](100) NULL,
	[checkoutbtn] [nvarchar](100) NULL,
	[addtocart] [nvarchar](100) NULL,
	[addtowl] [nvarchar](100) NULL,
	[register] [nvarchar](100) NULL,
	[cancel] [nvarchar](100) NULL,
	[remove] [nvarchar](100) NULL,
	[add2] [nvarchar](100) NULL,
	[login] [nvarchar](100) NULL,
	[login_checkout] [nvarchar](100) NULL,
	[back] [nvarchar](100) NULL,
	[register_checkout] [nvarchar](100) NULL,
	[customize] [nvarchar](100) NULL,
	[reconfigure] [nvarchar](100) NULL,
	[resetdefault] [nvarchar](100) NULL,
	[savequote] [nvarchar](100) NULL,
	[RevOrder] [nvarchar](100) NULL,
	[SubmitQuote] [nvarchar](100) NULL,
	[pcLO_requestQuote] [nvarchar](100) NULL,
	[pcLO_placeOrder] [nvarchar](100) NULL,
	[pcLO_checkoutWR] [nvarchar](100) NULL,
	[pcLO_processShip] [nvarchar](100) NULL,
	[pcLO_finalShip] [nvarchar](100) NULL,
	[pcLO_backtoOrder] [nvarchar](100) NULL,
	[pcLO_previous] [nvarchar](100) NULL,
	[pcLO_next] [nvarchar](100) NULL,
	[CreRegistry] [nvarchar](100) NULL,
	[DelRegistry] [nvarchar](100) NULL,
	[AddToRegistry] [nvarchar](100) NULL,
	[UpdRegistry] [nvarchar](100) NULL,
	[SendMsgs] [nvarchar](100) NULL,
	[RetRegistry] [nvarchar](100) NULL,
	[pcLO_Update] [nvarchar](155) NULL,
	[pcLO_Savecart] [nvarchar](155) NULL,
 CONSTRAINT [aaaaalayout_PK] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[klix]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[klix]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[klix](
	[idKlix] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[ssl_merchant_id] [nvarchar](255) NULL,
	[ssl_pin] [nvarchar](255) NULL,
	[CVV] [int] NULL,
	[ssl_avs] [int] NULL,
	[testmode] [int] NULL,
	[ssl_user_id] [nvarchar](255) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[ITransact]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[ITransact]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[ITransact](
	[Gateway_ID] [nvarchar](50) NULL,
	[URL] [nvarchar](255) NULL,
	[id] [int] NULL,
	[it_amex] [int] NULL,
	[it_diner] [int] NULL,
	[it_disc] [int] NULL,
	[it_mc] [int] NULL,
	[it_visa] [int] NULL,
	[ReqCVV] [int] NULL,
	[TransType] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[InternetSecure]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[InternetSecure]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[InternetSecure](
	[IsID] [int] NULL,
	[IsMerchantNumber] [nvarchar](250) NULL,
	[IsLanguage] [nvarchar](50) NULL,
	[IsCurrency] [nvarchar](50) NULL,
	[IsTestmode] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[icons]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[icons]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[icons](
	[id] [int] NULL,
	[erroricon] [nvarchar](155) NOT NULL,
	[requiredicon] [nvarchar](155) NOT NULL,
	[errorfieldicon] [nvarchar](155) NULL,
	[previousicon] [nvarchar](155) NULL,
	[nexticon] [nvarchar](155) NULL,
	[zoom] [nvarchar](155) NULL,
	[discount] [nvarchar](155) NULL,
	[arrowUp] [nvarchar](155) NULL,
	[arrowDown] [nvarchar](155) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[FlatShipTypes]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[FlatShipTypes]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[FlatShipTypes](
	[idFlatShipType] [int] IDENTITY(1,1) NOT NULL,
	[FlatShipTypeDesc] [nvarchar](150) NULL,
	[WQP] [nvarchar](10) NULL,
	[FlatShipTypeDelivery] [nvarchar](255) NULL,
	[startIncrement] [money] NULL,
 CONSTRAINT [aaaaaFlatShipTypes_PK] PRIMARY KEY NONCLUSTERED 
(
	[idFlatShipType] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[FlatShipTypeRules]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[FlatShipTypeRules]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[FlatShipTypeRules](
	[idFlatShipTypeRule] [int] IDENTITY(1,1) NOT NULL,
	[idFlatshipType] [int] NULL,
	[quantityFrom] [float] NULL,
	[quantityTo] [float] NULL,
	[shippingPrice] [float] NULL,
	[num] [int] NULL,
 CONSTRAINT [aaaaaFlatShipTypeRules_PK] PRIMARY KEY NONCLUSTERED 
(
	[idFlatShipTypeRule] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[FedExWSAPI]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[FedExWSAPI]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[FedExWSAPI](
	[FedExAPI_ID] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[FedExAPI_PersonName] [nvarchar](250) NULL,
	[FedExAPI_CompanyName] [nvarchar](100) NULL,
	[FedExAPI_Department] [nvarchar](50) NULL,
	[FedExAPI_PhoneNumber] [nvarchar](20) NULL,
	[FedExAPI_FaxNumber] [nvarchar](20) NULL,
	[FedExAPI_EmailAddress] [nvarchar](250) NULL,
	[FedExAPI_Line1] [nvarchar](250) NULL,
	[FedExAPI_Line2] [nvarchar](100) NULL,
	[FedExAPI_City] [nvarchar](100) NULL,
	[FedExAPI_State] [nvarchar](50) NULL,
	[FedExAPI_PostalCode] [nvarchar](50) NULL,
	[FedExAPI_Country] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[FedExAPI]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[FedExAPI]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[FedExAPI](
	[FedExAPI_ID] [int] IDENTITY(1,1) NOT NULL,
	[FedExAPI_PersonName] [nvarchar](100) NULL,
	[FedExAPI_CompanyName] [nvarchar](100) NULL,
	[FedExAPI_Department] [nvarchar](50) NULL,
	[FedExAPI_PhoneNumber] [nvarchar](20) NULL,
	[FedExAPI_PagerNumber] [nvarchar](20) NULL,
	[FedExAPI_FaxNumber] [nvarchar](20) NULL,
	[FedExAPI_EmailAddress] [nvarchar](250) NULL,
	[FedExAPI_Line1] [nvarchar](250) NULL,
	[FedExAPI_Line2] [nvarchar](100) NULL,
	[FedExAPI_City] [nvarchar](100) NULL,
	[FedExAPI_State] [nvarchar](50) NULL,
	[FedExAPI_PostalCode] [nvarchar](50) NULL,
	[FedExAPI_Country] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[fasttransact]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[fasttransact]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[fasttransact](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[AccountID] [nvarchar](50) NULL,
	[SiteTag] [nvarchar](50) NULL,
	[tran_type] [nvarchar](50) NULL,
	[card_types] [nvarchar](50) NULL,
	[CVV2] [int] NULL,
 CONSTRAINT [aaaaafasttransact_PK] PRIMARY KEY NONCLUSTERED 
(
	[id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[eWay]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[eWay]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[eWay](
	[eWayID] [int] NULL,
	[eWayCustomerId] [nvarchar](250) NULL,
	[eWayPostMethod] [nvarchar](250) NULL,
	[eWayTestmode] [int] NULL,
	[eWayCVV] [int] NULL,
	[eWayBeagleActive] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[emailSettings]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[emailSettings]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[emailSettings](
	[id] [int] NOT NULL,
	[ownerEmail] [nvarchar](150) NULL,
	[frmEmail] [nvarchar](150) NULL,
	[ConfirmEmail] [ntext] NULL,
	[PayPalEmail] [ntext] NULL,
	[ReceivedEmail] [ntext] NULL,
	[ShippedEmail] [ntext] NULL,
	[CancelledEmail] [ntext] NULL,
 CONSTRAINT [aaaaaemailSettings_PK] PRIMARY KEY NONCLUSTERED 
(
	[id] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[echo]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[echo]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[echo](
	[id] [int] NULL,
	[transaction_type] [nvarchar](50) NULL,
	[order_type] [nvarchar](50) NULL,
	[merchant_echo_id] [nvarchar](255) NULL,
	[merchant_pin] [nvarchar](255) NULL,
	[merchant_email] [nvarchar](255) NULL,
	[isp_echo_id] [nvarchar](255) NULL,
	[isp_pin] [nvarchar](255) NULL,
	[cnp_security] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[DProducts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DProducts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[DProducts](
	[idProduct] [int] NULL,
	[ProductURL] [nvarchar](255) NULL,
	[URLExpire] [int] NULL,
	[ExpireDays] [int] NULL,
	[License] [int] NULL,
	[LocalLG] [nvarchar](255) NULL,
	[RemoteLG] [nvarchar](255) NULL,
	[LicenseLabel1] [nvarchar](100) NULL,
	[LicenseLabel2] [nvarchar](100) NULL,
	[LicenseLabel3] [nvarchar](100) NULL,
	[LicenseLabel4] [nvarchar](100) NULL,
	[LicenseLabel5] [nvarchar](100) NULL,
	[AddToMail] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[DPRequests]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DPRequests]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[DPRequests](
	[idOrder] [int] NULL,
	[idProduct] [int] NULL,
	[idCustomer] [int] NULL,
	[RequestSTR] [nvarchar](255) NULL,
	[StartDate] [datetime] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[DPLicenses]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DPLicenses]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[DPLicenses](
	[idOrder] [int] NULL,
	[idProduct] [int] NULL,
	[Lic1] [nvarchar](255) NULL,
	[Lic2] [nvarchar](255) NULL,
	[Lic3] [nvarchar](255) NULL,
	[Lic4] [nvarchar](255) NULL,
	[Lic5] [nvarchar](255) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[discountsPerQuantity]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[discountsPerQuantity]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[discountsPerQuantity](
	[idDiscountPerQuantity] [int] IDENTITY(1,1) NOT NULL,
	[idproduct] [int] NULL,
	[idcategory] [int] NOT NULL,
	[discountDesc] [nvarchar](50) NULL,
	[quantityFrom] [int] NULL,
	[quantityUntil] [int] NULL,
	[discountPerUnit] [float] NULL,
	[num] [int] NULL,
	[percentage] [int] NULL,
	[discountPerWUnit] [float] NULL,
	[baseproductonly] [int] NULL,
 CONSTRAINT [aaaaadiscountsPerQuantity_PK] PRIMARY KEY NONCLUSTERED 
(
	[idDiscountPerQuantity] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[discounts]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[discounts]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[discounts](
	[iddiscount] [int] IDENTITY(1,1) NOT NULL,
	[discountdesc] [nvarchar](255) NULL,
	[pricetodiscount] [float] NULL,
	[percentagetodiscount] [float] NULL,
	[discountcode] [nvarchar](50) NULL,
	[active] [int] NULL,
	[used] [int] NULL,
	[onetime] [int] NULL,
	[quantityfrom] [int] NULL,
	[quantityuntil] [int] NULL,
	[weightfrom] [int] NULL,
	[weightuntil] [int] NULL,
	[pricefrom] [float] NULL,
	[priceuntil] [float] NULL,
	[idProduct] [int] NULL,
	[expDate] [datetime] NULL,
	[pcSeparate] [int] NULL,
	[pcDisc_Auto] [int] NULL,
	[pcDisc_StartDate] [datetime] NULL,
	[pcDisc_PerToFlatCartTotal] [money] NULL,
	[pcRetailFlag] [int] NULL,
	[pcDisc_PerToFlatDiscount] [money] NULL,
	[pcWholesaleFlag] [int] NULL,
	[pcDisc_IncExcPrd] [int] NOT NULL,
	[pcDisc_IncExcCat] [int] NOT NULL,
	[pcDisc_IncExcCust] [int] NOT NULL,
	[pcDisc_IncExcCPrice] [int] NOT NULL,
 CONSTRAINT [aaaaadiscounts_PK] PRIMARY KEY NONCLUSTERED 
(
	[iddiscount] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[customfields]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[customfields]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[customfields](
	[idcustom] [int] IDENTITY(1,1) NOT NULL,
	[custom] [nvarchar](255) NULL,
	[searchable] [bit] NOT NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[customers]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[customers]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[customers](
	[idcustomer] [int] IDENTITY(1,1) NOT NULL,
	[name] [nvarchar](255) NULL,
	[lastName] [nvarchar](255) NULL,
	[customerCompany] [nvarchar](255) NULL,
	[phone] [nvarchar](30) NULL,
	[email] [nvarchar](255) NULL,
	[password] [nvarchar](100) NULL,
	[address] [nvarchar](150) NULL,
	[zip] [nvarchar](20) NULL,
	[stateCode] [nvarchar](4) NULL,
	[state] [nvarchar](50) NULL,
	[city] [nvarchar](50) NULL,
	[countryCode] [nvarchar](4) NULL,
	[shippingaddress] [nvarchar](255) NULL,
	[shippingcity] [nvarchar](255) NULL,
	[shippingStateCode] [nvarchar](4) NULL,
	[shippingState] [nvarchar](255) NULL,
	[shippingCountryCode] [nvarchar](4) NULL,
	[shippingZip] [nvarchar](10) NULL,
	[customerType] [int] NOT NULL,
	[TotalOrders] [int] NULL,
	[TotalSales] [int] NULL,
	[iRewardPointsAccrued] [int] NULL,
	[iRewardPointsUsed] [int] NULL,
	[dtRewardsStarted] [datetime] NULL,
	[iRewardPointsHistory] [int] NULL,
	[iRewardPointsHistoryUsed] [int] NULL,
	[IDRefer] [int] NULL,
	[CI1] [nvarchar](255) NULL,
	[CI2] [nvarchar](255) NULL,
	[address2] [nvarchar](255) NULL,
	[shippingCompany] [nvarchar](255) NULL,
	[shippingAddress2] [nvarchar](150) NULL,
	[RecvNews] [int] NULL,
	[suspend] [int] NULL,
	[idCustomerCategory] [int] NULL,
	[fax] [nvarchar](30) NULL,
	[pcCust_DateCreated] [datetime] NULL,
	[pcCust_Locked] [int] NULL,
	[pcCust_SSN] [nvarchar](35) NULL,
	[shippingEmail] [nvarchar](250) NULL,
	[pcCust_VATID] [nvarchar](35) NULL,
	[shippingPhone] [nvarchar](30) NULL,
	[pcCust_EditedDate] [datetime] NULL,
	[pcCust_Guest] [int] NOT NULL,
	[pcCust_Residential] [int] NOT NULL,
	[shippingFax] [nvarchar](50) NULL,
	[pcCust_AgreeTerms] [int] NOT NULL,
	[pcCust_ConsolidateStr] [nvarchar](100) NULL,
	[pcCust_Notes] [nvarchar](4000) NULL,
	[pcCust_AllowReviewEmails] [int] NULL,
 CONSTRAINT [aaaaacustomers_PK] PRIMARY KEY NONCLUSTERED 
(
	[idcustomer] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM dbo.sysindexes WHERE id = OBJECT_ID(N'[dbo].[customers]') AND name = N'idx_customers_idcustomer')
CREATE NONCLUSTERED INDEX [idx_customers_idcustomer] ON [dbo].[customers] 
(
	[idcustomer] ASC
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[customCardTypes]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[customCardTypes]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[customCardTypes](
	[idCustomCardType] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[customCardDesc] [nvarchar](150) NULL,
 CONSTRAINT [aaaaacustomCardTypes_PK] PRIMARY KEY NONCLUSTERED 
(
	[idCustomCardType] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[customCardRules]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[customCardRules]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[customCardRules](
	[idCustomCardRules] [int] IDENTITY(1,1) NOT NULL,
	[idCustomCardType] [int] NULL,
	[ruleName] [nvarchar](250) NULL,
	[intruleRequired] [int] NULL,
	[intlengthOfField] [int] NULL,
	[intmaxInput] [int] NULL,
	[intOrder] [int] NULL,
 CONSTRAINT [aaaaacustomCardRules_PK] PRIMARY KEY NONCLUSTERED 
(
	[idCustomCardRules] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[customCardOrders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[customCardOrders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[customCardOrders](
	[idCCOrder] [int] IDENTITY(1,1) NOT NULL,
	[idOrder] [int] NULL,
	[idcustomCardType] [int] NULL,
	[idCustomCardRules] [int] NULL,
	[strFormValue] [ntext] NULL,
	[intOrderTotal] [money] NULL,
	[strRuleName] [nvarchar](250) NULL,
 CONSTRAINT [aaaaacustomCardOrders_PK] PRIMARY KEY NONCLUSTERED 
(
	[idCCOrder] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[CustCategoryPayTypes]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[CustCategoryPayTypes]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[CustCategoryPayTypes](
	[idCustCategoryPayType] [int] IDENTITY(1,1) NOT NULL,
	[idCustomerCategory] [int] NOT NULL,
	[idPayment] [int] NOT NULL,
 CONSTRAINT [PK_CustCategoryPayTypes] PRIMARY KEY CLUSTERED 
(
	[idCustCategoryPayType] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[cs_relationships]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[cs_relationships]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[cs_relationships](
	[idcrosssell] [int] IDENTITY(1,1) NOT NULL,
	[idproduct] [int] NULL,
	[idrelation] [int] NULL,
	[num] [int] NULL,
	[discount] [float] NULL,
	[isPercent] [int] NULL,
	[isRequired] [int] NULL,
	[cs_type] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[crossSelldata]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[crossSelldata]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[crossSelldata](
	[id] [int] NULL,
	[cs_status] [int] NULL,
	[cs_showprod] [int] NULL,
	[cs_showcart] [int] NULL,
	[cs_showimage] [int] NULL,
	[crossSellText] [nvarchar](255) NULL,
	[cs_ProductViewCnt] [int] NULL,
	[cs_CartViewCnt] [int] NULL,
	[cs_ImageHeight] [int] NULL,
	[cs_ImageWidth] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[creditCards]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[creditCards]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[creditCards](
	[idOrder] [int] NULL,
	[cardtype] [nvarchar](50) NULL,
	[cardnumber] [nvarchar](100) NULL,
	[expiration] [datetime] NULL,
	[seqcode] [nvarchar](10) NULL,
	[comments] [ntext] NULL,
	[pcSecurityKeyID] [int] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[countries]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[countries]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[countries](
	[countryName] [nvarchar](255) NULL,
	[countryCode] [nvarchar](255) NULL,
	[pcSubDivisionID] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[configWishlistSessions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[configWishlistSessions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[configWishlistSessions](
	[idconfigWishlistSession] [int] IDENTITY(1,1) NOT NULL,
	[configKey] [int] NULL,
	[idProduct] [int] NULL,
	[stringProducts] [ntext] NULL,
	[stringValues] [ntext] NULL,
	[stringCategories] [ntext] NULL,
	[stringOptions] [ntext] NULL,
	[idOptionA] [int] NULL,
	[idOptionB] [int] NULL,
	[xfdetails] [ntext] NULL,
	[dtCreated] [datetime] NULL,
	[fPrice] [float] NULL,
	[dPrice] [money] NULL,
	[stringQuantity] [ntext] NULL,
	[stringPrice] [ntext] NULL,
	[stringCProducts] [ntext] NULL,
	[stringCValues] [ntext] NULL,
	[stringCCategories] [ntext] NULL,
	[pcconf_Quantity] [int] NULL,
	[pcconf_QDiscount] [float] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[configSpec_products]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[configSpec_products]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[configSpec_products](
	[specProduct] [int] NULL,
	[configProduct] [int] NULL,
	[price] [float] NULL,
	[Wprice] [float] NULL,
	[cdefault] [bit] NOT NULL,
	[showInfo] [bit] NOT NULL,
	[requiredCategory] [bit] NOT NULL,
	[multiSelect] [bit] NOT NULL,
	[prdSort] [int] NULL,
	[catSort] [int] NULL,
	[configProductCategory] [int] NOT NULL,
	[displayQF] [bit] NOT NULL,
	[Notes] [ntext] NULL,
	[pcConfPro_ShowImg] [int] NULL,
	[pcConfPro_ImgWidth] [int] NULL,
	[pcConfPro_ShowSKU] [int] NULL,
	[pcConfPro_ShowDesc] [int] NULL,
	[pcConfPro_UseRadio] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[configSpec_Charges]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[configSpec_Charges]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[configSpec_Charges](
	[specProduct] [int] NULL,
	[configProduct] [int] NULL,
	[price] [float] NOT NULL,
	[Wprice] [float] NOT NULL,
	[cdefault] [bit] NOT NULL,
	[showInfo] [bit] NOT NULL,
	[requiredCategory] [bit] NOT NULL,
	[multiSelect] [bit] NOT NULL,
	[prdSort] [int] NOT NULL,
	[catSort] [int] NOT NULL,
	[configProductCategory] [int] NOT NULL,
	[displayQF] [bit] NOT NULL,
	[Notes] [ntext] NULL,
	[pcConfCha_ShowImg] [int] NULL,
	[pcConfCha_ImgWidth] [int] NULL,
	[pcConfCha_ShowSKU] [int] NULL,
	[pcConfCha_ShowDesc] [int] NULL,
	[pcConfCha_UseRadio] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[configSpec_categories]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[configSpec_categories]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[configSpec_categories](
	[idProduct] [int] NULL,
	[idCategory] [int] NULL,
	[catOrder] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[configSessions]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[configSessions]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[configSessions](
	[idconfigSession] [int] IDENTITY(1,1) NOT NULL,
	[configKey] [int] NULL,
	[idProduct] [int] NULL,
	[stringProducts] [ntext] NULL,
	[stringValues] [ntext] NULL,
	[stringCategories] [ntext] NULL,
	[stringOptions] [ntext] NULL,
	[dtCreated] [datetime] NULL,
	[stringQuantity] [ntext] NULL,
	[stringPrice] [ntext] NULL,
	[stringCProducts] [ntext] NULL,
	[stringCValues] [ntext] NULL,
	[stringCCategories] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[concord]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[concord]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[concord](
	[idConcord] [int] IDENTITY(1,1) NOT NULL,
	[StoreID] [nvarchar](255) NULL,
	[StoreKey] [nvarchar](255) NULL,
	[CVV] [int] NULL,
	[testmode] [int] NULL,
	[Curcode] [nvarchar](255) NULL,
	[MethodName] [nvarchar](255) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[CCTypes]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[CCTypes]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[CCTypes](
	[idCCType] [int] IDENTITY(1,1) NOT NULL,
	[CCType] [nvarchar](50) NULL,
	[active] [int] NOT NULL,
	[CCcode] [nvarchar](50) NULL,
 CONSTRAINT [aaaaaCCTypes_PK] PRIMARY KEY NONCLUSTERED 
(
	[idCCType] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[categories_products]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[categories_products]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[categories_products](
	[idProduct] [int] NOT NULL,
	[idCategory] [int] NOT NULL,
	[POrder] [int] NULL,
 CONSTRAINT [aaaaacategories_products_PK] PRIMARY KEY NONCLUSTERED 
(
	[idProduct] ASC,
	[idCategory] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[categories]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[categories]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[categories](
	[idCategory] [int] IDENTITY(1,1) NOT NULL,
	[idParentCategory] [int] NULL,
	[tier] [int] NOT NULL,
	[categoryDesc] [nvarchar](255) NULL,
	[serviceSpec] [bit] NOT NULL,
	[required] [bit] NOT NULL,
	[definePrd] [bit] NOT NULL,
	[priority] [int] NULL,
	[multi] [bit] NOT NULL,
	[details] [ntext] NULL,
	[image] [nvarchar](150) NULL,
	[largeimage] [nvarchar](150) NULL,
	[basePrice] [nvarchar](50) NULL,
	[iBTOhide] [int] NOT NULL,
	[SDesc] [ntext] NULL,
	[LDesc] [ntext] NULL,
	[HideDesc] [int] NULL,
	[pcCats_RetailHide] [int] NULL,
	[pcCats_BreadCrumbs] [ntext] NULL,
	[pcCats_SubCategoryView] [int] NULL,
	[pcCats_CategoryColumns] [int] NULL,
	[pcCats_CategoryRows] [int] NULL,
	[pcCats_PageStyle] [nvarchar](4) NULL,
	[pcCats_ProductColumns] [int] NULL,
	[pcCats_ProductRows] [int] NULL,
	[pcCats_FeaturedCategory] [int] NULL,
	[pcCats_FeaturedCategoryImage] [int] NULL,
	[pcCats_MetaKeywords] [ntext] NULL,
	[pcCats_DisplayLayout] [nvarchar](150) NULL,
	[pcCats_MetaTitle] [nvarchar](250) NULL,
	[pcCats_MetaDesc] [ntext] NULL,
	[pcCats_CreatedDate] [datetime] NULL,
	[pcCats_EditedDate] [datetime] NULL,
 CONSTRAINT [aaaaacategories_PK] PRIMARY KEY NONCLUSTERED 
(
	[idCategory] ASC
) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[Brands]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[Brands]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[Brands](
	[IdBrand] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[BrandName] [nvarchar](50) NOT NULL,
	[BrandLogo] [nvarchar](100) NULL,
	[pcBrands_Description] [nvarchar](4000) NULL,
	[pcBrands_Sdescription] [nvarchar](255) NULL,
	[pcBrands_SubBrandsView] [int] NOT NULL,
	[pcBrands_ProductsView] [nvarchar](10) NULL,
	[pcBrands_Active] [int] NOT NULL,
	[pcBrands_Order] [int] NOT NULL,
	[pcBrands_Parent] [int] NOT NULL,
	[pcBrands_MetaTitle] [nvarchar](255) NULL,
	[pcBrands_MetaDesc] [nvarchar](500) NULL,
	[pcBrands_MetaKeywords] [nvarchar](500) NULL,
	[pcBrands_BrandLogoLg] [nvarchar](50) NULL,
 CONSTRAINT [aaaaaBrands_PK] PRIMARY KEY NONCLUSTERED 
(
	[IdBrand] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[BluePay]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[BluePay]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[BluePay](
	[idBluePay] [int] NOT NULL,
	[BPMerchant] [nvarchar](250) NULL,
	[BPTestmode] [int] NOT NULL,
	[BPTransType] [nvarchar](255) NULL,
	[BPInterfaceType] [nvarchar](50) NULL,
	[BPCVC] [nvarchar](50) NULL,
	[BPSECRET_KEY] [nvarchar](250) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[Blackout]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[Blackout]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[Blackout](
	[Blackout_Date] [datetime] NULL,
	[Blackout_Message] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[authorizeNet]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[authorizeNet]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[authorizeNet](
	[id] [int] NULL,
	[x_Type] [nvarchar](50) NULL,
	[x_Login] [nvarchar](100) NULL,
	[x_Password] [nvarchar](100) NULL,
	[x_version] [nvarchar](4) NULL,
	[x_Curcode] [nvarchar](4) NULL,
	[x_Method] [nvarchar](4) NULL,
	[x_AIMType] [nvarchar](50) NULL,
	[x_CVV] [int] NULL,
	[x_testmode] [int] NULL,
	[x_eCheck] [int] NULL,
	[x_secureSource] [int] NULL,
	[x_eCheckPending] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[authorders]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[authorders]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[authorders](
	[idauthorder] [int] IDENTITY(1,1) NOT FOR REPLICATION NOT NULL,
	[idOrder] [int] NULL,
	[amount] [money] NULL,
	[paymentmethod] [nvarchar](250) NULL,
	[transtype] [nvarchar](250) NULL,
	[authcode] [nvarchar](250) NULL,
	[ccnum] [nvarchar](250) NULL,
	[ccexp] [nvarchar](250) NULL,
	[idCustomer] [int] NULL,
	[fname] [nvarchar](250) NULL,
	[lname] [nvarchar](250) NULL,
	[address] [nvarchar](250) NULL,
	[zip] [nvarchar](250) NULL,
	[captured] [int] NULL,
	[trans_id] [nvarchar](250) NULL,
	[pcSecurityKeyID] [int] NOT NULL,
 CONSTRAINT [aaaaaauthorders_PK] PRIMARY KEY NONCLUSTERED 
(
	[idauthorder] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[affiliates]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[affiliates]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[affiliates](
	[idAffiliate] [int] IDENTITY(1,1) NOT NULL,
	[affiliateName] [nvarchar](100) NULL,
	[affiliateEmail] [nvarchar](150) NULL,
	[commission] [float] NULL,
	[affiliateAddress] [nvarchar](150) NULL,
	[affiliateAddress2] [nvarchar](150) NULL,
	[affiliateCity] [nvarchar](50) NULL,
	[affiliateState] [nvarchar](50) NULL,
	[affiliatezip] [nvarchar](50) NULL,
	[affiliatecountryCode] [nvarchar](4) NULL,
	[affiliatephone] [nvarchar](30) NULL,
	[affiliatefax] [nvarchar](30) NULL,
	[affiliateCompany] [nvarchar](150) NULL,
	[pcaff_Password] [nvarchar](250) NULL,
	[pcaff_Active] [int] NULL,
	[pcaff_website] [nvarchar](150) NULL,
 CONSTRAINT [aaaaaaffiliates_PK] PRIMARY KEY NONCLUSTERED 
(
	[idAffiliate] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[admins]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[admins]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[admins](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[idadmin] [int] NULL,
	[adminname] [nvarchar](50) NULL,
	[adminpassword] [nvarchar](100) NULL,
	[adminlevel] [nvarchar](100) NULL,
	[lastlogin] [datetime] NULL,
	[pcSecurityKeyID] [int] NOT NULL,
	[adm_ContactName] [nvarchar](250) NULL,
	[adm_ContactEmail] [nvarchar](250) NULL,
 CONSTRAINT [aaaaaadmins_PK] PRIMARY KEY NONCLUSTERED 
(
	[id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
IF NOT EXISTS (SELECT * FROM ::fn_listextendedproperty(N'MS_Description' , N'USER',N'dbo', N'TABLE',N'admins', NULL,NULL))
EXEC dbo.sp_addextendedproperty @name=N'MS_Description', @value=N'Saves the login details for ProductCart Control Panels Users

' , @level0type=N'USER',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'admins'
GO
/****** Object:  Table [dbo].[ZipCodeValidation]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[ZipCodeValidation]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[ZipCodeValidation](
	[zipcode] [nvarchar](250) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[xfields]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[xfields]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[xfields](
	[idxfield] [int] IDENTITY(1,1) NOT NULL,
	[xfield] [nvarchar](200) NULL,
	[textarea] [int] NULL,
	[widthoffield] [int] NULL,
	[maxlength] [int] NULL,
	[rowlength] [int] NULL,
	[randnum] [int] NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[WorldPay]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[WorldPay]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[WorldPay](
	[wp_id] [int] NULL,
	[WP_Currency] [nvarchar](50) NULL,
	[WP_instID] [nvarchar](50) NULL,
	[WP_testmode] [nvarchar](50) NULL
) ON [PRIMARY]
END
GO
/****** Object:  Table [dbo].[wishList]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[wishList]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[wishList](
	[idCustomer] [int] NULL,
	[idProduct] [int] NULL,
	[idconfigWishlistSession] [int] NOT NULL,
	[QSubmit] [int] NULL,
	[QDate] [datetime] NULL,
	[IDQuote] [int] IDENTITY(1,1) NOT NULL,
	[DiscountCode] [nvarchar](100) NULL,
	[pcwishList_OptionsArray] [ntext] NULL,
	[idOptionA] [int] NULL,
	[idOptionB] [int] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO

/****** Object:  StoredProcedure [dbo].[uspRmvPrdFromSale]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspRmvPrdFromSale]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspRmvPrdFromSale]    Script Date: 02/25/2011 22:54:20 ******/
CREATE PROCEDURE [dbo].[uspRmvPrdFromSale]
@SCID nvarchar(10),
@IDPrd nvarchar(10),
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	DECLARE @TPrice int
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query=''UPDATE Products SET Products.pcSC_ID=Products.active,Products.active=0 WHERE Products.IDProduct ='' + @IDPrd + '';''
	EXEC(@query)
	
	SELECT TOP 1 @TPrice=pcSales_TargetPrice FROM pcSales_BackUp WHERE pcSC_ID=@SCID
	
	IF @TPrice=0
		SET @query=''UPDATE Products SET Products.Price=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID='' + @SCID + '' AND pcSales_BackUp.IDProduct='' + @IDPrd + '';''
	
	IF @TPrice=-1
		SET @query=''UPDATE Products SET Products.bToBPrice=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID='' + @SCID + '' AND pcSales_BackUp.IDProduct='' + @IDPrd + '';''
		
	IF @TPrice>0
		SET @query=''UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=pcSales_BackUp.pcSB_Price FROM pcCC_Pricing, pcSales_BackUp WHERE pcCC_Pricing.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID='' + @SCID + '' AND pcCC_Pricing.idcustomerCategory='' + @TPrice + '' AND pcSales_BackUp.IDProduct='' + @IDPrd + '';''
		
	EXEC(@query)
	
	SET @query=''UPDATE Products SET Products.active=Products.pcSC_ID,Products.pcSC_ID=0 WHERE Products.IDProduct ='' + @IDPrd + '';''
	EXEC(@query)
	
	SET @query=''DELETE FROM pcSales_BackUp WHERE pcSales_BackUp.pcSC_ID='' + @SCID + '' AND pcSales_BackUp.IDProduct='' + @IDPrd + '';''
	EXEC(@query)
	
	SET @query=''UPDATE pcSales_Completed SET pcSC_BUTotal=(SELECT Count(*) FROM Products WHERE Products.pcSC_ID='' + @SCID + '') WHERE pcSales_Completed.pcSC_ID='' + @SCID + '';''
	EXEC(@query)
	
END
' 
END
GO
/****** Object:  StoredProcedure [dbo].[uspRmvBackedUpRecords]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspRmvBackedUpRecords]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspRmvBackedUpRecords]    Script Date: 02/25/2011 22:53:44 ******/
CREATE PROCEDURE [dbo].[uspRmvBackedUpRecords]
@SCID nvarchar(10) ,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	DECLARE @TPrice int
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	DELETE FROM pcSales_BackUp WHERE pcSC_ID=@SCID
	SET @SMCount=@@ROWCOUNT

END
' 
END
GO
/****** Object:  StoredProcedure [dbo].[uspRestorePrices]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[uspRestorePrices]') AND OBJECTPROPERTY(id,N'IsProcedure') = 1)
BEGIN
EXEC dbo.sp_executesql @statement = N'/****** Object:  StoredProcedure [dbo].[uspRestorePrices]    Script Date: 02/25/2011 22:53:19 ******/
CREATE PROCEDURE [dbo].[uspRestorePrices]
@SCID nvarchar(10) ,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	DECLARE @TPrice int
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SELECT TOP 1 @TPrice=pcSales_TargetPrice FROM pcSales_BackUp WHERE pcSC_ID=@SCID
	
	IF @TPrice=0
		SET @query=''UPDATE Products SET Products.Price=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID='' + @SCID + '';''
	
	IF @TPrice=-1
		SET @query=''UPDATE Products SET Products.bToBPrice=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID='' + @SCID + '';''
		
	IF @TPrice>0
		SET @query=''UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=pcSales_BackUp.pcSB_Price FROM pcCC_Pricing, pcSales_BackUp WHERE pcCC_Pricing.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID='' + @SCID + '' AND pcCC_Pricing.idcustomerCategory='' + @TPrice + '';''
		
	EXEC(@query)
	SET @SMCount=@@ROWCOUNT

END
' 
END
GO
/****** Object:  Default [DBX_iddiscount_5438]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_iddiscount_5438]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_iddiscount_5438]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[used_discounts] ADD  CONSTRAINT [DBX_iddiscount_5438]  DEFAULT ((0)) FOR [iddiscount]
END


END
GO
/****** Object:  Default [DBX_idcustomer_10772]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idcustomer_10772]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idcustomer_10772]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[used_discounts] ADD  CONSTRAINT [DBX_idcustomer_10772]  DEFAULT ((0)) FOR [idcustomer]
END


END
GO
/****** Object:  Default [DF__ups_licen__idUPS__72B0FDB1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ups_licen__idUPS__72B0FDB1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ups_licen__idUPS__72B0FDB1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ups_license] ADD  CONSTRAINT [DF__ups_licen__idUPS__72B0FDB1]  DEFAULT ((0)) FOR [idUPS]
END


END
GO
/****** Object:  Default [DF__twoCheckou__v2co__76818E95]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__twoCheckou__v2co__76818E95]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__twoCheckou__v2co__76818E95]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[twoCheckout] ADD  CONSTRAINT [DF__twoCheckou__v2co__76818E95]  DEFAULT ((0)) FOR [v2co]
END


END
GO
/****** Object:  Default [DF_twoCheckout_v2co_TestMode]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_twoCheckout_v2co_TestMode]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_twoCheckout_v2co_TestMode]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[twoCheckout] ADD  CONSTRAINT [DF_twoCheckout_v2co_TestMode]  DEFAULT ((0)) FOR [v2co_TestMode]
END


END
GO
/****** Object:  Default [DF__taxPrd__idProduc__5614BF03]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__taxPrd__idProduc__5614BF03]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__taxPrd__idProduc__5614BF03]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[taxPrd] ADD  CONSTRAINT [DF__taxPrd__idProduc__5614BF03]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__taxPrd__countryC__5708E33C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__taxPrd__countryC__5708E33C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__taxPrd__countryC__5708E33C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[taxPrd] ADD  CONSTRAINT [DF__taxPrd__countryC__5708E33C]  DEFAULT ((0)) FOR [countryCodeEq]
END


END
GO
/****** Object:  Default [DF__taxPrd__stateCod__57FD0775]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__taxPrd__stateCod__57FD0775]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__taxPrd__stateCod__57FD0775]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[taxPrd] ADD  CONSTRAINT [DF__taxPrd__stateCod__57FD0775]  DEFAULT ((0)) FOR [stateCodeEq]
END


END
GO
/****** Object:  Default [DF__taxPrd__zipEq__58F12BAE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__taxPrd__zipEq__58F12BAE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__taxPrd__zipEq__58F12BAE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[taxPrd] ADD  CONSTRAINT [DF__taxPrd__zipEq__58F12BAE]  DEFAULT ((0)) FOR [zipEq]
END


END
GO
/****** Object:  Default [DF__taxPrd__taxPerPr__59E54FE7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__taxPrd__taxPerPr__59E54FE7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__taxPrd__taxPerPr__59E54FE7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[taxPrd] ADD  CONSTRAINT [DF__taxPrd__taxPerPr__59E54FE7]  DEFAULT ((0)) FOR [taxPerProduct]
END


END
GO
/****** Object:  Default [DF__taxLoc__countryC__1411F17C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__taxLoc__countryC__1411F17C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__taxLoc__countryC__1411F17C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[taxLoc] ADD  CONSTRAINT [DF__taxLoc__countryC__1411F17C]  DEFAULT ((0)) FOR [countryCodeEq]
END


END
GO
/****** Object:  Default [DF__taxLoc__stateCod__150615B5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__taxLoc__stateCod__150615B5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__taxLoc__stateCod__150615B5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[taxLoc] ADD  CONSTRAINT [DF__taxLoc__stateCod__150615B5]  DEFAULT ((0)) FOR [stateCodeEq]
END


END
GO
/****** Object:  Default [DF__taxLoc__zipEq__15FA39EE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__taxLoc__zipEq__15FA39EE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__taxLoc__zipEq__15FA39EE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[taxLoc] ADD  CONSTRAINT [DF__taxLoc__zipEq__15FA39EE]  DEFAULT ((0)) FOR [zipEq]
END


END
GO
/****** Object:  Default [DF__taxLoc__taxLoc__16EE5E27]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__taxLoc__taxLoc__16EE5E27]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__taxLoc__taxLoc__16EE5E27]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[taxLoc] ADD  CONSTRAINT [DF__taxLoc__taxLoc__16EE5E27]  DEFAULT ((0)) FOR [taxLoc]
END


END
GO
/****** Object:  Default [DF__suppliers__recei__1BB31344]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__suppliers__recei__1BB31344]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__suppliers__recei__1BB31344]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[suppliers] ADD  CONSTRAINT [DF__suppliers__recei__1BB31344]  DEFAULT ((0)) FOR [receiveSellEmail]
END


END
GO
/****** Object:  Default [DF__suppliers__recei__1CA7377D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__suppliers__recei__1CA7377D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__suppliers__recei__1CA7377D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[suppliers] ADD  CONSTRAINT [DF__suppliers__recei__1CA7377D]  DEFAULT ((0)) FOR [receiveUnderStockAlert]
END


END
GO
/****** Object:  Default [DF__shipServi__servi__253C7D7E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__shipServi__servi__253C7D7E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__shipServi__servi__253C7D7E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipService] ADD  CONSTRAINT [DF__shipServi__servi__253C7D7E]  DEFAULT ((0)) FOR [serviceActive]
END


END
GO
/****** Object:  Default [DF__shipServi__servi__2630A1B7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__shipServi__servi__2630A1B7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__shipServi__servi__2630A1B7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipService] ADD  CONSTRAINT [DF__shipServi__servi__2630A1B7]  DEFAULT ((0)) FOR [servicePriority]
END


END
GO
/****** Object:  Default [DF__shipServi__servi__2724C5F0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__shipServi__servi__2724C5F0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__shipServi__servi__2724C5F0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipService] ADD  CONSTRAINT [DF__shipServi__servi__2724C5F0]  DEFAULT ((0)) FOR [serviceFree]
END


END
GO
/****** Object:  Default [DF__shipServi__servi__2818EA29]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__shipServi__servi__2818EA29]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__shipServi__servi__2818EA29]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipService] ADD  CONSTRAINT [DF__shipServi__servi__2818EA29]  DEFAULT ((0)) FOR [serviceFreeOverAmt]
END


END
GO
/****** Object:  Default [DF__shipServi__servi__290D0E62]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__shipServi__servi__290D0E62]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__shipServi__servi__290D0E62]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipService] ADD  CONSTRAINT [DF__shipServi__servi__290D0E62]  DEFAULT ((0)) FOR [serviceHandlingFee]
END


END
GO
/****** Object:  Default [DF__shipServi__servi__2A01329B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__shipServi__servi__2A01329B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__shipServi__servi__2A01329B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipService] ADD  CONSTRAINT [DF__shipServi__servi__2A01329B]  DEFAULT ((0)) FOR [serviceHandlingIntFee]
END


END
GO
/****** Object:  Default [DF__shipServi__servi__2AF556D4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__shipServi__servi__2AF556D4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__shipServi__servi__2AF556D4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipService] ADD  CONSTRAINT [DF__shipServi__servi__2AF556D4]  DEFAULT ((0)) FOR [serviceShowHandlingFee]
END


END
GO
/****** Object:  Default [DF__shipServi__servi__2BE97B0D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__shipServi__servi__2BE97B0D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__shipServi__servi__2BE97B0D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipService] ADD  CONSTRAINT [DF__shipServi__servi__2BE97B0D]  DEFAULT ((0)) FOR [serviceLimitation]
END


END
GO
/****** Object:  Default [DF_shipService_serviceDefaultRate]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_shipService_serviceDefaultRate]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_shipService_serviceDefaultRate]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipService] ADD  CONSTRAINT [DF_shipService_serviceDefaultRate]  DEFAULT ((0)) FOR [serviceDefaultRate]
END


END
GO
/****** Object:  Default [DF__ShipmentT__price__30AE302A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ShipmentT__price__30AE302A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ShipmentT__price__30AE302A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ShipmentTypes] ADD  CONSTRAINT [DF__ShipmentT__price__30AE302A]  DEFAULT ((0)) FOR [priceToAdd]
END


END
GO
/****** Object:  Default [DF__ShipmentT__activ__31A25463]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ShipmentT__activ__31A25463]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ShipmentT__activ__31A25463]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ShipmentTypes] ADD  CONSTRAINT [DF__ShipmentT__activ__31A25463]  DEFAULT ((0)) FOR [active]
END


END
GO
/****** Object:  Default [DF__ShipmentT__inter__3296789C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ShipmentT__inter__3296789C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ShipmentT__inter__3296789C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ShipmentTypes] ADD  CONSTRAINT [DF__ShipmentT__inter__3296789C]  DEFAULT ((0)) FOR [international]
END


END
GO
/****** Object:  Default [DF__ShipmentT__ipric__338A9CD5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ShipmentT__ipric__338A9CD5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ShipmentT__ipric__338A9CD5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ShipmentTypes] ADD  CONSTRAINT [DF__ShipmentT__ipric__338A9CD5]  DEFAULT ((0)) FOR [ipriceToAdd]
END


END
GO
/****** Object:  Default [DF__shipAlert__shipE__384F51F2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__shipAlert__shipE__384F51F2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__shipAlert__shipE__384F51F2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[shipAlert] ADD  CONSTRAINT [DF__shipAlert__shipE__384F51F2]  DEFAULT ((0)) FOR [shipExists]
END


END
GO
/****** Object:  Default [DF__SB_Settin__Setti__6A70BD6B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Settin__Setti__6A70BD6B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Settin__Setti__6A70BD6B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Settings] ADD  CONSTRAINT [DF__SB_Settin__Setti__6A70BD6B]  DEFAULT ((0)) FOR [Setting_AutoReg]
END


END
GO
/****** Object:  Default [DF__SB_Settin__Setti__6B64E1A4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Settin__Setti__6B64E1A4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Settin__Setti__6B64E1A4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Settings] ADD  CONSTRAINT [DF__SB_Settin__Setti__6B64E1A4]  DEFAULT ((0)) FOR [Setting_RegSuccess]
END


END
GO
/****** Object:  Default [DF__SB_Packag__idPro__5669C4BE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__idPro__5669C4BE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__idPro__5669C4BE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__idPro__5669C4BE]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Is__575DE8F7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Is__575DE8F7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Is__575DE8F7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Is__575DE8F7]  DEFAULT ((0)) FOR [SB_IsLinked]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Am__58520D30]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Am__58520D30]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Am__58520D30]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Am__58520D30]  DEFAULT ((0)) FOR [SB_Amount]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Bi__59463169]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Bi__59463169]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Bi__59463169]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Bi__59463169]  DEFAULT ((0)) FOR [SB_BillingCycles]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Is__5A3A55A2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Is__5A3A55A2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Is__5A3A55A2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Is__5A3A55A2]  DEFAULT ((0)) FOR [SB_IsTrial]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Tr__5B2E79DB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Tr__5B2E79DB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Tr__5B2E79DB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Tr__5B2E79DB]  DEFAULT ((0)) FOR [SB_TrialAmount]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Tr__5C229E14]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Tr__5C229E14]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Tr__5C229E14]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Tr__5C229E14]  DEFAULT ((0)) FOR [SB_TrialBillingCycles]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_St__5D16C24D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_St__5D16C24D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_St__5D16C24D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_St__5D16C24D]  DEFAULT ((0)) FOR [SB_StartsImmediately]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Ag__5E0AE686]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Ag__5E0AE686]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Ag__5E0AE686]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Ag__5E0AE686]  DEFAULT ((0)) FOR [SB_Agree]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Ty__5EFF0ABF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Ty__5EFF0ABF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Ty__5EFF0ABF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Ty__5EFF0ABF]  DEFAULT ((0)) FOR [SB_Type]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Sh__5FF32EF8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Sh__5FF32EF8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Sh__5FF32EF8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Sh__5FF32EF8]  DEFAULT ((0)) FOR [SB_ShowTrialPrice]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Sh__60E75331]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Sh__60E75331]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Sh__60E75331]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Sh__60E75331]  DEFAULT ((0)) FOR [SB_ShowFreeTrial]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Sh__61DB776A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Sh__61DB776A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Sh__61DB776A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Sh__61DB776A]  DEFAULT ((0)) FOR [SB_ShowStartDate]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Sh__62CF9BA3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Sh__62CF9BA3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Sh__62CF9BA3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Sh__62CF9BA3]  DEFAULT ((0)) FOR [SB_ShowReoccurenceDate]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Sh__63C3BFDC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Sh__63C3BFDC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Sh__63C3BFDC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Sh__63C3BFDC]  DEFAULT ((0)) FOR [SB_ShowEOSDate]
END


END
GO
/****** Object:  Default [DF__SB_Packag__SB_Sh__64B7E415]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Packag__SB_Sh__64B7E415]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Packag__SB_Sh__64B7E415]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Packages] ADD  CONSTRAINT [DF__SB_Packag__SB_Sh__64B7E415]  DEFAULT ((0)) FOR [SB_ShowTrialDate]
END


END
GO
/****** Object:  Default [DF__SB_Orders__idOrd__679450C0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__SB_Orders__idOrd__679450C0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__SB_Orders__idOrd__679450C0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[SB_Orders] ADD  CONSTRAINT [DF__SB_Orders__idOrd__679450C0]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__Referrer__Remove__7B4643B2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__Referrer__Remove__7B4643B2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__Referrer__Remove__7B4643B2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[Referrer] ADD  CONSTRAINT [DF__Referrer__Remove__7B4643B2]  DEFAULT ((0)) FOR [Removed]
END


END
GO
/****** Object:  Default [DBX_idCustomer_31663]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idCustomer_31663]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idCustomer_31663]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[recipients] ADD  CONSTRAINT [DBX_idCustomer_31663]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF_recipients_Recipient_Residential]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_recipients_Recipient_Residential]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_recipients_Recipient_Residential]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[recipients] ADD  CONSTRAINT [DF_recipients_Recipient_Residential]  DEFAULT ((1)) FOR [Recipient_Residential]
END


END
GO
/****** Object:  Default [DF__PSIGate__id__76EBA2E9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__PSIGate__id__76EBA2E9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__PSIGate__id__76EBA2E9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[PSIGate] ADD  CONSTRAINT [DF__PSIGate__id__76EBA2E9]  DEFAULT ((0)) FOR [id]
END


END
GO
/****** Object:  Default [DF__protx__idProtx__7ABC33CD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__protx__idProtx__7ABC33CD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__protx__idProtx__7ABC33CD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[protx] ADD  CONSTRAINT [DF__protx__idProtx__7ABC33CD]  DEFAULT ((0)) FOR [idProtx]
END


END
GO
/****** Object:  Default [DF__protx__CVV__7BB05806]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__protx__CVV__7BB05806]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__protx__CVV__7BB05806]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[protx] ADD  CONSTRAINT [DF__protx__CVV__7BB05806]  DEFAULT ((0)) FOR [CVV]
END


END
GO
/****** Object:  Default [DF__protx__ProtxTest__7CA47C3F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__protx__ProtxTest__7CA47C3F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__protx__ProtxTest__7CA47C3F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[protx] ADD  CONSTRAINT [DF__protx__ProtxTest__7CA47C3F]  DEFAULT ((0)) FOR [ProtxTestmode]
END


END
GO
/****** Object:  Default [DF__protx__avs__7D98A078]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__protx__avs__7D98A078]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__protx__avs__7D98A078]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[protx] ADD  CONSTRAINT [DF__protx__avs__7D98A078]  DEFAULT ((0)) FOR [avs]
END


END
GO
/****** Object:  Default [DF_protx_ProtxApply3DSecure]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_protx_ProtxApply3DSecure]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_protx_ProtxApply3DSecure]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[protx] ADD  CONSTRAINT [DF_protx_ProtxApply3DSecure]  DEFAULT ((3)) FOR [ProtxApply3DSecure]
END


END
GO
/****** Object:  Default [DF__ProductsO__idOrd__0169315C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__idOrd__0169315C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__idOrd__0169315C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__idOrd__0169315C]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__ProductsO__idPro__025D5595]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__idPro__025D5595]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__idPro__025D5595]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__idPro__025D5595]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__ProductsO__servi__035179CE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__servi__035179CE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__servi__035179CE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__servi__035179CE]  DEFAULT ((0)) FOR [service]
END


END
GO
/****** Object:  Default [DF__ProductsO__quant__04459E07]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__quant__04459E07]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__quant__04459E07]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__quant__04459E07]  DEFAULT ((0)) FOR [quantity]
END


END
GO
/****** Object:  Default [DF__ProductsO__idOpt__0539C240]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__idOpt__0539C240]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__idOpt__0539C240]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__idOpt__0539C240]  DEFAULT ((0)) FOR [idOptionA]
END


END
GO
/****** Object:  Default [DF__ProductsO__idOpt__062DE679]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__idOpt__062DE679]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__idOpt__062DE679]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__idOpt__062DE679]  DEFAULT ((0)) FOR [idOptionB]
END


END
GO
/****** Object:  Default [DF__ProductsO__unitP__07220AB2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__unitP__07220AB2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__unitP__07220AB2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__unitP__07220AB2]  DEFAULT ((0)) FOR [unitPrice]
END


END
GO
/****** Object:  Default [DF__ProductsO__unitC__08162EEB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__unitC__08162EEB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__unitC__08162EEB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__unitC__08162EEB]  DEFAULT ((0)) FOR [unitCost]
END


END
GO
/****** Object:  Default [DF__ProductsO__idcon__090A5324]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__idcon__090A5324]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__idcon__090A5324]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__idcon__090A5324]  DEFAULT ((0)) FOR [idconfigSession]
END


END
GO
/****** Object:  Default [DF__ProductsO__rmaSu__09FE775D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__rmaSu__09FE775D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__rmaSu__09FE775D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__rmaSu__09FE775D]  DEFAULT ((0)) FOR [rmaSubmitted]
END


END
GO
/****** Object:  Default [DBX_QDiscounts_8053]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_QDiscounts_8053]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_QDiscounts_8053]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DBX_QDiscounts_8053]  DEFAULT ((0)) FOR [QDiscounts]
END


END
GO
/****** Object:  Default [DBX_ItemsDiscounts_20365]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_ItemsDiscounts_20365]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_ItemsDiscounts_20365]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DBX_ItemsDiscounts_20365]  DEFAULT ((0)) FOR [ItemsDiscounts]
END


END
GO
/****** Object:  Default [DF__ProductsO__pcPac__0AF29B96]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__pcPac__0AF29B96]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__pcPac__0AF29B96]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__pcPac__0AF29B96]  DEFAULT ((0)) FOR [pcPackageInfo_ID]
END


END
GO
/****** Object:  Default [DF__ProductsO__pcDro__0BE6BFCF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__pcDro__0BE6BFCF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__pcDro__0BE6BFCF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__pcDro__0BE6BFCF]  DEFAULT ((0)) FOR [pcDropShipper_ID]
END


END
GO
/****** Object:  Default [DF__ProductsO__pcPrd__0DCF0841]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__pcPrd__0DCF0841]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__pcPrd__0DCF0841]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__pcPrd__0DCF0841]  DEFAULT ((0)) FOR [pcPrdOrd_BackOrder]
END


END
GO
/****** Object:  Default [DF__ProductsO__pcPrd__0EC32C7A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__pcPrd__0EC32C7A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__pcPrd__0EC32C7A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__pcPrd__0EC32C7A]  DEFAULT ((0)) FOR [pcPrdOrd_SentNotice]
END


END
GO
/****** Object:  Default [DF__ProductsO__pcPO___0FB750B3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__pcPO___0FB750B3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__pcPO___0FB750B3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__pcPO___0FB750B3]  DEFAULT ((0)) FOR [pcPO_EPID]
END


END
GO
/****** Object:  Default [DF__ProductsO__pcPO___10AB74EC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__pcPO___10AB74EC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__pcPO___10AB74EC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__pcPO___10AB74EC]  DEFAULT ((0)) FOR [pcPO_GWOpt]
END


END
GO
/****** Object:  Default [DF__ProductsO__pcPO___119F9925]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ProductsO__pcPO___119F9925]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ProductsO__pcPO___119F9925]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF__ProductsO__pcPO___119F9925]  DEFAULT ((0)) FOR [pcPO_GWPrice]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPrdOrd_Shipped]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPrdOrd_Shipped]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPrdOrd_Shipped]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPrdOrd_Shipped]  DEFAULT ((0)) FOR [pcPrdOrd_Shipped]
END


END
GO
/****** Object:  Default [DBX_pcPrdOrd_BundledDisc_23038]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcPrdOrd_BundledDisc_23038]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcPrdOrd_BundledDisc_23038]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DBX_pcPrdOrd_BundledDisc_23038]  DEFAULT ((0)) FOR [pcPrdOrd_BundledDisc]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_SubFrequency]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_SubFrequency]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_SubFrequency]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_SubFrequency]  DEFAULT ((0)) FOR [pcPO_SubFrequency]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_SubCycles]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_SubCycles]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_SubCycles]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_SubCycles]  DEFAULT ((0)) FOR [pcPO_SubCycles]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_SubTrialFrequency]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_SubTrialFrequency]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_SubTrialFrequency]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_SubTrialFrequency]  DEFAULT ((0)) FOR [pcPO_SubTrialFrequency]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_SubTrialCycles]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_SubTrialCycles]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_SubTrialCycles]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_SubTrialCycles]  DEFAULT ((0)) FOR [pcPO_SubTrialCycles]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_IsTrial]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_IsTrial]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_IsTrial]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_IsTrial]  DEFAULT ((0)) FOR [pcPO_IsTrial]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_SubAmount]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_SubAmount]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_SubAmount]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_SubAmount]  DEFAULT ((0)) FOR [pcPO_SubAmount]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_SubTrialAmount]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_SubTrialAmount]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_SubTrialAmount]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_SubTrialAmount]  DEFAULT ((0)) FOR [pcPO_SubTrialAmount]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_SubAgree]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_SubAgree]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_SubAgree]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_SubAgree]  DEFAULT ((0)) FOR [pcPO_SubAgree]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_NoShipping]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_NoShipping]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_NoShipping]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_NoShipping]  DEFAULT ((0)) FOR [pcPO_NoShipping]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcPO_SubActive]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcPO_SubActive]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcPO_SubActive]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcPO_SubActive]  DEFAULT ((0)) FOR [pcPO_SubActive]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcSubscription_ID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcSubscription_ID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcSubscription_ID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcSubscription_ID]  DEFAULT ((0)) FOR [pcSubscription_ID]
END


END
GO
/****** Object:  Default [DF_ProductsOrdered_pcSC_ID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ProductsOrdered_pcSC_ID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ProductsOrdered_pcSC_ID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ProductsOrdered] ADD  CONSTRAINT [DF_ProductsOrdered_pcSC_ID]  DEFAULT ((0)) FOR [pcSC_ID]
END


END
GO
/****** Object:  Default [DF__products__idSupp__16644E42]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__idSupp__16644E42]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__idSupp__16644E42]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__idSupp__16644E42]  DEFAULT ((0)) FOR [idSupplier]
END


END
GO
/****** Object:  Default [DF__products__config__1758727B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__config__1758727B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__config__1758727B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__config__1758727B]  DEFAULT ((0)) FOR [configOnly]
END


END
GO
/****** Object:  Default [DF__products__servic__184C96B4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__servic__184C96B4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__servic__184C96B4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__servic__184C96B4]  DEFAULT ((0)) FOR [serviceSpec]
END


END
GO
/****** Object:  Default [DF__products__price__1940BAED]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__price__1940BAED]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__price__1940BAED]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__price__1940BAED]  DEFAULT ((0)) FOR [price]
END


END
GO
/****** Object:  Default [DF__products__listPr__1A34DF26]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__listPr__1A34DF26]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__listPr__1A34DF26]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__listPr__1A34DF26]  DEFAULT ((0)) FOR [listPrice]
END


END
GO
/****** Object:  Default [DF__products__bToBPr__1B29035F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__bToBPr__1B29035F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__bToBPr__1B29035F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__bToBPr__1B29035F]  DEFAULT ((0)) FOR [bToBPrice]
END


END
GO
/****** Object:  Default [DF__products__stock__1C1D2798]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__stock__1C1D2798]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__stock__1C1D2798]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__stock__1C1D2798]  DEFAULT ((0)) FOR [stock]
END


END
GO
/****** Object:  Default [DF__products__listHi__1D114BD1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__listHi__1D114BD1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__listHi__1D114BD1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__listHi__1D114BD1]  DEFAULT ((0)) FOR [listHidden]
END


END
GO
/****** Object:  Default [DF__products__weight__1E05700A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__weight__1E05700A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__weight__1E05700A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__weight__1E05700A]  DEFAULT ((0)) FOR [weight]
END


END
GO
/****** Object:  Default [DF__products__delive__1EF99443]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__delive__1EF99443]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__delive__1EF99443]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__delive__1EF99443]  DEFAULT ((0)) FOR [deliveringTime]
END


END
GO
/****** Object:  Default [DF__products__active__1FEDB87C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__active__1FEDB87C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__active__1FEDB87C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__active__1FEDB87C]  DEFAULT ((0)) FOR [active]
END


END
GO
/****** Object:  Default [DF__products__IdOpti__20E1DCB5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__IdOpti__20E1DCB5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__IdOpti__20E1DCB5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__IdOpti__20E1DCB5]  DEFAULT ((1)) FOR [IdOptionGroupA]
END


END
GO
/****** Object:  Default [DF__products__Arequi__21D600EE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__Arequi__21D600EE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__Arequi__21D600EE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__Arequi__21D600EE]  DEFAULT ((0)) FOR [Arequired]
END


END
GO
/****** Object:  Default [DF__products__IdOpti__22CA2527]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__IdOpti__22CA2527]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__IdOpti__22CA2527]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__IdOpti__22CA2527]  DEFAULT ((1)) FOR [IdOptionGroupB]
END


END
GO
/****** Object:  Default [DF__products__Brequi__23BE4960]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__Brequi__23BE4960]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__Brequi__23BE4960]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__Brequi__23BE4960]  DEFAULT ((0)) FOR [Brequired]
END


END
GO
/****** Object:  Default [DF__products__hotDea__24B26D99]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__hotDea__24B26D99]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__hotDea__24B26D99]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__hotDea__24B26D99]  DEFAULT ((0)) FOR [hotDeal]
END


END
GO
/****** Object:  Default [DF__products__cost__25A691D2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__cost__25A691D2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__cost__25A691D2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__cost__25A691D2]  DEFAULT ((0)) FOR [cost]
END


END
GO
/****** Object:  Default [DF__products__visits__269AB60B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__visits__269AB60B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__visits__269AB60B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__visits__269AB60B]  DEFAULT ((0)) FOR [visits]
END


END
GO
/****** Object:  Default [DF__products__sales__278EDA44]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__sales__278EDA44]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__sales__278EDA44]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__sales__278EDA44]  DEFAULT ((0)) FOR [sales]
END


END
GO
/****** Object:  Default [DF__products__stockL__2882FE7D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__stockL__2882FE7D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__stockL__2882FE7D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__stockL__2882FE7D]  DEFAULT ((0)) FOR [stockLevelAlert]
END


END
GO
/****** Object:  Default [DF__products__formQu__297722B6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__formQu__297722B6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__formQu__297722B6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__formQu__297722B6]  DEFAULT ((0)) FOR [formQuantity]
END


END
GO
/****** Object:  Default [DF__products__showIn__2A6B46EF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__showIn__2A6B46EF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__showIn__2A6B46EF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__showIn__2A6B46EF]  DEFAULT ((0)) FOR [showInHome]
END


END
GO
/****** Object:  Default [DF__products__rndNum__2B5F6B28]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__rndNum__2B5F6B28]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__rndNum__2B5F6B28]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__rndNum__2B5F6B28]  DEFAULT ((0)) FOR [rndNum]
END


END
GO
/****** Object:  Default [DF__products__priori__2C538F61]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__priori__2C538F61]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__priori__2C538F61]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__priori__2C538F61]  DEFAULT ((0)) FOR [priority]
END


END
GO
/****** Object:  Default [DF__products__notax__2D47B39A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__notax__2D47B39A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__notax__2D47B39A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__notax__2D47B39A]  DEFAULT ((0)) FOR [notax]
END


END
GO
/****** Object:  Default [DF__products__noship__2E3BD7D3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__noship__2E3BD7D3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__noship__2E3BD7D3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__noship__2E3BD7D3]  DEFAULT ((0)) FOR [noshipping]
END


END
GO
/****** Object:  Default [DF__products__remove__2F2FFC0C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__remove__2F2FFC0C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__remove__2F2FFC0C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__remove__2F2FFC0C]  DEFAULT ((0)) FOR [removed]
END


END
GO
/****** Object:  Default [DF__products__custom__30242045]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__custom__30242045]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__custom__30242045]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__custom__30242045]  DEFAULT ((0)) FOR [custom1]
END


END
GO
/****** Object:  Default [DF__products__custom__3118447E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__custom__3118447E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__custom__3118447E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__custom__3118447E]  DEFAULT ((0)) FOR [custom2]
END


END
GO
/****** Object:  Default [DF__products__custom__320C68B7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__custom__320C68B7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__custom__320C68B7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__custom__320C68B7]  DEFAULT ((0)) FOR [custom3]
END


END
GO
/****** Object:  Default [DF__products__xfield__33008CF0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__xfield__33008CF0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__xfield__33008CF0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__xfield__33008CF0]  DEFAULT ((0)) FOR [xfield1]
END


END
GO
/****** Object:  Default [DF__products__x1req__33F4B129]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__x1req__33F4B129]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__x1req__33F4B129]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__x1req__33F4B129]  DEFAULT ((0)) FOR [x1req]
END


END
GO
/****** Object:  Default [DF__products__xfield__34E8D562]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__xfield__34E8D562]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__xfield__34E8D562]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__xfield__34E8D562]  DEFAULT ((0)) FOR [xfield2]
END


END
GO
/****** Object:  Default [DF__products__x2req__35DCF99B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__x2req__35DCF99B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__x2req__35DCF99B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__x2req__35DCF99B]  DEFAULT ((0)) FOR [x2req]
END


END
GO
/****** Object:  Default [DF__products__xfield__36D11DD4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__xfield__36D11DD4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__xfield__36D11DD4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__xfield__36D11DD4]  DEFAULT ((0)) FOR [xfield3]
END


END
GO
/****** Object:  Default [DF__products__x3req__37C5420D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__x3req__37C5420D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__x3req__37C5420D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__x3req__37C5420D]  DEFAULT ((0)) FOR [x3req]
END


END
GO
/****** Object:  Default [DF__products__iRewar__38B96646]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__iRewar__38B96646]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__iRewar__38B96646]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__iRewar__38B96646]  DEFAULT ((0)) FOR [iRewardPoints]
END


END
GO
/****** Object:  Default [DF__products__NoPric__39AD8A7F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__NoPric__39AD8A7F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__NoPric__39AD8A7F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__NoPric__39AD8A7F]  DEFAULT ((0)) FOR [NoPrices]
END


END
GO
/****** Object:  Default [DF__products__IDBran__3AA1AEB8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__IDBran__3AA1AEB8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__IDBran__3AA1AEB8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__IDBran__3AA1AEB8]  DEFAULT ((0)) FOR [IDBrand]
END


END
GO
/****** Object:  Default [DF__products__Downlo__3B95D2F1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__Downlo__3B95D2F1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__Downlo__3B95D2F1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__Downlo__3B95D2F1]  DEFAULT ((0)) FOR [Downloadable]
END


END
GO
/****** Object:  Default [DF__products__noStoc__3C89F72A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__noStoc__3C89F72A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__noStoc__3C89F72A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__noStoc__3C89F72A]  DEFAULT ((0)) FOR [noStock]
END


END
GO
/****** Object:  Default [DF__products__noship__3D7E1B63]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__noship__3D7E1B63]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__noship__3D7E1B63]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__noship__3D7E1B63]  DEFAULT ((0)) FOR [noshippingtext]
END


END
GO
/****** Object:  Default [DF__products__pcprod__3E723F9C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcprod__3E723F9C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcprod__3E723F9C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcprod__3E723F9C]  DEFAULT ((0)) FOR [pcprod_HideBTOPrice]
END


END
GO
/****** Object:  Default [DF__products__pcprod__3F6663D5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcprod__3F6663D5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcprod__3F6663D5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcprod__3F6663D5]  DEFAULT ((0)) FOR [pcprod_QtyValidate]
END


END
GO
/****** Object:  Default [DF__products__pcprod__405A880E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcprod__405A880E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcprod__405A880E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcprod__405A880E]  DEFAULT ((0)) FOR [pcprod_MinimumQty]
END


END
GO
/****** Object:  Default [DF__products__pcprod__414EAC47]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcprod__414EAC47]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcprod__414EAC47]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcprod__414EAC47]  DEFAULT ((0)) FOR [pcprod_QtyToPound]
END


END
GO
/****** Object:  Default [DF__products__pcprod__4242D080]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcprod__4242D080]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcprod__4242D080]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcprod__4242D080]  DEFAULT ((0)) FOR [pcprod_OrdInHome]
END


END
GO
/****** Object:  Default [DF__products__pcprod__4336F4B9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcprod__4336F4B9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcprod__4336F4B9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcprod__4336F4B9]  DEFAULT ((0)) FOR [pcprod_HideDefConfig]
END


END
GO
/****** Object:  Default [DF__products__pcProd__442B18F2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcProd__442B18F2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcProd__442B18F2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcProd__442B18F2]  DEFAULT ((0)) FOR [pcProdImage_Columns]
END


END
GO
/****** Object:  Default [DF__products__pcProd__451F3D2B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcProd__451F3D2B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcProd__451F3D2B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcProd__451F3D2B]  DEFAULT ((0)) FOR [pcProd_NotifyStock]
END


END
GO
/****** Object:  Default [DF__products__pcProd__46136164]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcProd__46136164]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcProd__46136164]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcProd__46136164]  DEFAULT ((0)) FOR [pcProd_ReorderLevel]
END


END
GO
/****** Object:  Default [DF__products__pcProd__4707859D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcProd__4707859D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcProd__4707859D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcProd__4707859D]  DEFAULT ((0)) FOR [pcProd_SentNotice]
END


END
GO
/****** Object:  Default [DF__products__pcSupp__47FBA9D6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcSupp__47FBA9D6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcSupp__47FBA9D6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcSupp__47FBA9D6]  DEFAULT ((0)) FOR [pcSupplier_ID]
END


END
GO
/****** Object:  Default [DF__products__pcProd__4AD81681]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcProd__4AD81681]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcProd__4AD81681]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcProd__4AD81681]  DEFAULT ((0)) FOR [pcProd_IsDropShipped]
END


END
GO
/****** Object:  Default [DF__products__pcDrop__4BCC3ABA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcDrop__4BCC3ABA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcDrop__4BCC3ABA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcDrop__4BCC3ABA]  DEFAULT ((0)) FOR [pcDropShipper_ID]
END


END
GO
/****** Object:  Default [DF__products__pcProd__4CC05EF3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcProd__4CC05EF3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcProd__4CC05EF3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcProd__4CC05EF3]  DEFAULT ((0)) FOR [pcProd_BackOrder]
END


END
GO
/****** Object:  Default [DF__products__pcProd__4DB4832C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcProd__4DB4832C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcProd__4DB4832C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcProd__4DB4832C]  DEFAULT ((0)) FOR [pcProd_ShipNDays]
END


END
GO
/****** Object:  Default [DF__products__pcprod__4EA8A765]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcprod__4EA8A765]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcprod__4EA8A765]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcprod__4EA8A765]  DEFAULT ((0)) FOR [pcprod_GC]
END


END
GO
/****** Object:  Default [DF__products__pcProd__4F9CCB9E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__products__pcProd__4F9CCB9E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__products__pcProd__4F9CCB9E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF__products__pcProd__4F9CCB9E]  DEFAULT ((0)) FOR [pcProd_SkipDetailsPage]
END


END
GO
/****** Object:  Default [DF_products_pcProd_BTODefaultPrice]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_products_pcProd_BTODefaultPrice]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_products_pcProd_BTODefaultPrice]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_pcProd_BTODefaultPrice]  DEFAULT ((0)) FOR [pcProd_BTODefaultPrice]
END


END
GO
/****** Object:  Default [DF_products_pcProd_BTODefaultWPrice]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_products_pcProd_BTODefaultWPrice]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_products_pcProd_BTODefaultWPrice]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_pcProd_BTODefaultWPrice]  DEFAULT ((0)) FOR [pcProd_BTODefaultWPrice]
END


END
GO
/****** Object:  Default [DBX_pcProd_HideSKU_20244]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcProd_HideSKU_20244]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcProd_HideSKU_20244]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DBX_pcProd_HideSKU_20244]  DEFAULT ((0)) FOR [pcProd_HideSKU]
END


END
GO
/****** Object:  Default [DF_products_pcProd_SavedTimes]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_products_pcProd_SavedTimes]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_products_pcProd_SavedTimes]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_pcProd_SavedTimes]  DEFAULT ((0)) FOR [pcProd_SavedTimes]
END


END
GO
/****** Object:  Default [DF_products_pcProd_Surcharge]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_products_pcProd_Surcharge]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_products_pcProd_Surcharge]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_pcProd_Surcharge]  DEFAULT ((0)) FOR [pcProd_Surcharge1]
END


END
GO
/****** Object:  Default [DF_products_pcProd_Surcharge2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_products_pcProd_Surcharge2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_products_pcProd_Surcharge2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_pcProd_Surcharge2]  DEFAULT ((0)) FOR [pcProd_Surcharge2]
END


END
GO
/****** Object:  Default [DF_products_pcProd_multiQty]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_products_pcProd_multiQty]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_products_pcProd_multiQty]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_pcProd_multiQty]  DEFAULT ((0)) FOR [pcProd_multiQty]
END


END
GO
/****** Object:  Default [DF_products_pcProd_MaxSelect]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_products_pcProd_MaxSelect]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_products_pcProd_MaxSelect]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_pcProd_MaxSelect]  DEFAULT ((0)) FOR [pcProd_MaxSelect]
END


END
GO
/****** Object:  Default [DF_products_pcPrd_MojoZoom]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_products_pcPrd_MojoZoom]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_products_pcPrd_MojoZoom]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_pcPrd_MojoZoom]  DEFAULT ((0)) FOR [pcPrd_MojoZoom]
END


END
GO
/****** Object:  Default [DF_products_pcProd_AvgRating]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_products_pcProd_AvgRating]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_products_pcProd_AvgRating]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_products_pcProd_AvgRating]  DEFAULT ((0)) FOR [pcProd_AvgRating]
END


END
GO
/****** Object:  Default [DF_Products_pcSC_ID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_Products_pcSC_ID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_Products_pcSC_ID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[products] ADD  CONSTRAINT [DF_Products_pcSC_ID]  DEFAULT ((0)) FOR [pcSC_ID]
END


END
GO
/****** Object:  Default [DF__pfporders__idOrd__5EDF0F2E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pfporders__idOrd__5EDF0F2E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pfporders__idOrd__5EDF0F2E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pfporders] ADD  CONSTRAINT [DF__pfporders__idOrd__5EDF0F2E]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__pfporders__amt__5FD33367]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pfporders__amt__5FD33367]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pfporders__amt__5FD33367]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pfporders] ADD  CONSTRAINT [DF__pfporders__amt__5FD33367]  DEFAULT ((0)) FOR [amt]
END


END
GO
/****** Object:  Default [DF_idCustomer_pfporders]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_idCustomer_pfporders]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_idCustomer_pfporders]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pfporders] ADD  CONSTRAINT [DF_idCustomer_pfporders]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pfporders__captu__60C757A0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pfporders__captu__60C757A0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pfporders__captu__60C757A0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pfporders] ADD  CONSTRAINT [DF__pfporders__captu__60C757A0]  DEFAULT ((0)) FOR [captured]
END


END
GO
/****** Object:  Default [DF_pfporders_pcSecurityKeyID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pfporders_pcSecurityKeyID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pfporders_pcSecurityKeyID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pfporders] ADD  CONSTRAINT [DF_pfporders_pcSecurityKeyID]  DEFAULT ((0)) FOR [pcSecurityKeyID]
END


END
GO
/****** Object:  Default [DF__pcXMLSett__pcXML__028C36D1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLSett__pcXML__028C36D1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLSett__pcXML__028C36D1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLSettings] ADD  CONSTRAINT [DF__pcXMLSett__pcXML__028C36D1]  DEFAULT ((0)) FOR [pcXMLSet_Log]
END


END
GO
/****** Object:  Default [DF__pcXMLSett__pcXML__03805B0A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLSett__pcXML__03805B0A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLSett__pcXML__03805B0A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLSettings] ADD  CONSTRAINT [DF__pcXMLSett__pcXML__03805B0A]  DEFAULT ((0)) FOR [pcXMLSet_LogErrors]
END


END
GO
/****** Object:  Default [DF__pcXMLSett__pcXML__04747F43]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLSett__pcXML__04747F43]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLSett__pcXML__04747F43]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLSettings] ADD  CONSTRAINT [DF__pcXMLSett__pcXML__04747F43]  DEFAULT ((0)) FOR [pcXMLSet_CaptureRequest]
END


END
GO
/****** Object:  Default [DF__pcXMLSett__pcXML__0568A37C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLSett__pcXML__0568A37C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLSett__pcXML__0568A37C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLSettings] ADD  CONSTRAINT [DF__pcXMLSett__pcXML__0568A37C]  DEFAULT ((0)) FOR [pcXMLSet_CaptureResponse]
END


END
GO
/****** Object:  Default [DF__pcXMLSett__pcXML__065CC7B5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLSett__pcXML__065CC7B5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLSett__pcXML__065CC7B5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLSettings] ADD  CONSTRAINT [DF__pcXMLSett__pcXML__065CC7B5]  DEFAULT ((0)) FOR [pcXMLSet_EnforceHTTPs]
END


END
GO
/****** Object:  Default [DF__pcXMLPart__pcXP___7DC781B4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLPart__pcXP___7DC781B4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLPart__pcXP___7DC781B4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLPartners] ADD  CONSTRAINT [DF__pcXMLPart__pcXP___7DC781B4]  DEFAULT ((0)) FOR [pcXP_Status]
END


END
GO
/****** Object:  Default [DF__pcXMLPart__pcXP___7EBBA5ED]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLPart__pcXP___7EBBA5ED]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLPart__pcXP___7EBBA5ED]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLPartners] ADD  CONSTRAINT [DF__pcXMLPart__pcXP___7EBBA5ED]  DEFAULT ((0)) FOR [pcXP_Removed]
END


END
GO
/****** Object:  Default [DF__pcXMLPart__pcXP___7FAFCA26]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLPart__pcXP___7FAFCA26]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLPart__pcXP___7FAFCA26]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLPartners] ADD  CONSTRAINT [DF__pcXMLPart__pcXP___7FAFCA26]  DEFAULT ((0)) FOR [pcXP_ExportAdmin]
END


END
GO
/****** Object:  Default [DF__pcXMLLogs__pcXP___743E177A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLLogs__pcXP___743E177A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLLogs__pcXP___743E177A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLLogs] ADD  CONSTRAINT [DF__pcXMLLogs__pcXP___743E177A]  DEFAULT ((0)) FOR [pcXP_id]
END


END
GO
/****** Object:  Default [DF__pcXMLLogs__pcXL___75323BB3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLLogs__pcXL___75323BB3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLLogs__pcXL___75323BB3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLLogs] ADD  CONSTRAINT [DF__pcXMLLogs__pcXL___75323BB3]  DEFAULT ((0)) FOR [pcXL_RequestType]
END


END
GO
/****** Object:  Default [DF__pcXMLLogs__pcXL___76265FEC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLLogs__pcXL___76265FEC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLLogs__pcXL___76265FEC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLLogs] ADD  CONSTRAINT [DF__pcXMLLogs__pcXL___76265FEC]  DEFAULT ((0)) FOR [pcXL_UpdatedID]
END


END
GO
/****** Object:  Default [DF__pcXMLLogs__pcXL___771A8425]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLLogs__pcXL___771A8425]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLLogs__pcXL___771A8425]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLLogs] ADD  CONSTRAINT [DF__pcXMLLogs__pcXL___771A8425]  DEFAULT ((0)) FOR [pcXL_Undo]
END


END
GO
/****** Object:  Default [DF__pcXMLLogs__pcXL___780EA85E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLLogs__pcXL___780EA85E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLLogs__pcXL___780EA85E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLLogs] ADD  CONSTRAINT [DF__pcXMLLogs__pcXL___780EA85E]  DEFAULT ((0)) FOR [pcXL_ResultCount]
END


END
GO
/****** Object:  Default [DF__pcXMLLogs__pcXL___7902CC97]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLLogs__pcXL___7902CC97]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLLogs__pcXL___7902CC97]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLLogs] ADD  CONSTRAINT [DF__pcXMLLogs__pcXL___7902CC97]  DEFAULT ((0)) FOR [pcXL_LastID]
END


END
GO
/****** Object:  Default [DF__pcXMLLogs__pcXL___79F6F0D0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLLogs__pcXL___79F6F0D0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLLogs__pcXL___79F6F0D0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLLogs] ADD  CONSTRAINT [DF__pcXMLLogs__pcXL___79F6F0D0]  DEFAULT ((0)) FOR [pcXL_UndoID]
END


END
GO
/****** Object:  Default [DF__pcXMLLogs__pcXL___7AEB1509]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLLogs__pcXL___7AEB1509]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLLogs__pcXL___7AEB1509]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLLogs] ADD  CONSTRAINT [DF__pcXMLLogs__pcXL___7AEB1509]  DEFAULT ((0)) FOR [pcXL_Status]
END


END
GO
/****** Object:  Default [DF__pcXMLIPs__pcXIP___7161AACF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLIPs__pcXIP___7161AACF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLIPs__pcXIP___7161AACF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLIPs] ADD  CONSTRAINT [DF__pcXMLIPs__pcXIP___7161AACF]  DEFAULT ((0)) FOR [pcXIP_TurnOn]
END


END
GO
/****** Object:  Default [DF__pcXMLExpo__pcXP___6C9CF5B2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLExpo__pcXP___6C9CF5B2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLExpo__pcXP___6C9CF5B2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLExportLogs] ADD  CONSTRAINT [DF__pcXMLExpo__pcXP___6C9CF5B2]  DEFAULT ((0)) FOR [pcXP_ID]
END


END
GO
/****** Object:  Default [DF__pcXMLExpo__pcXEL__6D9119EB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLExpo__pcXEL__6D9119EB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLExpo__pcXEL__6D9119EB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLExportLogs] ADD  CONSTRAINT [DF__pcXMLExpo__pcXEL__6D9119EB]  DEFAULT ((0)) FOR [pcXEL_ExportedID]
END


END
GO
/****** Object:  Default [DF__pcXMLExpo__pcXEL__6E853E24]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcXMLExpo__pcXEL__6E853E24]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcXMLExpo__pcXEL__6E853E24]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcXMLExportLogs] ADD  CONSTRAINT [DF__pcXMLExpo__pcXEL__6E853E24]  DEFAULT ((0)) FOR [pcXEL_IDType]
END


END
GO
/****** Object:  Default [DF_pcVATRates_pcVATRate_Rate]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcVATRates_pcVATRate_Rate]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcVATRates_pcVATRate_Rate]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcVATRates] ADD  CONSTRAINT [DF_pcVATRates_pcVATRate_Rate]  DEFAULT ((0)) FOR [pcVATRate_Rate]
END


END
GO
/****** Object:  Default [DF__pcUploadF__pcUpl__6C390A4C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcUploadF__pcUpl__6C390A4C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcUploadF__pcUpl__6C390A4C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcUploadFiles] ADD  CONSTRAINT [DF__pcUploadF__pcUpl__6C390A4C]  DEFAULT ((0)) FOR [pcUpld_IDFeedback]
END


END
GO
/****** Object:  Default [DF__pcTaxZone__pcTax__03A67F89]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaxZone__pcTax__03A67F89]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaxZone__pcTax__03A67F89]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxZonesGroups] ADD  CONSTRAINT [DF__pcTaxZone__pcTax__03A67F89]  DEFAULT ((0)) FOR [pcTaxZoneRate_ID]
END


END
GO
/****** Object:  Default [DF__pcTaxZone__pcTax__049AA3C2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaxZone__pcTax__049AA3C2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaxZone__pcTax__049AA3C2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxZonesGroups] ADD  CONSTRAINT [DF__pcTaxZone__pcTax__049AA3C2]  DEFAULT ((0)) FOR [pcTaxZoneDesc_ID]
END


END
GO
/****** Object:  Default [DF__pcTaxZone__pcTax__10416098]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaxZone__pcTax__10416098]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaxZone__pcTax__10416098]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxZoneRates] ADD  CONSTRAINT [DF__pcTaxZone__pcTax__10416098]  DEFAULT ((0)) FOR [pcTaxZone_ID]
END



END
GO
/****** Object:  Default [DF__pcTaxZone__pcTax__113584D1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaxZone__pcTax__113584D1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaxZone__pcTax__113584D1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxZoneRates] ADD  CONSTRAINT [DF__pcTaxZone__pcTax__113584D1]  DEFAULT ((0)) FOR [pcTaxZoneRate_Order]
END


END
GO
/****** Object:  Default [DF__pcTaxZone__pcTax__1229A90A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaxZone__pcTax__1229A90A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaxZone__pcTax__1229A90A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxZoneRates] ADD  CONSTRAINT [DF__pcTaxZone__pcTax__1229A90A]  DEFAULT ((0)) FOR [pcTaxZoneRate_Rate]
END


END
GO
/****** Object:  Default [DF__pcTaxZone__pcTax__131DCD43]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaxZone__pcTax__131DCD43]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaxZone__pcTax__131DCD43]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxZoneRates] ADD  CONSTRAINT [DF__pcTaxZone__pcTax__131DCD43]  DEFAULT ((0)) FOR [pcTaxZoneRate_ApplyToSH]
END


END
GO
/****** Object:  Default [DF__pcTaxZone__pcTax__1411F17C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaxZone__pcTax__1411F17C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaxZone__pcTax__1411F17C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxZoneRates] ADD  CONSTRAINT [DF__pcTaxZone__pcTax__1411F17C]  DEFAULT ((0)) FOR [pcTaxZoneRate_Taxable]
END


END
GO
/****** Object:  Default [DF_pcTaxZoneRates_pcTaxZoneRate_LocalZone]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcTaxZoneRates_pcTaxZoneRate_LocalZone]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcTaxZoneRates_pcTaxZoneRate_LocalZone]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxZoneRates] ADD  CONSTRAINT [DF_pcTaxZoneRates_pcTaxZoneRate_LocalZone]  DEFAULT ((0)) FOR [pcTaxZoneRate_LocalZone]
END


END
GO
/****** Object:  Default [DF__pcTaxGrou__pcTax__7FD5EEA5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaxGrou__pcTax__7FD5EEA5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaxGrou__pcTax__7FD5EEA5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxGroups] ADD  CONSTRAINT [DF__pcTaxGrou__pcTax__7FD5EEA5]  DEFAULT ((0)) FOR [pcTaxZoneDesc_ID]
END


END
GO
/****** Object:  Default [DF__pcTaxGrou__pcTax__00CA12DE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaxGrou__pcTax__00CA12DE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaxGrou__pcTax__00CA12DE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxGroups] ADD  CONSTRAINT [DF__pcTaxGrou__pcTax__00CA12DE]  DEFAULT ((0)) FOR [pcTaxZone_ID]
END


END
GO

/****** Object:  Default [DF_PFL_Authorize_idOrder]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_PFL_Authorize_idOrder]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_PFL_Authorize_idOrder]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PFL_Authorize] ADD  CONSTRAINT [DF_PFL_Authorize_idOrder]  DEFAULT ((0)) FOR [idOrder]
END


END
GO

/****** Object:  Default [DF_PFL_Authorize_captured]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_PFL_Authorize_captured]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_PFL_Authorize_captured]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PFL_Authorize] ADD  CONSTRAINT [DF_PFL_Authorize_captured]  DEFAULT ((0)) FOR [captured]
END


END
GO

/****** Object:  Default [DF_pcTaxEptCust_idCustomer]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcTaxEptCust_idCustomer]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcTaxEptCust_idCustomer]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxEptCust] ADD  CONSTRAINT [DF_pcTaxEptCust_idCustomer]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF_pcTaxEptCust_pcTaxZoneRate_ID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcTaxEptCust_pcTaxZoneRate_ID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcTaxEptCust_pcTaxZoneRate_ID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxEptCust] ADD  CONSTRAINT [DF_pcTaxEptCust_pcTaxZoneRate_ID]  DEFAULT ((0)) FOR [pcTaxZoneRate_ID]
END


END
GO
/****** Object:  Default [DF_pcTaxEpt_pcTEpt_EptAll]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcTaxEpt_pcTEpt_EptAll]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcTaxEpt_pcTEpt_EptAll]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxEpt] ADD  CONSTRAINT [DF_pcTaxEpt_pcTEpt_EptAll]  DEFAULT ((0)) FOR [pcTEpt_EptAll]
END


END
GO
/****** Object:  Default [DF_pcTaxEpt_pcTaxZoneRate_ID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcTaxEpt_pcTaxZoneRate_ID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcTaxEpt_pcTaxZoneRate_ID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaxEpt] ADD  CONSTRAINT [DF_pcTaxEpt_pcTaxZoneRate_ID]  DEFAULT ((0)) FOR [pcTaxZoneRate_ID]
END


END
GO
/****** Object:  Default [DF__pcTaskMan__pcTas__789EE131]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaskMan__pcTas__789EE131]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaskMan__pcTas__789EE131]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaskManager] ADD  CONSTRAINT [DF__pcTaskMan__pcTas__789EE131]  DEFAULT ((0)) FOR [pcTaskNum]
END


END
GO
/****** Object:  Default [DF__pcTaskMan__pcTas__7993056A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcTaskMan__pcTas__7993056A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcTaskMan__pcTas__7993056A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcTaskManager] ADD  CONSTRAINT [DF__pcTaskMan__pcTas__7993056A]  DEFAULT ((0)) FOR [pcTaskComplete]
END


END
GO
/****** Object:  Default [DF__pcSupplie__pcSup__7D63964E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSupplie__pcSup__7D63964E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSupplie__pcSup__7D63964E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSuppliers] ADD  CONSTRAINT [DF__pcSupplie__pcSup__7D63964E]  DEFAULT ((0)) FOR [pcSupplier_IsDropShipper]
END


END
GO
/****** Object:  Default [DF__pcSupplie__pcSup__7E57BA87]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSupplie__pcSup__7E57BA87]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSupplie__pcSup__7E57BA87]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSuppliers] ADD  CONSTRAINT [DF__pcSupplie__pcSup__7E57BA87]  DEFAULT ((0)) FOR [pcSupplier_NoticeType]
END


END
GO
/****** Object:  Default [DF__pcSupplie__pcSup__7F4BDEC0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSupplie__pcSup__7F4BDEC0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSupplie__pcSup__7F4BDEC0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSuppliers] ADD  CONSTRAINT [DF__pcSupplie__pcSup__7F4BDEC0]  DEFAULT ((0)) FOR [pcSupplier_NotifyManually]
END


END
GO
/****** Object:  Default [DF__pcSupplie__pcSup__004002F9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSupplie__pcSup__004002F9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSupplie__pcSup__004002F9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSuppliers] ADD  CONSTRAINT [DF__pcSupplie__pcSup__004002F9]  DEFAULT ((0)) FOR [pcSupplier_CustNotifyUpdates]
END


END
GO
/****** Object:  Default [DF_pcStoreVersions_pcStoreVersion_SP]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcStoreVersions_pcStoreVersion_SP]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcStoreVersions_pcStoreVersion_SP]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreVersions] ADD  CONSTRAINT [DF_pcStoreVersions_pcStoreVersion_SP]  DEFAULT ((0)) FOR [pcStoreVersion_SP]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__041093DD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__041093DD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__041093DD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__041093DD]  DEFAULT ((0)) FOR [pcStoreSettings_QtyLimit]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__0504B816]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__0504B816]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__0504B816]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__0504B816]  DEFAULT ((0)) FOR [pcStoreSettings_AddLimit]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__05F8DC4F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__05F8DC4F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__05F8DC4F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__05F8DC4F]  DEFAULT ((0)) FOR [pcStoreSettings_Pre]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__06ED0088]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__06ED0088]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__06ED0088]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__06ED0088]  DEFAULT ((0)) FOR [pcStoreSettings_CustPre]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__07E124C1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__07E124C1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__07E124C1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__07E124C1]  DEFAULT ((0)) FOR [pcStoreSettings_CatImages]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__08D548FA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__08D548FA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__08D548FA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__08D548FA]  DEFAULT ((0)) FOR [pcStoreSettings_ShowStockLmt]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__09C96D33]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__09C96D33]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__09C96D33]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__09C96D33]  DEFAULT ((0)) FOR [pcStoreSettings_OutOfStockPurchase]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__0ABD916C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__0ABD916C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__0ABD916C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__0ABD916C]  DEFAULT ((0)) FOR [pcStoreSettings_MinPurchase]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__0BB1B5A5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__0BB1B5A5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__0BB1B5A5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__0BB1B5A5]  DEFAULT ((0)) FOR [pcStoreSettings_WholesaleMinPurchase]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__0CA5D9DE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__0CA5D9DE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__0CA5D9DE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__0CA5D9DE]  DEFAULT ((0)) FOR [pcStoreSettings_PrdRow]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__0D99FE17]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__0D99FE17]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__0D99FE17]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__0D99FE17]  DEFAULT ((0)) FOR [pcStoreSettings_PrdRowsPerPage]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__0E8E2250]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__0E8E2250]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__0E8E2250]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__0E8E2250]  DEFAULT ((0)) FOR [pcStoreSettings_CatRow]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__0F824689]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__0F824689]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__0F824689]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__0F824689]  DEFAULT ((0)) FOR [pcStoreSettings_CatRowsPerPage]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__10766AC2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__10766AC2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__10766AC2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__10766AC2]  DEFAULT ((0)) FOR [pcStoreSettings_WL]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__116A8EFB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__116A8EFB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__116A8EFB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__116A8EFB]  DEFAULT ((0)) FOR [pcStoreSettings_TF]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__125EB334]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__125EB334]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__125EB334]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__125EB334]  DEFAULT ((0)) FOR [pcStoreSettings_DisplayStock]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__1352D76D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__1352D76D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__1352D76D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__1352D76D]  DEFAULT ((0)) FOR [pcStoreSettings_HideCategory]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__1446FBA6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__1446FBA6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__1446FBA6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__1446FBA6]  DEFAULT ((0)) FOR [pcStoreSettings_AllowNews]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__153B1FDF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__153B1FDF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__153B1FDF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__153B1FDF]  DEFAULT ((0)) FOR [pcStoreSettings_NewsCheckOut]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__162F4418]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__162F4418]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__162F4418]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__162F4418]  DEFAULT ((0)) FOR [pcStoreSettings_NewsReg]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__17236851]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__17236851]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__17236851]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__17236851]  DEFAULT ((0)) FOR [pcStoreSettings_PCOrd]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__18178C8A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__18178C8A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__18178C8A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__18178C8A]  DEFAULT ((0)) FOR [pcStoreSettings_HideSortPro]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__190BB0C3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__190BB0C3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__190BB0C3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__190BB0C3]  DEFAULT ((0)) FOR [pcStoreSettings_ViewRefer]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__19FFD4FC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__19FFD4FC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__19FFD4FC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__19FFD4FC]  DEFAULT ((0)) FOR [pcStoreSettings_RefNewCheckout]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__1AF3F935]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__1AF3F935]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__1AF3F935]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__1AF3F935]  DEFAULT ((0)) FOR [pcStoreSettings_RefNewReg]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__1BE81D6E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__1BE81D6E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__1BE81D6E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__1BE81D6E]  DEFAULT ((0)) FOR [pcStoreSettings_BrandLogo]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__1CDC41A7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__1CDC41A7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__1CDC41A7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__1CDC41A7]  DEFAULT ((0)) FOR [pcStoreSettings_BrandPro]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__1DD065E0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__1DD065E0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__1DD065E0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__1DD065E0]  DEFAULT ((0)) FOR [pcStoreSettings_RewardsActive]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__1EC48A19]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__1EC48A19]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__1EC48A19]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__1EC48A19]  DEFAULT ((0)) FOR [pcStoreSettings_RewardsIncludeWholesale]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__1FB8AE52]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__1FB8AE52]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__1FB8AE52]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__1FB8AE52]  DEFAULT ((0)) FOR [pcStoreSettings_RewardsPercent]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__20ACD28B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__20ACD28B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__20ACD28B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__20ACD28B]  DEFAULT ((0)) FOR [pcStoreSettings_RewardsReferral]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__21A0F6C4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__21A0F6C4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__21A0F6C4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__21A0F6C4]  DEFAULT ((0)) FOR [pcStoreSettings_RewardsFlat]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__22951AFD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__22951AFD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__22951AFD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__22951AFD]  DEFAULT ((0)) FOR [pcStoreSettings_RewardsFlatValue]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__23893F36]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__23893F36]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__23893F36]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__23893F36]  DEFAULT ((0)) FOR [pcStoreSettings_RewardsPerc]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__247D636F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__247D636F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__247D636F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__247D636F]  DEFAULT ((0)) FOR [pcStoreSettings_RewardsPercValue]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__257187A8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__257187A8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__257187A8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__257187A8]  DEFAULT ((0)) FOR [pcStoreSettings_QDiscountType]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2665ABE1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2665ABE1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2665ABE1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2665ABE1]  DEFAULT ((0)) FOR [pcStoreSettings_BTOdisplayType]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2759D01A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2759D01A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2759D01A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2759D01A]  DEFAULT ((0)) FOR [pcStoreSettings_BTOOutofStockPurchase]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__284DF453]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__284DF453]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__284DF453]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__284DF453]  DEFAULT ((0)) FOR [pcStoreSettings_BTOShowImage]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2942188C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2942188C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2942188C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2942188C]  DEFAULT ((0)) FOR [pcStoreSettings_BTOQuote]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2A363CC5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2A363CC5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2A363CC5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2A363CC5]  DEFAULT ((0)) FOR [pcStoreSettings_BTOQuoteSubmit]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2B2A60FE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2B2A60FE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2B2A60FE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2B2A60FE]  DEFAULT ((0)) FOR [pcStoreSettings_BTOQuoteSubmitOnly]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2C1E8537]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2C1E8537]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2C1E8537]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2C1E8537]  DEFAULT ((0)) FOR [pcStoreSettings_BTODetLinkType]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2D12A970]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2D12A970]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2D12A970]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2D12A970]  DEFAULT ((0)) FOR [pcStoreSettings_BTOPopWidth]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2E06CDA9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2E06CDA9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2E06CDA9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2E06CDA9]  DEFAULT ((0)) FOR [pcStoreSettings_BTOPopHeight]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2EFAF1E2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2EFAF1E2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2EFAF1E2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2EFAF1E2]  DEFAULT ((0)) FOR [pcStoreSettings_BTOPopImage]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__2FEF161B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__2FEF161B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__2FEF161B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__2FEF161B]  DEFAULT ((0)) FOR [pcStoreSettings_ConfigPurchaseOnly]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__30E33A54]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__30E33A54]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__30E33A54]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__30E33A54]  DEFAULT ((0)) FOR [pcStoreSettings_ShowSKU]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__31D75E8D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__31D75E8D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__31D75E8D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__31D75E8D]  DEFAULT ((0)) FOR [pcStoreSettings_ShowSmallImg]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__32CB82C6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__32CB82C6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__32CB82C6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__32CB82C6]  DEFAULT ((0)) FOR [pcStoreSettings_Terms]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__33BFA6FF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__33BFA6FF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__33BFA6FF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__33BFA6FF]  DEFAULT ((0)) FOR [pcStoreSettings_HideRMA]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__34B3CB38]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__34B3CB38]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__34B3CB38]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__34B3CB38]  DEFAULT ((0)) FOR [pcStoreSettings_ShowHD]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__35A7EF71]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__35A7EF71]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__35A7EF71]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__35A7EF71]  DEFAULT ((0)) FOR [pcStoreSettings_StoreUseToolTip]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__369C13AA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__369C13AA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__369C13AA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__369C13AA]  DEFAULT ((1)) FOR [pcStoreSettings_ErrorHandler]
END


END
GO
/****** Object:  Default [DF__pcStoreSe__pcSto__379037E3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcStoreSe__pcSto__379037E3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcStoreSe__pcSto__379037E3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF__pcStoreSe__pcSto__379037E3]  DEFAULT ((1)) FOR [pcStoreSettings_AllowCheckoutWR]
END


END
GO
/****** Object:  Default [DBX_pcStoreSettings_TermsShown_pcStoreSettings]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcStoreSettings_TermsShown_pcStoreSettings]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcStoreSettings_TermsShown_pcStoreSettings]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DBX_pcStoreSettings_TermsShown_pcStoreSettings]  DEFAULT ((0)) FOR [pcStoreSettings_TermsShown]
END


END
GO
/****** Object:  Default [DBX_pcStoreSettings_DisableGiftRegistry_pcStoreSettings]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcStoreSettings_DisableGiftRegistry_pcStoreSettings]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcStoreSettings_DisableGiftRegistry_pcStoreSettings]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DBX_pcStoreSettings_DisableGiftRegistry_pcStoreSettings]  DEFAULT ((0)) FOR [pcStoreSettings_DisableGiftRegistry]
END


END
GO
/****** Object:  Default [DBX_pcStoreSettings_DisableDiscountCodes_pcStoreSettings]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcStoreSettings_DisableDiscountCodes_pcStoreSettings]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcStoreSettings_DisableDiscountCodes_pcStoreSettings]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DBX_pcStoreSettings_DisableDiscountCodes_pcStoreSettings]  DEFAULT ((1)) FOR [pcStoreSettings_DisableDiscountCodes]
END


END
GO
/****** Object:  Default [DF_pcStoreSettings_pcStoreSettings_seoURLs]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcStoreSettings_pcStoreSettings_seoURLs]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcStoreSettings_pcStoreSettings_seoURLs]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF_pcStoreSettings_pcStoreSettings_seoURLs]  DEFAULT ((0)) FOR [pcStoreSettings_seoURLs]
END


END
GO
/****** Object:  Default [DF_pcStoreSettings_pcStoreSettings_QuickBuy]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcStoreSettings_pcStoreSettings_QuickBuy]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcStoreSettings_pcStoreSettings_QuickBuy]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF_pcStoreSettings_pcStoreSettings_QuickBuy]  DEFAULT ((0)) FOR [pcStoreSettings_QuickBuy]
END


END
GO
/****** Object:  Default [DF_pcStoreSettings_pcStoreSettings_ATCEnabled]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcStoreSettings_pcStoreSettings_ATCEnabled]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcStoreSettings_pcStoreSettings_ATCEnabled]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF_pcStoreSettings_pcStoreSettings_ATCEnabled]  DEFAULT ((0)) FOR [pcStoreSettings_ATCEnabled]
END


END
GO
/****** Object:  Default [DF_pcStoreSettings_pcStoreSettings_GuestCheckoutOpt]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcStoreSettings_pcStoreSettings_GuestCheckoutOpt]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcStoreSettings_pcStoreSettings_GuestCheckoutOpt]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF_pcStoreSettings_pcStoreSettings_GuestCheckoutOpt]  DEFAULT ((0)) FOR [pcStoreSettings_GuestCheckoutOpt]
END


END
GO
/****** Object:  Default [DF_pcStoreSettings_pcStoreSettings_RestoreCart]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcStoreSettings_pcStoreSettings_RestoreCart]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcStoreSettings_pcStoreSettings_RestoreCart]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF_pcStoreSettings_pcStoreSettings_RestoreCart]  DEFAULT ((1)) FOR [pcStoreSettings_RestoreCart]
END


END
GO
/****** Object:  Default [DF_pcStoreSettings_pcStoreSettings_AddThisDisplay]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcStoreSettings_pcStoreSettings_AddThisDisplay]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcStoreSettings_pcStoreSettings_AddThisDisplay]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF_pcStoreSettings_pcStoreSettings_AddThisDisplay]  DEFAULT ((0)) FOR [pcStoreSettings_AddThisDisplay]
END


END
GO
/****** Object:  Default [DF_pcStoreSettings_pcStoreSettings_PinterestDisplay]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcStoreSettings_pcStoreSettings_PinterestDisplay]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcStoreSettings_pcStoreSettings_PinterestDisplay]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcStoreSettings] ADD  CONSTRAINT [DF_pcStoreSettings_pcStoreSettings_PinterestDisplay]  DEFAULT ((0)) FOR [pcStoreSettings_PinterestDisplay]
END


END
GO
/****** Object:  Default [DF_pcSecurityKeys_pcActiveKey]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSecurityKeys_pcActiveKey]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSecurityKeys_pcActiveKey]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSecurityKeys] ADD  CONSTRAINT [DF_pcSecurityKeys_pcActiveKey]  DEFAULT ((0)) FOR [pcActiveKey]
END


END
GO
/****** Object:  Default [DF__pcSearchF__idPro__5614BF03]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSearchF__idPro__5614BF03]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSearchF__idPro__5614BF03]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields_Products] ADD  CONSTRAINT [DF__pcSearchF__idPro__5614BF03]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__pcSearchF__idSea__5708E33C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSearchF__idSea__5708E33C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSearchF__idSea__5708E33C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields_Products] ADD  CONSTRAINT [DF__pcSearchF__idSea__5708E33C]  DEFAULT ((0)) FOR [idSearchData]
END


END
GO
/****** Object:  Default [DF__pcSearchF__idSea__55209ACA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSearchF__idSea__55209ACA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSearchF__idSea__55209ACA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields_Mappings] ADD  CONSTRAINT [DF__pcSearchF__idSea__55209ACA]  DEFAULT ((0)) FOR [idSearchField]
END


END
GO
/****** Object:  Default [DF__pcSearchF__idCat__53385258]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSearchF__idCat__53385258]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSearchF__idCat__53385258]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields_Categories] ADD  CONSTRAINT [DF__pcSearchF__idCat__53385258]  DEFAULT ((0)) FOR [idCategory]
END


END
GO
/****** Object:  Default [DF__pcSearchF__idSea__542C7691]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSearchF__idSea__542C7691]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSearchF__idSea__542C7691]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields_Categories] ADD  CONSTRAINT [DF__pcSearchF__idSea__542C7691]  DEFAULT ((0)) FOR [idSearchData]
END


END
GO
/****** Object:  Default [DF_pcSearchFields_pcSearchFieldShow]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSearchFields_pcSearchFieldShow]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSearchFields_pcSearchFieldShow]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields] ADD  CONSTRAINT [DF_pcSearchFields_pcSearchFieldShow]  DEFAULT ((0)) FOR [pcSearchFieldShow]
END


END
GO
/****** Object:  Default [DF_pcSearchFields_pcSearchFieldOrder]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSearchFields_pcSearchFieldOrder]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSearchFields_pcSearchFieldOrder]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields] ADD  CONSTRAINT [DF_pcSearchFields_pcSearchFieldOrder]  DEFAULT ((0)) FOR [pcSearchFieldOrder]
END


END
GO
/****** Object:  Default [DF_pcSearchFields_pcSearchFieldCPShow]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSearchFields_pcSearchFieldCPShow]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSearchFields_pcSearchFieldCPShow]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields] ADD  CONSTRAINT [DF_pcSearchFields_pcSearchFieldCPShow]  DEFAULT ((0)) FOR [pcSearchFieldCPShow]
END


END
GO
/****** Object:  Default [DF_pcSearchFields_pcSearchFieldSearch]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSearchFields_pcSearchFieldSearch]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSearchFields_pcSearchFieldSearch]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields] ADD  CONSTRAINT [DF_pcSearchFields_pcSearchFieldSearch]  DEFAULT ((0)) FOR [pcSearchFieldSearch]
END


END
GO
/****** Object:  Default [DF_pcSearchFields_pcSearchFieldCPSearch]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSearchFields_pcSearchFieldCPSearch]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSearchFields_pcSearchFieldCPSearch]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchFields] ADD  CONSTRAINT [DF_pcSearchFields_pcSearchFieldCPSearch]  DEFAULT ((0)) FOR [pcSearchFieldCPSearch]
END


END
GO
/****** Object:  Default [DF_pcSearchData_idSearchField]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSearchData_idSearchField]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSearchData_idSearchField]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchData] ADD  CONSTRAINT [DF_pcSearchData_idSearchField]  DEFAULT ((0)) FOR [idSearchField]
END


END
GO
/****** Object:  Default [DF_pcSearchData_pcSearchDataOrder]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSearchData_pcSearchDataOrder]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSearchData_pcSearchDataOrder]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSearchData] ADD  CONSTRAINT [DF_pcSearchData_pcSearchDataOrder]  DEFAULT ((0)) FOR [pcSearchDataOrder]
END


END
GO
/****** Object:  Default [DF__pcSavedPr__idPro__3BB5CE82]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSavedPr__idPro__3BB5CE82]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSavedPr__idPro__3BB5CE82]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSavedPrdStats] ADD  CONSTRAINT [DF__pcSavedPr__idPro__3BB5CE82]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__pcSavedPr__pcSPS__3CA9F2BB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcSavedPr__pcSPS__3CA9F2BB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcSavedPr__pcSPS__3CA9F2BB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSavedPrdStats] ADD  CONSTRAINT [DF__pcSavedPr__pcSPS__3CA9F2BB]  DEFAULT ((0)) FOR [pcSPS_SavedTimes]
END


END
GO
/****** Object:  Default [DF_pcSavedCartStatistics_pcSCMonth]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSavedCartStatistics_pcSCMonth]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSavedCartStatistics_pcSCMonth]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSavedCartStatistics] ADD  CONSTRAINT [DF_pcSavedCartStatistics_pcSCMonth]  DEFAULT ((0)) FOR [pcSCMonth]
END


END
GO
/****** Object:  Default [DF_pcSavedCartStatistics_pcSCYear]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSavedCartStatistics_pcSCYear]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSavedCartStatistics_pcSCYear]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSavedCartStatistics] ADD  CONSTRAINT [DF_pcSavedCartStatistics_pcSCYear]  DEFAULT ((0)) FOR [pcSCYear]
END


END
GO
/****** Object:  Default [DF_pcSavedCartStatistics_pcSCTotals]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSavedCartStatistics_pcSCTotals]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSavedCartStatistics_pcSCTotals]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSavedCartStatistics] ADD  CONSTRAINT [DF_pcSavedCartStatistics_pcSCTotals]  DEFAULT ((0)) FOR [pcSCTotals]
END


END
GO
/****** Object:  Default [DF_pcSavedCartStatistics_pcSCAnonymous]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSavedCartStatistics_pcSCAnonymous]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSavedCartStatistics_pcSCAnonymous]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSavedCartStatistics] ADD  CONSTRAINT [DF_pcSavedCartStatistics_pcSCAnonymous]  DEFAULT ((0)) FOR [pcSCAnonymous]
END


END
GO
/****** Object:  Default [DF_pcSavedCarts_idcustomer]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSavedCarts_idcustomer]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSavedCarts_idcustomer]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSavedCarts] ADD  CONSTRAINT [DF_pcSavedCarts_idcustomer]  DEFAULT ((0)) FOR [idcustomer]
END


END
GO
/****** Object:  Default [DF_pcSavedCartArray_SavedCartID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSavedCartArray_SavedCartID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSavedCartArray_SavedCartID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSavedCartArray] ADD  CONSTRAINT [DF_pcSavedCartArray_SavedCartID]  DEFAULT ((0)) FOR [SavedCartID]
END


END
GO
/****** Object:  Default [DF_pcSales_Completed_pcSales_ID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_Completed_pcSales_ID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_Completed_pcSales_ID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_Completed] ADD  CONSTRAINT [DF_pcSales_Completed_pcSales_ID]  DEFAULT ((0)) FOR [pcSales_ID]
END


END
GO
/****** Object:  Default [DF_pcSales_Completed_pcSC_Status]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_Completed_pcSC_Status]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_Completed_pcSC_Status]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_Completed] ADD  CONSTRAINT [DF_pcSales_Completed_pcSC_Status]  DEFAULT ((0)) FOR [pcSC_Status]
END


END
GO
/****** Object:  Default [DF_pcSales_Completed_pcSC_BUTotal]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_Completed_pcSC_BUTotal]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_Completed_pcSC_BUTotal]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_Completed] ADD  CONSTRAINT [DF_pcSales_Completed_pcSC_BUTotal]  DEFAULT ((0)) FOR [pcSC_BUTotal]
END


END
GO
/****** Object:  Default [DF_pcSales_Completed_pcSC_RETotal]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_Completed_pcSC_RETotal]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_Completed_pcSC_RETotal]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_Completed] ADD  CONSTRAINT [DF_pcSales_Completed_pcSC_RETotal]  DEFAULT ((0)) FOR [pcSC_RETotal]
END


END
GO
/****** Object:  Default [DF_pcSales_Completed_pcSC_Archived]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_Completed_pcSC_Archived]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_Completed_pcSC_Archived]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_Completed] ADD  CONSTRAINT [DF_pcSales_Completed_pcSC_Archived]  DEFAULT ((0)) FOR [pcSC_Archived]
END


END
GO
/****** Object:  Default [DF_pcSales_BackUp_pcSC_ID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_BackUp_pcSC_ID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_BackUp_pcSC_ID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_BackUp] ADD  CONSTRAINT [DF_pcSales_BackUp_pcSC_ID]  DEFAULT ((0)) FOR [pcSC_ID]
END


END
GO
/****** Object:  Default [DF_pcSales_BackUp_pcSales_ID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_BackUp_pcSales_ID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_BackUp_pcSales_ID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_BackUp] ADD  CONSTRAINT [DF_pcSales_BackUp_pcSales_ID]  DEFAULT ((0)) FOR [pcSales_ID]
END


END
GO
/****** Object:  Default [DF_pcSales_BackUp_IDProduct]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_BackUp_IDProduct]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_BackUp_IDProduct]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_BackUp] ADD  CONSTRAINT [DF_pcSales_BackUp_IDProduct]  DEFAULT ((0)) FOR [IDProduct]
END


END
GO
/****** Object:  Default [DF_pcSales_BackUp_pcSales_Type]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_BackUp_pcSales_Type]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_BackUp_pcSales_Type]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_BackUp] ADD  CONSTRAINT [DF_pcSales_BackUp_pcSales_Type]  DEFAULT ((0)) FOR [pcSales_TargetPrice]
END


END
GO
/****** Object:  Default [DF_pcSales_BackUp_pcSB_Price]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_BackUp_pcSB_Price]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_BackUp_pcSB_Price]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales_BackUp] ADD  CONSTRAINT [DF_pcSales_BackUp_pcSB_Price]  DEFAULT ((0)) FOR [pcSB_Price]
END


END
GO
/****** Object:  Default [DF_pcSales_pcSales_TargetPrice]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_pcSales_TargetPrice]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_pcSales_TargetPrice]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales] ADD  CONSTRAINT [DF_pcSales_pcSales_TargetPrice]  DEFAULT ((0)) FOR [pcSales_TargetPrice]
END


END
GO
/****** Object:  Default [DF_pcSales_pcSales_Type]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_pcSales_Type]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_pcSales_Type]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales] ADD  CONSTRAINT [DF_pcSales_pcSales_Type]  DEFAULT ((0)) FOR [pcSales_Type]
END


END
GO
/****** Object:  Default [DF_Table_1_pcSales_]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_Table_1_pcSales_]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_Table_1_pcSales_]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales] ADD  CONSTRAINT [DF_Table_1_pcSales_]  DEFAULT ((0)) FOR [pcSales_Relative]
END


END
GO
/****** Object:  Default [DF_pcSales_pcSales_Amount]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_pcSales_Amount]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_pcSales_Amount]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales] ADD  CONSTRAINT [DF_pcSales_pcSales_Amount]  DEFAULT ((0)) FOR [pcSales_Amount]
END


END
GO
/****** Object:  Default [DF_pcSales_pcSales_Round]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_pcSales_Round]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_pcSales_Round]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales] ADD  CONSTRAINT [DF_pcSales_pcSales_Round]  DEFAULT ((0)) FOR [pcSales_Round]
END


END
GO
/****** Object:  Default [DF_pcSales_pcSales_Removed]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcSales_pcSales_Removed]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcSales_pcSales_Removed]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcSales] ADD  CONSTRAINT [DF_pcSales_pcSales_Removed]  DEFAULT ((0)) FOR [pcSales_Removed]
END


END
GO
/****** Object:  Default [DF__pcRevSett__pcRS___3C54ED00]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevSett__pcRS___3C54ED00]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevSett__pcRS___3C54ED00]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF__pcRevSett__pcRS___3C54ED00]  DEFAULT ((0)) FOR [pcRS_RatingType]
END


END
GO
/****** Object:  Default [DF__pcRevSett__pcRS___3D491139]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevSett__pcRS___3D491139]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevSett__pcRS___3D491139]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF__pcRevSett__pcRS___3D491139]  DEFAULT ((0)) FOR [pcRS_MaxRating]
END


END
GO
/****** Object:  Default [DF__pcRevSett__pcRS___3E3D3572]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevSett__pcRS___3E3D3572]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevSett__pcRS___3E3D3572]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF__pcRevSett__pcRS___3E3D3572]  DEFAULT ((0)) FOR [pcRS_Active]
END


END
GO
/****** Object:  Default [DF__pcRevSett__pcRS___3F3159AB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevSett__pcRS___3F3159AB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevSett__pcRS___3F3159AB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF__pcRevSett__pcRS___3F3159AB]  DEFAULT ((0)) FOR [pcRS_ShowRatSum]
END


END
GO
/****** Object:  Default [DF__pcRevSett__pcRS___40257DE4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevSett__pcRS___40257DE4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevSett__pcRS___40257DE4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF__pcRevSett__pcRS___40257DE4]  DEFAULT ((0)) FOR [pcRS_RevCount]
END


END
GO
/****** Object:  Default [DF__pcRevSett__pcRS___4119A21D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevSett__pcRS___4119A21D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevSett__pcRS___4119A21D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF__pcRevSett__pcRS___4119A21D]  DEFAULT ((0)) FOR [pcRS_NeedCheck]
END


END
GO
/****** Object:  Default [DF__pcRevSett__pcRS___420DC656]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevSett__pcRS___420DC656]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevSett__pcRS___420DC656]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF__pcRevSett__pcRS___420DC656]  DEFAULT ((0)) FOR [pcRS_LockPost]
END


END
GO
/****** Object:  Default [DF__pcRevSett__pcRS___4301EA8F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevSett__pcRS___4301EA8F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevSett__pcRS___4301EA8F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF__pcRevSett__pcRS___4301EA8F]  DEFAULT ((0)) FOR [pcRS_PostCount]
END


END
GO
/****** Object:  Default [DF__pcRevSett__pcRS___43F60EC8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevSett__pcRS___43F60EC8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevSett__pcRS___43F60EC8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF__pcRevSett__pcRS___43F60EC8]  DEFAULT ((0)) FOR [pcRS_CalMain]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_SendReviewReminder]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_SendReviewReminder]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_SendReviewReminder]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_SendReviewReminder]  DEFAULT ((0)) FOR [pcRS_SendReviewReminder]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_sendReviewReminderDays]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_sendReviewReminderDays]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_sendReviewReminderDays]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_sendReviewReminderDays]  DEFAULT ((0)) FOR [pcRS_sendReviewReminderDays]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_sendReviewReminderType]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_sendReviewReminderType]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_sendReviewReminderType]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_sendReviewReminderType]  DEFAULT ((0)) FOR [pcRS_sendReviewReminderType]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_sendReviewReminderFormat]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_sendReviewReminderFormat]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_sendReviewReminderFormat]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_sendReviewReminderFormat]  DEFAULT ((0)) FOR [pcRS_sendReviewReminderFormat]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_RewardForReview]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_RewardForReview]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_RewardForReview]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_RewardForReview]  DEFAULT ((0)) FOR [pcRS_RewardForReview]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_RewardForReviewFirstPts]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_RewardForReviewFirstPts]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_RewardForReviewFirstPts]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_RewardForReviewFirstPts]  DEFAULT ((0)) FOR [pcRS_RewardForReviewFirstPts]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_RewardForReviewAdditionalPts]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_RewardForReviewAdditionalPts]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_RewardForReviewAdditionalPts]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_RewardForReviewAdditionalPts]  DEFAULT ((0)) FOR [pcRS_RewardForReviewAdditionalPts]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_RewardForReviewMinLength]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_RewardForReviewMinLength]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_RewardForReviewMinLength]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_RewardForReviewMinLength]  DEFAULT ((0)) FOR [pcRS_RewardForReviewMinLength]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_RewardForReviewMaxPts]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_RewardForReviewMaxPts]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_RewardForReviewMaxPts]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_RewardForReviewMaxPts]  DEFAULT ((0)) FOR [pcRS_RewardForReviewMaxPts]
END


END
GO
/****** Object:  Default [DF_pcRevSettings_pcRS_DisplayRatings]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRevSettings_pcRS_DisplayRatings]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRevSettings_pcRS_DisplayRatings]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevSettings] ADD  CONSTRAINT [DF_pcRevSettings_pcRS_DisplayRatings]  DEFAULT ((0)) FOR [pcRS_DisplayRatings]
END


END
GO
/****** Object:  Default [DF__pcRevList__pcRL___47C69FAC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevList__pcRL___47C69FAC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevList__pcRL___47C69FAC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevLists] ADD  CONSTRAINT [DF__pcRevList__pcRL___47C69FAC]  DEFAULT ((0)) FOR [pcRL_IDField]
END


END
GO
/****** Object:  Default [DF__pcReviewS__pcRS___4B973090]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviewS__pcRS___4B973090]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviewS__pcRS___4B973090]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewSpecials] ADD  CONSTRAINT [DF__pcReviewS__pcRS___4B973090]  DEFAULT ((0)) FOR [pcRS_IDProduct]
END


END
GO
/****** Object:  Default [DF__pcReviews__pcRD___0D2FE9C3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviews__pcRD___0D2FE9C3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviews__pcRD___0D2FE9C3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewsData] ADD  CONSTRAINT [DF__pcReviews__pcRD___0D2FE9C3]  DEFAULT ((0)) FOR [pcRD_IDReview]
END


END
GO
/****** Object:  Default [DF__pcReviews__pcRD___0E240DFC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviews__pcRD___0E240DFC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviews__pcRD___0E240DFC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewsData] ADD  CONSTRAINT [DF__pcReviews__pcRD___0E240DFC]  DEFAULT ((0)) FOR [pcRD_IDField]
END


END
GO
/****** Object:  Default [DF__pcReviews__pcRD___0F183235]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviews__pcRD___0F183235]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviews__pcRD___0F183235]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewsData] ADD  CONSTRAINT [DF__pcReviews__pcRD___0F183235]  DEFAULT ((0)) FOR [pcRD_Feel]
END


END
GO
/****** Object:  Default [DF__pcReviews__pcRD___100C566E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviews__pcRD___100C566E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviews__pcRD___100C566E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewsData] ADD  CONSTRAINT [DF__pcReviews__pcRD___100C566E]  DEFAULT ((0)) FOR [pcRD_Rate]
END


END
GO
/****** Object:  Default [DF__pcReviews__pcRev__658C0CBD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviews__pcRev__658C0CBD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviews__pcRev__658C0CBD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviews] ADD  CONSTRAINT [DF__pcReviews__pcRev__658C0CBD]  DEFAULT ((0)) FOR [pcRev_IDProduct]
END


END
GO
/****** Object:  Default [DF__pcReviews__pcRev__668030F6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviews__pcRev__668030F6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviews__pcRev__668030F6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviews] ADD  CONSTRAINT [DF__pcReviews__pcRev__668030F6]  DEFAULT ((0)) FOR [pcRev_Active]
END


END
GO
/****** Object:  Default [DF__pcReviews__pcRev__6774552F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviews__pcRev__6774552F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviews__pcRev__6774552F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviews] ADD  CONSTRAINT [DF__pcReviews__pcRev__6774552F]  DEFAULT ((0)) FOR [pcRev_MainRate]
END


END
GO
/****** Object:  Default [DF__pcReviews__pcRev__68687968]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviews__pcRev__68687968]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviews__pcRev__68687968]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviews] ADD  CONSTRAINT [DF__pcReviews__pcRev__68687968]  DEFAULT ((0)) FOR [pcRev_MainDRate]
END


END
GO
/****** Object:  Default [DF_pcReviews_pcRev_IDOrder]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcReviews_pcRev_IDOrder]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcReviews_pcRev_IDOrder]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviews] ADD  CONSTRAINT [DF_pcReviews_pcRev_IDOrder]  DEFAULT ((0)) FOR [pcRev_IDOrder]
END


END
GO
/****** Object:  Default [DF_pcReviews_pcRev_IDCustomer]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcReviews_pcRev_IDCustomer]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcReviews_pcRev_IDCustomer]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviews] ADD  CONSTRAINT [DF_pcReviews_pcRev_IDCustomer]  DEFAULT ((1)) FOR [pcRev_IDCustomer]
END


END
GO
/****** Object:  Default [DF__pcReviewP__pcRP___37E53D9E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviewP__pcRP___37E53D9E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviewP__pcRP___37E53D9E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewPoints] ADD  CONSTRAINT [DF__pcReviewP__pcRP___37E53D9E]  DEFAULT ((0)) FOR [pcRP_IDReview]
END


END
GO
/****** Object:  Default [DF__pcReviewP__pcRP___38D961D7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviewP__pcRP___38D961D7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviewP__pcRP___38D961D7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewPoints] ADD  CONSTRAINT [DF__pcReviewP__pcRP___38D961D7]  DEFAULT ((0)) FOR [pcRP_IDCustomer]
END


END
GO
/****** Object:  Default [DF__pcReviewP__pcRP___39CD8610]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviewP__pcRP___39CD8610]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviewP__pcRP___39CD8610]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewPoints] ADD  CONSTRAINT [DF__pcReviewP__pcRP___39CD8610]  DEFAULT ((0)) FOR [pcRP_PointsAwarded]
END


END
GO
/****** Object:  Default [DF__pcReviewN__pcRN___3508D0F3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviewN__pcRN___3508D0F3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviewN__pcRN___3508D0F3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewNotifications] ADD  CONSTRAINT [DF__pcReviewN__pcRN___3508D0F3]  DEFAULT ((0)) FOR [pcRN_idCustomer]
END


END
GO
/****** Object:  Default [DF__pcReviewN__pcRN___35FCF52C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcReviewN__pcRN___35FCF52C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcReviewN__pcRN___35FCF52C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcReviewNotifications] ADD  CONSTRAINT [DF__pcReviewN__pcRN___35FCF52C]  DEFAULT ((0)) FOR [pcRN_idOrder]
END


END
GO
/****** Object:  Default [DF__pcRevFiel__pcRF___4F32B74A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevFiel__pcRF___4F32B74A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevFiel__pcRF___4F32B74A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevFields] ADD  CONSTRAINT [DF__pcRevFiel__pcRF___4F32B74A]  DEFAULT ((0)) FOR [pcRF_Type]
END


END
GO
/****** Object:  Default [DF__pcRevFiel__pcRF___5026DB83]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevFiel__pcRF___5026DB83]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevFiel__pcRF___5026DB83]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevFields] ADD  CONSTRAINT [DF__pcRevFiel__pcRF___5026DB83]  DEFAULT ((0)) FOR [pcRF_Active]
END


END
GO
/****** Object:  Default [DF__pcRevFiel__pcRF___511AFFBC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevFiel__pcRF___511AFFBC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevFiel__pcRF___511AFFBC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevFields] ADD  CONSTRAINT [DF__pcRevFiel__pcRF___511AFFBC]  DEFAULT ((0)) FOR [pcRF_Required]
END


END
GO
/****** Object:  Default [DF__pcRevFiel__pcRF___520F23F5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevFiel__pcRF___520F23F5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevFiel__pcRF___520F23F5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevFields] ADD  CONSTRAINT [DF__pcRevFiel__pcRF___520F23F5]  DEFAULT ((0)) FOR [pcRF_Order]
END


END
GO
/****** Object:  Default [DF__pcRevExc__pcRE_I__5F3414E9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcRevExc__pcRE_I__5F3414E9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcRevExc__pcRE_I__5F3414E9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRevExc] ADD  CONSTRAINT [DF__pcRevExc__pcRE_I__5F3414E9]  DEFAULT ((0)) FOR [pcRE_IDProduct]
END


END
GO
/****** Object:  Default [DF__PCReturns__rmaAp__65E11278]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__PCReturns__rmaAp__65E11278]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__PCReturns__rmaAp__65E11278]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[PCReturns] ADD  CONSTRAINT [DF__PCReturns__rmaAp__65E11278]  DEFAULT ((0)) FOR [rmaApproved]
END


END
GO
/****** Object:  Default [DF_pcRecentRevSettings_pcRR_RecentRevCount]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRecentRevSettings_pcRR_RecentRevCount]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRecentRevSettings_pcRR_RecentRevCount]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRecentRevSettings] ADD  CONSTRAINT [DF_pcRecentRevSettings_pcRR_RecentRevCount]  DEFAULT ((0)) FOR [pcRR_RecentRevCount]
END


END
GO
/****** Object:  Default [DF_pcRecentRevSettings_pcRR_RevDays]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRecentRevSettings_pcRR_RevDays]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRecentRevSettings_pcRR_RevDays]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRecentRevSettings] ADD  CONSTRAINT [DF_pcRecentRevSettings_pcRR_RevDays]  DEFAULT ((0)) FOR [pcRR_RevDays]
END


END
GO
/****** Object:  Default [DF_pcRecentRevSettings_pcRR_NotForSale]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRecentRevSettings_pcRR_NotForSale]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRecentRevSettings_pcRR_NotForSale]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRecentRevSettings] ADD  CONSTRAINT [DF_pcRecentRevSettings_pcRR_NotForSale]  DEFAULT ((0)) FOR [pcRR_NotForSale]
END


END
GO
/****** Object:  Default [DF_pcRecentRevSettings_pcRR_OutOfStock]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRecentRevSettings_pcRR_OutOfStock]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRecentRevSettings_pcRR_OutOfStock]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRecentRevSettings] ADD  CONSTRAINT [DF_pcRecentRevSettings_pcRR_OutOfStock]  DEFAULT ((0)) FOR [pcRR_OutOfStock]
END


END
GO
/****** Object:  Default [DF_pcRecentRevSettings_pcRR_SKU]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRecentRevSettings_pcRR_SKU]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRecentRevSettings_pcRR_SKU]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRecentRevSettings] ADD  CONSTRAINT [DF_pcRecentRevSettings_pcRR_SKU]  DEFAULT ((0)) FOR [pcRR_SKU]
END


END
GO
/****** Object:  Default [DF_pcRecentRevSettings_pcRR_ShowImg]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRecentRevSettings_pcRR_ShowImg]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRecentRevSettings_pcRR_ShowImg]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRecentRevSettings] ADD  CONSTRAINT [DF_pcRecentRevSettings_pcRR_ShowImg]  DEFAULT ((0)) FOR [pcRR_ShowImg]
END


END
GO
/****** Object:  Default [DF_pcRecentRevSettings_pcRR_ReviewsPerProduct]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcRecentRevSettings_pcRR_ReviewsPerProduct]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcRecentRevSettings_pcRR_ReviewsPerProduct]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcRecentRevSettings] ADD  CONSTRAINT [DF_pcRecentRevSettings_pcRR_ReviewsPerProduct]  DEFAULT ((1)) FOR [pcRR_ReviewsPerProduct]
END


END
GO
/****** Object:  Default [DF__pcProduct__idPro__6AA5C795]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcProduct__idPro__6AA5C795]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcProduct__idPro__6AA5C795]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcProductsOrderedOptions] ADD  CONSTRAINT [DF__pcProduct__idPro__6AA5C795]  DEFAULT ((0)) FOR [idProductOrdered]
END


END
GO
/****** Object:  Default [DF__pcProduct__idopt__6B99EBCE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcProduct__idopt__6B99EBCE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcProduct__idopt__6B99EBCE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcProductsOrderedOptions] ADD  CONSTRAINT [DF__pcProduct__idopt__6B99EBCE]  DEFAULT ((0)) FOR [idoptoptgrp]
END


END
GO
/****** Object:  Default [DF__pcProduct__idPro__705EA0EB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcProduct__idPro__705EA0EB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcProduct__idPro__705EA0EB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcProductsOptions] ADD  CONSTRAINT [DF__pcProduct__idPro__705EA0EB]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__pcProduct__idOpt__7152C524]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcProduct__idOpt__7152C524]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcProduct__idOpt__7152C524]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcProductsOptions] ADD  CONSTRAINT [DF__pcProduct__idOpt__7152C524]  DEFAULT ((0)) FOR [idOptionGroup]
END


END
GO
/****** Object:  Default [DF__pcProduct__pcPro__7246E95D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcProduct__pcPro__7246E95D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcProduct__pcPro__7246E95D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcProductsOptions] ADD  CONSTRAINT [DF__pcProduct__pcPro__7246E95D]  DEFAULT ((0)) FOR [pcProdOpt_Required]
END


END
GO
/****** Object:  Default [DF__pcProduct__pcPro__733B0D96]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcProduct__pcPro__733B0D96]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcProduct__pcPro__733B0D96]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcProductsOptions] ADD  CONSTRAINT [DF__pcProduct__pcPro__733B0D96]  DEFAULT ((0)) FOR [pcProdOpt_Order]
END


END
GO
/****** Object:  Default [DF__pcProduct__idPro__77FFC2B3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcProduct__idPro__77FFC2B3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcProduct__idPro__77FFC2B3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcProductsImages] ADD  CONSTRAINT [DF__pcProduct__idPro__77FFC2B3]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__pcProduct__pcPro__78F3E6EC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcProduct__pcPro__78F3E6EC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcProduct__pcPro__78F3E6EC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcProductsImages] ADD  CONSTRAINT [DF__pcProduct__pcPro__78F3E6EC]  DEFAULT ((0)) FOR [pcProdImage_Order]
END


END
GO
/****** Object:  Default [DF__pcProduct__pcPE___7DB89C09]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcProduct__pcPE___7DB89C09]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcProduct__pcPE___7DB89C09]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcProductsExc] ADD  CONSTRAINT [DF__pcProduct__pcPE___7DB89C09]  DEFAULT ((0)) FOR [pcPE_IDProduct]
END


END
GO
/****** Object:  Default [DF__pcPriorit__pcPri__39CD8610]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPriorit__pcPri__39CD8610]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPriorit__pcPri__39CD8610]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPriority] ADD  CONSTRAINT [DF__pcPriorit__pcPri__39CD8610]  DEFAULT ((0)) FOR [pcPri_ShowImg]
END


END
GO
/****** Object:  Default [DF_pcPrdPromotions_idproduct]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPrdPromotions_idproduct]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPrdPromotions_idproduct]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPrdPromotions] ADD  CONSTRAINT [DF_pcPrdPromotions_idproduct]  DEFAULT ((0)) FOR [idproduct]
END


END
GO
/****** Object:  Default [DF_pcPrdPromotions_pcPrdPro_QtyTrigger]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPrdPromotions_pcPrdPro_QtyTrigger]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPrdPromotions_pcPrdPro_QtyTrigger]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPrdPromotions] ADD  CONSTRAINT [DF_pcPrdPromotions_pcPrdPro_QtyTrigger]  DEFAULT ((0)) FOR [pcPrdPro_QtyTrigger]
END


END
GO
/****** Object:  Default [DF_pcPrdPromotions_pcPrdPro_DiscountType]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPrdPromotions_pcPrdPro_DiscountType]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPrdPromotions_pcPrdPro_DiscountType]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPrdPromotions] ADD  CONSTRAINT [DF_pcPrdPromotions_pcPrdPro_DiscountType]  DEFAULT ((0)) FOR [pcPrdPro_DiscountType]
END


END
GO
/****** Object:  Default [DF_pcPrdPromotions_pcPrdPro_ApplyUnits]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPrdPromotions_pcPrdPro_ApplyUnits]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPrdPromotions_pcPrdPro_ApplyUnits]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPrdPromotions] ADD  CONSTRAINT [DF_pcPrdPromotions_pcPrdPro_ApplyUnits]  DEFAULT ((0)) FOR [pcPrdPro_ApplyUnits]
END


END
GO
/****** Object:  Default [DF__pcPrdProm__pcPrd__092A4EB5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPrdProm__pcPrd__092A4EB5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPrdProm__pcPrd__092A4EB5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPrdPromotions] ADD  CONSTRAINT [DF__pcPrdProm__pcPrd__092A4EB5]  DEFAULT ((0)) FOR [pcPrdPro_Inactive]
END


END
GO
/****** Object:  Default [DF__pcPrdProm__pcPrd__0A1E72EE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPrdProm__pcPrd__0A1E72EE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPrdProm__pcPrd__0A1E72EE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPrdPromotions] ADD  CONSTRAINT [DF__pcPrdProm__pcPrd__0A1E72EE]  DEFAULT ((0)) FOR [pcPrdPro_IncExcCust]
END


END
GO
/****** Object:  Default [DF__pcPrdProm__pcPrd__0B129727]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPrdProm__pcPrd__0B129727]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPrdProm__pcPrd__0B129727]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPrdPromotions] ADD  CONSTRAINT [DF__pcPrdProm__pcPrd__0B129727]  DEFAULT ((0)) FOR [pcPrdPro_IncExcCPrice]
END


END
GO
/****** Object:  Default [DF__pcPrdProm__pcPrd__0C06BB60]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPrdProm__pcPrd__0C06BB60]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPrdProm__pcPrd__0C06BB60]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPrdPromotions] ADD  CONSTRAINT [DF__pcPrdProm__pcPrd__0C06BB60]  DEFAULT ((0)) FOR [pcPrdPro_RetailFlag]
END


END
GO
/****** Object:  Default [DF__pcPrdProm__pcPrd__0CFADF99]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPrdProm__pcPrd__0CFADF99]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPrdProm__pcPrd__0CFADF99]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPrdPromotions] ADD  CONSTRAINT [DF__pcPrdProm__pcPrd__0CFADF99]  DEFAULT ((0)) FOR [pcPrdPro_WholesaleFlag]
END


END
GO
/****** Object:  Default [DF_pcPPFProducts_pcPrdPro_id]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPPFProducts_pcPrdPro_id]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPPFProducts_pcPrdPro_id]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPPFProducts] ADD  CONSTRAINT [DF_pcPPFProducts_pcPrdPro_id]  DEFAULT ((0)) FOR [pcPrdPro_id]
END


END
GO
/****** Object:  Default [DF_pcPPFProducts_idproduct]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPPFProducts_idproduct]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPPFProducts_idproduct]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPPFProducts] ADD  CONSTRAINT [DF_pcPPFProducts_idproduct]  DEFAULT ((0)) FOR [idproduct]
END


END
GO
/****** Object:  Default [DF__pcPPFCust__pcPrd__1A54DAB7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPPFCust__pcPrd__1A54DAB7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPPFCust__pcPrd__1A54DAB7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPPFCusts] ADD  CONSTRAINT [DF__pcPPFCust__pcPrd__1A54DAB7]  DEFAULT ((0)) FOR [pcPrdPro_id]
END


END
GO
/****** Object:  Default [DF__pcPPFCust__idCus__1B48FEF0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPPFCust__idCus__1B48FEF0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPPFCust__idCus__1B48FEF0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPPFCusts] ADD  CONSTRAINT [DF__pcPPFCust__idCus__1B48FEF0]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcPPFCust__pcPrd__2E5BD364]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPPFCust__pcPrd__2E5BD364]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPPFCust__pcPrd__2E5BD364]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPPFCustPriceCats] ADD  CONSTRAINT [DF__pcPPFCust__pcPrd__2E5BD364]  DEFAULT ((0)) FOR [pcPrdPro_id]
END


END
GO
/****** Object:  Default [DF__pcPPFCust__idCus__2F4FF79D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPPFCust__idCus__2F4FF79D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPPFCust__idCus__2F4FF79D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPPFCustPriceCats] ADD  CONSTRAINT [DF__pcPPFCust__idCus__2F4FF79D]  DEFAULT ((0)) FOR [idCustomerCategory]
END


END
GO
/****** Object:  Default [DF_pcPPFCategories_pcPrdPro_id]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPPFCategories_pcPrdPro_id]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPPFCategories_pcPrdPro_id]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPPFCategories] ADD  CONSTRAINT [DF_pcPPFCategories_pcPrdPro_id]  DEFAULT ((0)) FOR [pcPrdPro_id]
END


END
GO
/****** Object:  Default [DF_pcPPFCategories_idcategory]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPPFCategories_idcategory]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPPFCategories_idcategory]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPPFCategories] ADD  CONSTRAINT [DF_pcPPFCategories_idcategory]  DEFAULT ((0)) FOR [idcategory]
END


END
GO
/****** Object:  Default [DF_pcPPFCategories_pcPPFCats_IncSubCats]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPPFCategories_pcPPFCats_IncSubCats]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPPFCategories_pcPPFCats_IncSubCats]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPPFCategories] ADD  CONSTRAINT [DF_pcPPFCategories_pcPPFCats_IncSubCats]  DEFAULT ((0)) FOR [pcPPFCats_IncSubCats]
END


END
GO
/****** Object:  Default [DF__pcPay_USA__idOrd__064DE20A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_USA__idOrd__064DE20A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_USA__idOrd__064DE20A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay_Orders] ADD  CONSTRAINT [DF__pcPay_USA__idOrd__064DE20A]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__pcPay_USA__Amoun__07420643]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_USA__Amoun__07420643]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_USA__Amoun__07420643]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay_Orders] ADD  CONSTRAINT [DF__pcPay_USA__Amoun__07420643]  DEFAULT ((0)) FOR [Amount]
END


END
GO
/****** Object:  Default [DF__pcPay_USA__idCus__08362A7C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_USA__idCus__08362A7C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_USA__idCus__08362A7C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay_Orders] ADD  CONSTRAINT [DF__pcPay_USA__idCus__08362A7C]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcPay_USA__captu__092A4EB5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_USA__captu__092A4EB5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_USA__captu__092A4EB5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay_Orders] ADD  CONSTRAINT [DF__pcPay_USA__captu__092A4EB5]  DEFAULT ((0)) FOR [captured]
END


END
GO
/****** Object:  Default [DF_pcPay_USAePay_Orders_pcSecurityKeyID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPay_USAePay_Orders_pcSecurityKeyID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPay_USAePay_Orders_pcSecurityKeyID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay_Orders] ADD  CONSTRAINT [DF_pcPay_USAePay_Orders_pcSecurityKeyID]  DEFAULT ((0)) FOR [pcSecurityKeyID]
END


END
GO
/****** Object:  Default [DF__pcPay_USA__pcPay__5792F321]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_USA__pcPay__5792F321]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_USA__pcPay__5792F321]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay] ADD  CONSTRAINT [DF__pcPay_USA__pcPay__5792F321]  DEFAULT ((0)) FOR [pcPay_Uep_Id]
END


END
GO
/****** Object:  Default [DF__pcPay_USA__pcPay__5887175A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_USA__pcPay__5887175A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_USA__pcPay__5887175A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay] ADD  CONSTRAINT [DF__pcPay_USA__pcPay__5887175A]  DEFAULT ((0)) FOR [pcPay_Uep_TransType]
END


END
GO
/****** Object:  Default [DF__pcPay_USA__pcPay__597B3B93]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_USA__pcPay__597B3B93]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_USA__pcPay__597B3B93]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay] ADD  CONSTRAINT [DF__pcPay_USA__pcPay__597B3B93]  DEFAULT ((0)) FOR [pcPay_Uep_TestMode]
END


END
GO
/****** Object:  Default [DF__pcPay_USA__pcPay__5A6F5FCC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_USA__pcPay__5A6F5FCC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_USA__pcPay__5A6F5FCC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay] ADD  CONSTRAINT [DF__pcPay_USA__pcPay__5A6F5FCC]  DEFAULT ((0)) FOR [pcPay_Uep_Checking]
END


END
GO
/****** Object:  Default [DF__pcPay_USA__pcPay__5B638405]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_USA__pcPay__5B638405]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_USA__pcPay__5B638405]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_USAePay] ADD  CONSTRAINT [DF__pcPay_USA__pcPay__5B638405]  DEFAULT ((0)) FOR [pcPay_Uep_CheckPending]
END


END
GO
/****** Object:  Default [DF__pcPay_Tri__pcPay__12B3B8EF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Tri__pcPay__12B3B8EF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Tri__pcPay__12B3B8EF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_TripleDeal] ADD  CONSTRAINT [DF__pcPay_Tri__pcPay__12B3B8EF]  DEFAULT ((0)) FOR [pcPay_TD_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_Tri__pcPay__13A7DD28]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Tri__pcPay__13A7DD28]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Tri__pcPay__13A7DD28]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_TripleDeal] ADD  CONSTRAINT [DF__pcPay_Tri__pcPay__13A7DD28]  DEFAULT ((0)) FOR [pcPay_TD_PayPeriod]
END


END
GO
/****** Object:  Default [DF__pcPay_Tri__pcPay__149C0161]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Tri__pcPay__149C0161]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Tri__pcPay__149C0161]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_TripleDeal] ADD  CONSTRAINT [DF__pcPay_Tri__pcPay__149C0161]  DEFAULT ((0)) FOR [pcPay_TD_TestMode]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__idOrd__1C8112FE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__idOrd__1C8112FE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__idOrd__1C8112FE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal_Authorize] ADD  CONSTRAINT [DF__pcPay_Pay__idOrd__1C8112FE]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__order__1D753737]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__order__1D753737]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__order__1D753737]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal_Authorize] ADD  CONSTRAINT [DF__pcPay_Pay__order__1D753737]  DEFAULT ((2)) FOR [orderStatus]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__idCus__1E695B70]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__idCus__1E695B70]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__idCus__1E695B70]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal_Authorize] ADD  CONSTRAINT [DF__pcPay_Pay__idCus__1E695B70]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__captu__1F5D7FA9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__captu__1F5D7FA9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__captu__1F5D7FA9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal_Authorize] ADD  CONSTRAINT [DF__pcPay_Pay__captu__1F5D7FA9]  DEFAULT ((0)) FOR [captured]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__gwCode__1F5D7FA9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__gwCode__1F5D7FA9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__gwCode__1F5D7FA9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal_Authorize] ADD  CONSTRAINT [DF__pcPay_Pay__gwCode__1F5D7FA9]  DEFAULT ((0)) FOR [gwCode]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__pcPay__186C9245]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__pcPay__186C9245]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__pcPay__186C9245]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal] ADD  CONSTRAINT [DF__pcPay_Pay__pcPay__186C9245]  DEFAULT ((0)) FOR [pcPay_PayPal_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__pcPay__1960B67E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__pcPay__1960B67E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__pcPay__1960B67E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal] ADD  CONSTRAINT [DF__pcPay_Pay__pcPay__1960B67E]  DEFAULT ((0)) FOR [pcPay_PayPal_TransType]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__pcPay__1A54DAB7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__pcPay__1A54DAB7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__pcPay__1A54DAB7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal] ADD  CONSTRAINT [DF__pcPay_Pay__pcPay__1A54DAB7]  DEFAULT ((0)) FOR [pcPay_PayPal_AVS]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__pcPay__1B48FEF0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__pcPay__1B48FEF0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__pcPay__1B48FEF0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal] ADD  CONSTRAINT [DF__pcPay_Pay__pcPay__1B48FEF0]  DEFAULT ((0)) FOR [pcPay_PayPal_CVC]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__pcPay__1C3D2329]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__pcPay__1C3D2329]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__pcPay__1C3D2329]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_PayPal] ADD  CONSTRAINT [DF__pcPay_Pay__pcPay__1C3D2329]  DEFAULT ((0)) FOR [pcPay_PayPal_Sandbox]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__pcPay__66E41C5C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__pcPay__66E41C5C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__pcPay__66E41C5C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_Paymentech] ADD  CONSTRAINT [DF__pcPay_Pay__pcPay__66E41C5C]  DEFAULT ((0)) FOR [pcPay_PT_Id]
END


END
GO
/****** Object:  Default [DF__pcPay_Pay__pcPay__67D84095]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Pay__pcPay__67D84095]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Pay__pcPay__67D84095]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_Paymentech] ADD  CONSTRAINT [DF__pcPay_Pay__pcPay__67D84095]  DEFAULT ((0)) FOR [pcPay_PT_CVC]
END


END
GO
/****** Object:  Default [DF__pcPay_Par__pcPay__27AED5D5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Par__pcPay__27AED5D5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Par__pcPay__27AED5D5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_ParaData] ADD  CONSTRAINT [DF__pcPay_Par__pcPay__27AED5D5]  DEFAULT ((0)) FOR [pcPay_ParaData_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_Par__pcPay__28A2FA0E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Par__pcPay__28A2FA0E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Par__pcPay__28A2FA0E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_ParaData] ADD  CONSTRAINT [DF__pcPay_Par__pcPay__28A2FA0E]  DEFAULT ((0)) FOR [pcPay_ParaData_CVC]
END


END
GO
/****** Object:  Default [DF__pcPay_Par__pcPay__29971E47]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Par__pcPay__29971E47]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Par__pcPay__29971E47]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_ParaData] ADD  CONSTRAINT [DF__pcPay_Par__pcPay__29971E47]  DEFAULT ((0)) FOR [pcPay_ParaData_AVS]
END


END
GO
/****** Object:  Default [DF__pcPay_Par__pcPay__2A8B4280]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Par__pcPay__2A8B4280]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Par__pcPay__2A8B4280]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_ParaData] ADD  CONSTRAINT [DF__pcPay_Par__pcPay__2A8B4280]  DEFAULT ((0)) FOR [pcPay_ParaData_TestMode]
END


END
GO
/****** Object:  Default [DF__pcPay_Ord__pcPay__2E5BD364]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Ord__pcPay__2E5BD364]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Ord__pcPay__2E5BD364]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_OrdersMoneris] ADD  CONSTRAINT [DF__pcPay_Ord__pcPay__2E5BD364]  DEFAULT ((0)) FOR [pcPay_MOrder_OrderID]
END


END
GO
/****** Object:  Default [DF__pcPay_NET__pcPay__322C6448]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_NET__pcPay__322C6448]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_NET__pcPay__322C6448]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_NETOne] ADD  CONSTRAINT [DF__pcPay_NET__pcPay__322C6448]  DEFAULT ((0)) FOR [pcPay_NETOne_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_NET__pcPay__33208881]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_NET__pcPay__33208881]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_NET__pcPay__33208881]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_NETOne] ADD  CONSTRAINT [DF__pcPay_NET__pcPay__33208881]  DEFAULT ((0)) FOR [pcPay_NETOne_CVV]
END


END
GO
/****** Object:  Default [DF__pcPay_Mon__pcPay__2042BE37]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Mon__pcPay__2042BE37]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Mon__pcPay__2042BE37]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_Moneris] ADD  CONSTRAINT [DF__pcPay_Mon__pcPay__2042BE37]  DEFAULT ((0)) FOR [pcPay_Moneris_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_Mon__pcPay__2136E270]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Mon__pcPay__2136E270]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Mon__pcPay__2136E270]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_Moneris] ADD  CONSTRAINT [DF__pcPay_Mon__pcPay__2136E270]  DEFAULT ((0)) FOR [pcPay_Moneris_TestMode]
END


END
GO
/****** Object:  Default [DF_pcPay_Moneris_pcPay_Moneris_CVVEnabled]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPay_Moneris_pcPay_Moneris_CVVEnabled]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPay_Moneris_pcPay_Moneris_CVVEnabled]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_Moneris] ADD  CONSTRAINT [DF_pcPay_Moneris_pcPay_Moneris_CVVEnabled]  DEFAULT ((0)) FOR [pcPay_Moneris_CVVEnabled]
END


END
GO
/****** Object:  Default [DF__pcPay_Lin__pcPay__438BFA74]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Lin__pcPay__438BFA74]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Lin__pcPay__438BFA74]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_LinkPointAPI] ADD  CONSTRAINT [DF__pcPay_Lin__pcPay__438BFA74]  DEFAULT ((0)) FOR [pcPay_LPAPI_OrderStatus]
END


END
GO
/****** Object:  Default [DF__pcPay_Lin__pcPay__44801EAD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Lin__pcPay__44801EAD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Lin__pcPay__44801EAD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_LinkPointAPI] ADD  CONSTRAINT [DF__pcPay_Lin__pcPay__44801EAD]  DEFAULT ((0)) FOR [pcPay_LPAPI_Amount]
END


END
GO
/****** Object:  Default [DF__pcPay_Lin__idCus__457442E6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Lin__idCus__457442E6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Lin__idCus__457442E6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_LinkPointAPI] ADD  CONSTRAINT [DF__pcPay_Lin__idCus__457442E6]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcPay_Lin__pcPay__4668671F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Lin__pcPay__4668671F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Lin__pcPay__4668671F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_LinkPointAPI] ADD  CONSTRAINT [DF__pcPay_Lin__pcPay__4668671F]  DEFAULT ((0)) FOR [pcPay_LPAPI_Captured]
END


END
GO
/****** Object:  Default [DF_pcPay_LinkPointAPI_pcSecurityKeyID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPay_LinkPointAPI_pcSecurityKeyID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPay_LinkPointAPI_pcSecurityKeyID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_LinkPointAPI] ADD  CONSTRAINT [DF_pcPay_LinkPointAPI_pcSecurityKeyID]  DEFAULT ((0)) FOR [pcSecurityKeyID]
END


END
GO
/****** Object:  Default [DF__pcPay_HSB__pcPay__0CFADF99]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_HSB__pcPay__0CFADF99]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_HSB__pcPay__0CFADF99]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_HSBC] ADD  CONSTRAINT [DF__pcPay_HSB__pcPay__0CFADF99]  DEFAULT ((0)) FOR [pcPay_HSBC_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_HSB__pcPay__0DEF03D2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_HSB__pcPay__0DEF03D2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_HSB__pcPay__0DEF03D2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_HSBC] ADD  CONSTRAINT [DF__pcPay_HSB__pcPay__0DEF03D2]  DEFAULT ((0)) FOR [pcPay_HSBC_CVV]
END


END
GO
/****** Object:  Default [DF__pcPay_HSB__pcPay__0EE3280B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_HSB__pcPay__0EE3280B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_HSB__pcPay__0EE3280B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_HSBC] ADD  CONSTRAINT [DF__pcPay_HSB__pcPay__0EE3280B]  DEFAULT ((0)) FOR [pcPay_HSBC_TestMode]
END


END
GO
/****** Object:  Default [DF_PAY1_AMOUNT_pcPay_GestPay_Response]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_PAY1_AMOUNT_pcPay_GestPay_Response]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_PAY1_AMOUNT_pcPay_GestPay_Response]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_GestPay_Response] ADD  CONSTRAINT [DF_PAY1_AMOUNT_pcPay_GestPay_Response]  DEFAULT ((0)) FOR [PAY1_AMOUNT]
END


END
GO
/****** Object:  Default [DBX_pcPay_GestPay_OTP_pcPay_GestPay_OTP]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcPay_GestPay_OTP_pcPay_GestPay_OTP]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcPay_GestPay_OTP_pcPay_GestPay_OTP]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_GestPay_OTP] ADD  CONSTRAINT [DBX_pcPay_GestPay_OTP_pcPay_GestPay_OTP]  DEFAULT ((0)) FOR [pcPay_GestPay_OTP]
END


END
GO
/****** Object:  Default [DF__pcPay_Ges__pcPay__027D5126]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Ges__pcPay__027D5126]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Ges__pcPay__027D5126]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_GestPay_OTP] ADD  CONSTRAINT [DF__pcPay_Ges__pcPay__027D5126]  DEFAULT ((0)) FOR [pcPay_GestPay_OTP_Used]
END


END
GO
/****** Object:  Default [DF__pcPay_Ges__pcPay__6521F869]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Ges__pcPay__6521F869]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Ges__pcPay__6521F869]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_GestPay] ADD  CONSTRAINT [DF__pcPay_Ges__pcPay__6521F869]  DEFAULT ((0)) FOR [pcPay_GestPay_Id]
END


END
GO
/****** Object:  Default [DF__pcPay_Ges__pcPay__66161CA2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Ges__pcPay__66161CA2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Ges__pcPay__66161CA2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_GestPay] ADD  CONSTRAINT [DF__pcPay_Ges__pcPay__66161CA2]  DEFAULT ((0)) FOR [pcPay_GestPay_idLanguage]
END


END
GO
/****** Object:  Default [DF__pcPay_Ges__pcPay__670A40DB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Ges__pcPay__670A40DB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Ges__pcPay__670A40DB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_GestPay] ADD  CONSTRAINT [DF__pcPay_Ges__pcPay__670A40DB]  DEFAULT ((0)) FOR [pcPay_GestPay_idCurrency]
END


END
GO
/****** Object:  Default [DF__pcPay_Fas__pcPay__6ADAD1BF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Fas__pcPay__6ADAD1BF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Fas__pcPay__6ADAD1BF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_FastCharge] ADD  CONSTRAINT [DF__pcPay_Fas__pcPay__6ADAD1BF]  DEFAULT ((0)) FOR [pcPay_FAC_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_Fas__pcPay__6BCEF5F8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Fas__pcPay__6BCEF5F8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Fas__pcPay__6BCEF5F8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_FastCharge] ADD  CONSTRAINT [DF__pcPay_Fas__pcPay__6BCEF5F8]  DEFAULT ((0)) FOR [pcPay_FAC_TransType]
END


END
GO
/****** Object:  Default [DF__pcPay_Fas__pcPay__6CC31A31]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Fas__pcPay__6CC31A31]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Fas__pcPay__6CC31A31]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_FastCharge] ADD  CONSTRAINT [DF__pcPay_Fas__pcPay__6CC31A31]  DEFAULT ((0)) FOR [pcPay_FAC_CVV]
END


END
GO
/****** Object:  Default [DF__pcPay_Fas__pcPay__6DB73E6A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Fas__pcPay__6DB73E6A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Fas__pcPay__6DB73E6A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_FastCharge] ADD  CONSTRAINT [DF__pcPay_Fas__pcPay__6DB73E6A]  DEFAULT ((0)) FOR [pcPay_FAC_Checking]
END


END
GO
/****** Object:  Default [DF__pcPay_Fas__pcPay__6EAB62A3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Fas__pcPay__6EAB62A3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Fas__pcPay__6EAB62A3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_FastCharge] ADD  CONSTRAINT [DF__pcPay_Fas__pcPay__6EAB62A3]  DEFAULT ((0)) FOR [pcPay_FAC_CheckPending]
END


END
GO
/****** Object:  Default [DF__pcPay_EPN__pcPay__727BF387]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EPN__pcPay__727BF387]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EPN__pcPay__727BF387]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EPN] ADD  CONSTRAINT [DF__pcPay_EPN__pcPay__727BF387]  DEFAULT ((0)) FOR [pcPay_EPN_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_EPN__pcPay__737017C0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EPN__pcPay__737017C0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EPN__pcPay__737017C0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EPN] ADD  CONSTRAINT [DF__pcPay_EPN__pcPay__737017C0]  DEFAULT ((0)) FOR [pcPay_EPN_CVV]
END


END
GO
/****** Object:  Default [DF__pcPay_EPN__pcPay__74643BF9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EPN__pcPay__74643BF9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EPN__pcPay__74643BF9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EPN] ADD  CONSTRAINT [DF__pcPay_EPN__pcPay__74643BF9]  DEFAULT ((0)) FOR [pcPay_EPN_TestMode]
END


END
GO
/****** Object:  Default [DF__pcPay_eMe__pcPay__5303482E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_eMe__pcPay__5303482E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_eMe__pcPay__5303482E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_eMerchant] ADD  CONSTRAINT [DF__pcPay_eMe__pcPay__5303482E]  DEFAULT ((0)) FOR [pcPay_eMerch_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_eMe__pcPay__53F76C67]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_eMe__pcPay__53F76C67]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_eMe__pcPay__53F76C67]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_eMerchant] ADD  CONSTRAINT [DF__pcPay_eMe__pcPay__53F76C67]  DEFAULT ((0)) FOR [pcPay_eMerch_CVD]
END


END
GO
/****** Object:  Default [DF__pcPay_eMe__pcPay__54EB90A0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_eMe__pcPay__54EB90A0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_eMe__pcPay__54EB90A0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_eMerchant] ADD  CONSTRAINT [DF__pcPay_eMe__pcPay__54EB90A0]  DEFAULT ((0)) FOR [pcPay_eMerch_TestMode]
END


END
GO
/****** Object:  Default [DF__pcPay_eMe__idOrd__4E3E9311]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_eMe__idOrd__4E3E9311]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_eMe__idOrd__4E3E9311]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_eMerch_Orders] ADD  CONSTRAINT [DF__pcPay_eMe__idOrd__4E3E9311]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__pcPay_eMe__idCus__4F32B74A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_eMe__idCus__4F32B74A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_eMe__idCus__4F32B74A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_eMerch_Orders] ADD  CONSTRAINT [DF__pcPay_eMe__idCus__4F32B74A]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcPay_eMe__pcPay__5026DB83]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_eMe__pcPay__5026DB83]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_eMe__pcPay__5026DB83]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_eMerch_Orders] ADD  CONSTRAINT [DF__pcPay_eMe__pcPay__5026DB83]  DEFAULT ((0)) FOR [pcPay_eMerch_Ord_Captured]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__idOrd__4E9398CC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__idOrd__4E9398CC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__idOrd__4E9398CC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG_Vault] ADD  CONSTRAINT [DF__pcPay_EIG__idOrd__4E9398CC]  DEFAULT ((1)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__idCus__4F87BD05]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__idCus__4F87BD05]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__idCus__4F87BD05]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG_Vault] ADD  CONSTRAINT [DF__pcPay_EIG__idCus__4F87BD05]  DEFAULT ((1)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__IsSav__507BE13E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__IsSav__507BE13E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__IsSav__507BE13E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG_Vault] ADD  CONSTRAINT [DF__pcPay_EIG__IsSav__507BE13E]  DEFAULT ((1)) FOR [IsSaved]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__idOrd__526429B0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__idOrd__526429B0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__idOrd__526429B0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG_Authorize] ADD  CONSTRAINT [DF__pcPay_EIG__idOrd__526429B0]  DEFAULT ((1)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__idCus__53584DE9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__idCus__53584DE9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__idCus__53584DE9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG_Authorize] ADD  CONSTRAINT [DF__pcPay_EIG__idCus__53584DE9]  DEFAULT ((1)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__captu__544C7222]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__captu__544C7222]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__captu__544C7222]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG_Authorize] ADD  CONSTRAINT [DF__pcPay_EIG__captu__544C7222]  DEFAULT ((1)) FOR [captured]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__pcSec__5540965B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__pcSec__5540965B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__pcSec__5540965B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG_Authorize] ADD  CONSTRAINT [DF__pcPay_EIG__pcSec__5540965B]  DEFAULT ((1)) FOR [pcSecurityKeyID]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__amoun__5634BA94]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__amoun__5634BA94]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__amoun__5634BA94]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG_Authorize] ADD  CONSTRAINT [DF__pcPay_EIG__amoun__5634BA94]  DEFAULT ((0)) FOR [amount]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__pcPay__48DABF76]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__pcPay__48DABF76]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__pcPay__48DABF76]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG] ADD  CONSTRAINT [DF__pcPay_EIG__pcPay__48DABF76]  DEFAULT ((1)) FOR [pcPay_EIG_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__pcPay__49CEE3AF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__pcPay__49CEE3AF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__pcPay__49CEE3AF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG] ADD  CONSTRAINT [DF__pcPay_EIG__pcPay__49CEE3AF]  DEFAULT ((0)) FOR [pcPay_EIG_CVV]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__pcPay__4AC307E8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__pcPay__4AC307E8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__pcPay__4AC307E8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG] ADD  CONSTRAINT [DF__pcPay_EIG__pcPay__4AC307E8]  DEFAULT ((0)) FOR [pcPay_EIG_SaveCards]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__pcPay__4BB72C21]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__pcPay__4BB72C21]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__pcPay__4BB72C21]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG] ADD  CONSTRAINT [DF__pcPay_EIG__pcPay__4BB72C21]  DEFAULT ((0)) FOR [pcPay_EIG_UseVault]
END


END
GO
/****** Object:  Default [DF__pcPay_EIG__pcPay__4CAB505A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_EIG__pcPay__4CAB505A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_EIG__pcPay__4CAB505A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_EIG] ADD  CONSTRAINT [DF__pcPay_EIG__pcPay__4CAB505A]  DEFAULT ((0)) FOR [pcPay_EIG_TestMode]
END


END
GO
/****** Object:  Default [DF__pcPay_Cyb__pcPay__7834CCDD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Cyb__pcPay__7834CCDD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Cyb__pcPay__7834CCDD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_CyberSource] ADD  CONSTRAINT [DF__pcPay_Cyb__pcPay__7834CCDD]  DEFAULT ((0)) FOR [pcPay_Cys_Id]
END


END
GO
/****** Object:  Default [DF__pcPay_Cyb__pcPay__7928F116]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Cyb__pcPay__7928F116]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Cyb__pcPay__7928F116]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_CyberSource] ADD  CONSTRAINT [DF__pcPay_Cyb__pcPay__7928F116]  DEFAULT ((0)) FOR [pcPay_Cys_TransType]
END


END
GO
/****** Object:  Default [DF__pcPay_Cyb__pcPay__7A1D154F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Cyb__pcPay__7A1D154F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Cyb__pcPay__7A1D154F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_CyberSource] ADD  CONSTRAINT [DF__pcPay_Cyb__pcPay__7A1D154F]  DEFAULT ((0)) FOR [pcPay_Cys_CVV]
END


END
GO
/****** Object:  Default [DF__pcPay_Cyb__pcPay__7B113988]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Cyb__pcPay__7B113988]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Cyb__pcPay__7B113988]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_CyberSource] ADD  CONSTRAINT [DF__pcPay_Cyb__pcPay__7B113988]  DEFAULT ((0)) FOR [pcPay_Cys_TestMode]
END


END
GO
/****** Object:  Default [DBX_pcPay_Cys_eCheck_20179]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcPay_Cys_eCheck_20179]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcPay_Cys_eCheck_20179]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_CyberSource] ADD  CONSTRAINT [DBX_pcPay_Cys_eCheck_20179]  DEFAULT ((0)) FOR [pcPay_Cys_eCheck]
END


END
GO
/****** Object:  Default [DBX_pcPay_Cys_eCheckPending_22226]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcPay_Cys_eCheckPending_22226]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcPay_Cys_eCheckPending_22226]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_CyberSource] ADD  CONSTRAINT [DBX_pcPay_Cys_eCheckPending_22226]  DEFAULT ((0)) FOR [pcPay_Cys_eCheckPending]
END


END
GO
/****** Object:  Default [DF__pcPay_Cen__pcPay__7EE1CA6C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Cen__pcPay__7EE1CA6C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Cen__pcPay__7EE1CA6C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_Centinel_Orders] ADD  CONSTRAINT [DF__pcPay_Cen__pcPay__7EE1CA6C]  DEFAULT ((0)) FOR [pcPay_CentOrd_OrderID]
END


END
GO
/****** Object:  Default [DF__pcPay_Cen__pcPay__02B25B50]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Cen__pcPay__02B25B50]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Cen__pcPay__02B25B50]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_Centinel] ADD  CONSTRAINT [DF__pcPay_Cen__pcPay__02B25B50]  DEFAULT ((0)) FOR [pcPay_Cent_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_Cen__pcPay__03A67F89]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_Cen__pcPay__03A67F89]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_Cen__pcPay__03A67F89]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_Centinel] ADD  CONSTRAINT [DF__pcPay_Cen__pcPay__03A67F89]  DEFAULT ((0)) FOR [pcPay_Cent_Active]
END


END
GO
/****** Object:  Default [DF__pcPay_CBN__pcPay__0777106D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_CBN__pcPay__0777106D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_CBN__pcPay__0777106D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_CBN] ADD  CONSTRAINT [DF__pcPay_CBN__pcPay__0777106D]  DEFAULT ((0)) FOR [pcPay_CBN_id]
END


END
GO
/****** Object:  Default [DF__pcPay_CBN__pcPay__086B34A6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_CBN__pcPay__086B34A6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_CBN__pcPay__086B34A6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_CBN] ADD  CONSTRAINT [DF__pcPay_CBN__pcPay__086B34A6]  DEFAULT ((0)) FOR [pcPay_CBN_test]
END


END
GO
/****** Object:  Default [DF__pcPay_CBN__pcPay__095F58DF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_CBN__pcPay__095F58DF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_CBN__pcPay__095F58DF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_CBN] ADD  CONSTRAINT [DF__pcPay_CBN__pcPay__095F58DF]  DEFAULT ((0)) FOR [pcPay_CBN_status]
END


END
GO
/****** Object:  Default [DF__pcPay_ACH__pcPay__51DA19CB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_ACH__pcPay__51DA19CB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_ACH__pcPay__51DA19CB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_ACHDirect] ADD  CONSTRAINT [DF__pcPay_ACH__pcPay__51DA19CB]  DEFAULT ((0)) FOR [pcPay_ACH_ID]
END


END
GO
/****** Object:  Default [DF__pcPay_ACH__pcPay__52CE3E04]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_ACH__pcPay__52CE3E04]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_ACH__pcPay__52CE3E04]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_ACHDirect] ADD  CONSTRAINT [DF__pcPay_ACH__pcPay__52CE3E04]  DEFAULT ((0)) FOR [pcPay_ACH_TestMode]
END


END
GO
/****** Object:  Default [DF__pcPay_ACH__pcPay__53C2623D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPay_ACH__pcPay__53C2623D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPay_ACH__pcPay__53C2623D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPay_ACHDirect] ADD  CONSTRAINT [DF__pcPay_ACH__pcPay__53C2623D]  DEFAULT ((0)) FOR [pcPay_ACH_CVV]
END


END
GO
/****** Object:  Default [DF__pcPackage__idOrd__13DCE752]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__idOrd__13DCE752]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__idOrd__13DCE752]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__idOrd__13DCE752]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__pcPackage__pcPac__14D10B8B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__pcPac__14D10B8B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__pcPac__14D10B8B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__pcPac__14D10B8B]  DEFAULT ((0)) FOR [pcPackageInfo_PackageNumber]
END


END
GO
/****** Object:  Default [DF__pcPackage__pcPac__15C52FC4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__pcPac__15C52FC4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__pcPac__15C52FC4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__pcPac__15C52FC4]  DEFAULT ((0)) FOR [pcPackageInfo_PackageWeight]
END


END
GO
/****** Object:  Default [DF__pcPackage__pcPac__16B953FD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__pcPac__16B953FD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__pcPac__16B953FD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__pcPac__16B953FD]  DEFAULT ((0)) FOR [pcPackageInfo_ShipToResidential]
END


END
GO
/****** Object:  Default [DF__pcPackage__pcPac__17AD7836]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__pcPac__17AD7836]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__pcPac__17AD7836]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__pcPac__17AD7836]  DEFAULT ((0)) FOR [pcPackageInfo_AddSaturdayDelivery]
END


END
GO
/****** Object:  Default [DF__pcPackage__pcPac__18A19C6F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__pcPac__18A19C6F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__pcPac__18A19C6F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__pcPac__18A19C6F]  DEFAULT ((0)) FOR [pcPackageInfo_AddVerbalConfirmation]
END


END
GO
/****** Object:  Default [DF__pcPackage__pcPac__1995C0A8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__pcPac__1995C0A8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__pcPac__1995C0A8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__pcPac__1995C0A8]  DEFAULT ((0)) FOR [pcPackageInfo_AddAdditionalHandling]
END


END
GO
/****** Object:  Default [DF__pcPackage__pcPac__1A89E4E1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__pcPac__1A89E4E1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__pcPac__1A89E4E1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__pcPac__1A89E4E1]  DEFAULT ((0)) FOR [pcPackageInfo_OverSizedIndicator]
END


END
GO
/****** Object:  Default [DF__pcPackage__pcPac__1B7E091A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__pcPac__1B7E091A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__pcPac__1B7E091A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__pcPac__1B7E091A]  DEFAULT ((0)) FOR [pcPackageInfo_UPSCODFunds]
END


END
GO
/****** Object:  Default [DF__pcPackage__pcPac__1C722D53]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcPackage__pcPac__1C722D53]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcPackage__pcPac__1C722D53]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF__pcPackage__pcPac__1C722D53]  DEFAULT ((0)) FOR [pcPackageInfo_Status]
END


END
GO
/****** Object:  Default [DF_pcPackageInfo_pcPackageInfo_MethodFlag]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPackageInfo_pcPackageInfo_MethodFlag]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPackageInfo_pcPackageInfo_MethodFlag]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF_pcPackageInfo_pcPackageInfo_MethodFlag]  DEFAULT ((0)) FOR [pcPackageInfo_MethodFlag]
END


END
GO
/****** Object:  Default [DF_pcPackageInfo_pcPackageInfo_Endicia]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPackageInfo_pcPackageInfo_Endicia]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPackageInfo_pcPackageInfo_Endicia]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF_pcPackageInfo_pcPackageInfo_Endicia]  DEFAULT ((0)) FOR [pcPackageInfo_Endicia]
END


END
GO
/****** Object:  Default [DF_pcPackageInfo_pcPackageInfo_EndiciaIsPIC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcPackageInfo_pcPackageInfo_EndiciaIsPIC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcPackageInfo_pcPackageInfo_EndiciaIsPIC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcPackageInfo] ADD  CONSTRAINT [DF_pcPackageInfo_pcPackageInfo_EndiciaIsPIC]  DEFAULT ((0)) FOR [pcPackageInfo_EndiciaIsPIC]
END


END
GO
/****** Object:  Default [DF__pcNewArri__pcNAS__5C8CB268]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcNewArri__pcNAS__5C8CB268]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcNewArri__pcNAS__5C8CB268]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcNewArrivalsSettings] ADD  CONSTRAINT [DF__pcNewArri__pcNAS__5C8CB268]  DEFAULT ((0)) FOR [pcNAS_NewArrCount]
END


END
GO
/****** Object:  Default [DF__pcNewArri__pcNAS__5D80D6A1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcNewArri__pcNAS__5D80D6A1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcNewArri__pcNAS__5D80D6A1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcNewArrivalsSettings] ADD  CONSTRAINT [DF__pcNewArri__pcNAS__5D80D6A1]  DEFAULT ((0)) FOR [pcNAS_NDays]
END


END
GO
/****** Object:  Default [DF__pcNewArri__pcNAS__5E74FADA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcNewArri__pcNAS__5E74FADA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcNewArri__pcNAS__5E74FADA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcNewArrivalsSettings] ADD  CONSTRAINT [DF__pcNewArri__pcNAS__5E74FADA]  DEFAULT ((0)) FOR [pcNAS_NotForSale]
END


END
GO
/****** Object:  Default [DF__pcNewArri__pcNAS__5F691F13]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcNewArri__pcNAS__5F691F13]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcNewArri__pcNAS__5F691F13]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcNewArrivalsSettings] ADD  CONSTRAINT [DF__pcNewArri__pcNAS__5F691F13]  DEFAULT ((0)) FOR [pcNAS_OutOfStock]
END


END
GO
/****** Object:  Default [DF__pcNewArri__pcNAS__605D434C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcNewArri__pcNAS__605D434C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcNewArri__pcNAS__605D434C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcNewArrivalsSettings] ADD  CONSTRAINT [DF__pcNewArri__pcNAS__605D434C]  DEFAULT ((0)) FOR [pcNAS_SKU]
END


END
GO
/****** Object:  Default [DF__pcNewArri__pcNAS__61516785]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcNewArri__pcNAS__61516785]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcNewArri__pcNAS__61516785]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcNewArrivalsSettings] ADD  CONSTRAINT [DF__pcNewArri__pcNAS__61516785]  DEFAULT ((0)) FOR [pcNAS_ShowImg]
END


END
GO
/****** Object:  Default [DF__pcMailUpS__idCus__002AF460]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpS__idCus__002AF460]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpS__idCus__002AF460]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpSubs] ADD  CONSTRAINT [DF__pcMailUpS__idCus__002AF460]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcMailUpS__pcMai__011F1899]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpS__pcMai__011F1899]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpS__pcMai__011F1899]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpSubs] ADD  CONSTRAINT [DF__pcMailUpS__pcMai__011F1899]  DEFAULT ((0)) FOR [pcMailUpLists_ID]
END


END
GO
/****** Object:  Default [DF__pcMailUpS__pcMai__02133CD2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpS__pcMai__02133CD2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpS__pcMai__02133CD2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpSubs] ADD  CONSTRAINT [DF__pcMailUpS__pcMai__02133CD2]  DEFAULT ((0)) FOR [pcMailUpSubs_SyncNeeded]
END


END
GO
/****** Object:  Default [DF__pcMailUpS__pcMai__0307610B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpS__pcMai__0307610B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpS__pcMai__0307610B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpSubs] ADD  CONSTRAINT [DF__pcMailUpS__pcMai__0307610B]  DEFAULT ((0)) FOR [pcMailUpSubs_Optout]
END


END
GO
/****** Object:  Default [DF__pcMailUpS__pcMai__6D181FEC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpS__pcMai__6D181FEC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpS__pcMai__6D181FEC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpSettings] ADD  CONSTRAINT [DF__pcMailUpS__pcMai__6D181FEC]  DEFAULT ((0)) FOR [pcMailUpSett_AutoReg]
END


END
GO
/****** Object:  Default [DF__pcMailUpS__pcMai__6E0C4425]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpS__pcMai__6E0C4425]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpS__pcMai__6E0C4425]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpSettings] ADD  CONSTRAINT [DF__pcMailUpS__pcMai__6E0C4425]  DEFAULT ((0)) FOR [pcMailUpSett_RegSuccess]
END


END
GO
/****** Object:  Default [DF__pcMailUpS__pcMai__6F00685E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpS__pcMai__6F00685E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpS__pcMai__6F00685E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpSettings] ADD  CONSTRAINT [DF__pcMailUpS__pcMai__6F00685E]  DEFAULT ((0)) FOR [pcMailUpSett_BulkRegister]
END


END
GO
/****** Object:  Default [DF_pcMailUpSettings_pcMailUpSett_TurnOff]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcMailUpSettings_pcMailUpSett_TurnOff]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcMailUpSettings_pcMailUpSett_TurnOff]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpSettings] ADD  CONSTRAINT [DF_pcMailUpSettings_pcMailUpSett_TurnOff]  DEFAULT ((0)) FOR [pcMailUpSett_TurnOff]
END


END
GO
/****** Object:  Default [DF__pcMailUpL__pcMai__73C51D7B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpL__pcMai__73C51D7B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpL__pcMai__73C51D7B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpLists] ADD  CONSTRAINT [DF__pcMailUpL__pcMai__73C51D7B]  DEFAULT ((0)) FOR [pcMailUpLists_ListID]
END


END
GO
/****** Object:  Default [DF__pcMailUpL__pcMai__74B941B4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpL__pcMai__74B941B4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpL__pcMai__74B941B4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpLists] ADD  CONSTRAINT [DF__pcMailUpL__pcMai__74B941B4]  DEFAULT ((0)) FOR [pcMailUpLists_Active]
END


END
GO
/****** Object:  Default [DF__pcMailUpL__pcMai__75AD65ED]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpL__pcMai__75AD65ED]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpL__pcMai__75AD65ED]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpLists] ADD  CONSTRAINT [DF__pcMailUpL__pcMai__75AD65ED]  DEFAULT ((0)) FOR [pcMailUpLists_Removed]
END


END
GO
/****** Object:  Default [DF__pcMailUpG__pcMai__7A721B0A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpG__pcMai__7A721B0A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpG__pcMai__7A721B0A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpGroups] ADD  CONSTRAINT [DF__pcMailUpG__pcMai__7A721B0A]  DEFAULT ((0)) FOR [pcMailUpLists_ID]
END


END
GO
/****** Object:  Default [DF__pcMailUpG__pcMai__7B663F43]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcMailUpG__pcMai__7B663F43]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcMailUpG__pcMai__7B663F43]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcMailUpGroups] ADD  CONSTRAINT [DF__pcMailUpG__pcMai__7B663F43]  DEFAULT ((0)) FOR [pcMailUpGroups_GroupID]
END


END
GO
/****** Object:  Default [DF__pcImageDi__pcImg__7387885E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcImageDi__pcImg__7387885E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcImageDi__pcImg__7387885E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcImageDirectory] ADD  CONSTRAINT [DF__pcImageDi__pcImg__7387885E]  DEFAULT ((0)) FOR [pcImgDir_Size]
END


END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__25077354]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__25077354]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__25077354]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__25077354]  DEFAULT ((0)) FOR [pcHPS_FeaturedCount]
END


END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__25FB978D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__25FB978D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__25FB978D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__25FB978D]  DEFAULT ((0)) FOR [pcHPS_First]
END


END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__26EFBBC6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__26EFBBC6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__26EFBBC6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__26EFBBC6]  DEFAULT ((0)) FOR [pcHPS_ShowSKU]
END


END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__27E3DFFF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__27E3DFFF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__27E3DFFF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__27E3DFFF]  DEFAULT ((0)) FOR [pcHPS_ShowImg]
END


END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__28D80438]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__28D80438]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__28D80438]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__28D80438]  DEFAULT ((0)) FOR [pcHPS_SpcCount]
END


END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__29CC2871]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__29CC2871]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__29CC2871]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__29CC2871]  DEFAULT ((0)) FOR [pcHPS_SpcOrder]
END


END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__2AC04CAA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__2AC04CAA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__2AC04CAA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__2AC04CAA]  DEFAULT ((0)) FOR [pcHPS_NewCount]
END


END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__2BB470E3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__2BB470E3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__2BB470E3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__2BB470E3]  DEFAULT ((0)) FOR [pcHPS_NewOrder]
END



END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__2CA8951C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__2CA8951C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__2CA8951C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__2CA8951C]  DEFAULT ((0)) FOR [pcHPS_BestCount]
END


END
GO
/****** Object:  Default [DF__pcHomePag__pcHPS__2D9CB955]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcHomePag__pcHPS__2D9CB955]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcHomePag__pcHPS__2D9CB955]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcHomePageSettings] ADD  CONSTRAINT [DF__pcHomePag__pcHPS__2D9CB955]  DEFAULT ((0)) FOR [pcHPS_BestOrder]
END


END
GO
/****** Object:  Default [DF__pcGWSetti__pcGWS__32616E72]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGWSetti__pcGWS__32616E72]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGWSetti__pcGWS__32616E72]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGWSettings] ADD  CONSTRAINT [DF__pcGWSetti__pcGWS__32616E72]  DEFAULT ((0)) FOR [pcGWSet_Show]
END


END
GO
/****** Object:  Default [DF__pcGWSetti__pcGWS__335592AB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGWSetti__pcGWS__335592AB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGWSetti__pcGWS__335592AB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGWSettings] ADD  CONSTRAINT [DF__pcGWSetti__pcGWS__335592AB]  DEFAULT ((0)) FOR [pcGWSet_Overview]
END


END
GO
/****** Object:  Default [DF_pcGWSettings_pcGWSet_OverviewCart]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcGWSettings_pcGWSet_OverviewCart]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcGWSettings_pcGWSet_OverviewCart]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGWSettings] ADD  CONSTRAINT [DF_pcGWSettings_pcGWSet_OverviewCart]  DEFAULT ((0)) FOR [pcGWSet_OverviewCart]
END


END
GO
/****** Object:  Default [DF__pcGWOptio__pcGW___3726238F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGWOptio__pcGW___3726238F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGWOptio__pcGW___3726238F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGWOptions] ADD  CONSTRAINT [DF__pcGWOptio__pcGW___3726238F]  DEFAULT ((0)) FOR [pcGW_OptPrice]
END


END
GO
/****** Object:  Default [DF__pcGWOptio__pcGW___381A47C8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGWOptio__pcGW___381A47C8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGWOptio__pcGW___381A47C8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGWOptions] ADD  CONSTRAINT [DF__pcGWOptio__pcGW___381A47C8]  DEFAULT ((0)) FOR [pcGW_Removed]
END


END
GO
/****** Object:  Default [DF_pcGWOptions_pcGW_OptActive]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcGWOptions_pcGW_OptActive]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcGWOptions_pcGW_OptActive]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGWOptions] ADD  CONSTRAINT [DF_pcGWOptions_pcGW_OptActive]  DEFAULT ((1)) FOR [pcGW_OptActive]
END


END
GO
/****** Object:  Default [DF_pcGWOptions_pcGW_OptOrder]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcGWOptions_pcGW_OptOrder]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcGWOptions_pcGW_OptOrder]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGWOptions] ADD  CONSTRAINT [DF_pcGWOptions_pcGW_OptOrder]  DEFAULT ((0)) FOR [pcGW_OptOrder]
END


END
GO
/****** Object:  Default [DF__pcGCOrder__pcGO___3BEAD8AC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGCOrder__pcGO___3BEAD8AC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGCOrder__pcGO___3BEAD8AC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGCOrdered] ADD  CONSTRAINT [DF__pcGCOrder__pcGO___3BEAD8AC]  DEFAULT ((0)) FOR [pcGO_IDProduct]
END


END
GO
/****** Object:  Default [DF__pcGCOrder__pcGO___3CDEFCE5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGCOrder__pcGO___3CDEFCE5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGCOrder__pcGO___3CDEFCE5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGCOrdered] ADD  CONSTRAINT [DF__pcGCOrder__pcGO___3CDEFCE5]  DEFAULT ((0)) FOR [pcGO_IDOrder]
END


END
GO
/****** Object:  Default [DF__pcGCOrder__pcGO___3DD3211E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGCOrder__pcGO___3DD3211E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGCOrder__pcGO___3DD3211E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGCOrdered] ADD  CONSTRAINT [DF__pcGCOrder__pcGO___3DD3211E]  DEFAULT ((0)) FOR [pcGO_Amount]
END


END
GO
/****** Object:  Default [DF__pcGCOrder__pcGO___3EC74557]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGCOrder__pcGO___3EC74557]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGCOrder__pcGO___3EC74557]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGCOrdered] ADD  CONSTRAINT [DF__pcGCOrder__pcGO___3EC74557]  DEFAULT ((0)) FOR [pcGO_Status]
END


END
GO
/****** Object:  Default [DF__pcGC__pcGC_IDPro__4297D63B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGC__pcGC_IDPro__4297D63B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGC__pcGC_IDPro__4297D63B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGC] ADD  CONSTRAINT [DF__pcGC__pcGC_IDPro__4297D63B]  DEFAULT ((0)) FOR [pcGC_IDProduct]
END


END
GO
/****** Object:  Default [DF__pcGC__pcGC_Exp__438BFA74]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGC__pcGC_Exp__438BFA74]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGC__pcGC_Exp__438BFA74]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGC] ADD  CONSTRAINT [DF__pcGC__pcGC_Exp__438BFA74]  DEFAULT ((0)) FOR [pcGC_Exp]
END


END
GO
/****** Object:  Default [DF__pcGC__pcGC_ExpDa__44801EAD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGC__pcGC_ExpDa__44801EAD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGC__pcGC_ExpDa__44801EAD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGC] ADD  CONSTRAINT [DF__pcGC__pcGC_ExpDa__44801EAD]  DEFAULT ((0)) FOR [pcGC_ExpDays]
END


END
GO
/****** Object:  Default [DF__pcGC__pcGC_EOnly__457442E6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGC__pcGC_EOnly__457442E6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGC__pcGC_EOnly__457442E6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGC] ADD  CONSTRAINT [DF__pcGC__pcGC_EOnly__457442E6]  DEFAULT ((0)) FOR [pcGC_EOnly]
END


END
GO
/****** Object:  Default [DF__pcGC__pcGC_CodeG__4668671F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcGC__pcGC_CodeG__4668671F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcGC__pcGC_CodeG__4668671F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcGC] ADD  CONSTRAINT [DF__pcGC__pcGC_CodeG__4668671F]  DEFAULT ((0)) FOR [pcGC_CodeGen]
END


END
GO
/****** Object:  Default [DF__pcFTypes__pcFTyp__4A38F803]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcFTypes__pcFTyp__4A38F803]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcFTypes__pcFTyp__4A38F803]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcFTypes] ADD  CONSTRAINT [DF__pcFTypes__pcFTyp__4A38F803]  DEFAULT ((0)) FOR [pcFType_ShowImg]
END


END
GO
/****** Object:  Default [DF__pcFStatus__pcFSt__4E0988E7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcFStatus__pcFSt__4E0988E7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcFStatus__pcFSt__4E0988E7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcFStatus] ADD  CONSTRAINT [DF__pcFStatus__pcFSt__4E0988E7]  DEFAULT ((0)) FOR [pcFStat_ShowImg]
END


END
GO
/****** Object:  Default [DF__pcExportG__idpro__64FBD3EA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcExportG__idpro__64FBD3EA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcExportG__idpro__64FBD3EA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcExportGoogle] ADD  CONSTRAINT [DF__pcExportG__idpro__64FBD3EA]  DEFAULT ((0)) FOR [idproduct]
END


END
GO
/****** Object:  Default [DF__pcExportC__idpro__39AD8A7F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcExportC__idpro__39AD8A7F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcExportC__idpro__39AD8A7F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcExportCashback] ADD  CONSTRAINT [DF__pcExportC__idpro__39AD8A7F]  DEFAULT ((0)) FOR [idproduct]
END


END
GO
/****** Object:  Default [DF__pcEvProdu__pcEP___62E4AA3C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvProdu__pcEP___62E4AA3C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvProdu__pcEP___62E4AA3C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvProducts] ADD  CONSTRAINT [DF__pcEvProdu__pcEP___62E4AA3C]  DEFAULT ((0)) FOR [pcEP_IDEvent]
END


END
GO
/****** Object:  Default [DF__pcEvProdu__pcEP___63D8CE75]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvProdu__pcEP___63D8CE75]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvProdu__pcEP___63D8CE75]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvProducts] ADD  CONSTRAINT [DF__pcEvProdu__pcEP___63D8CE75]  DEFAULT ((0)) FOR [pcEP_IDProduct]
END


END
GO
/****** Object:  Default [DF__pcEvProdu__pcEP___64CCF2AE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvProdu__pcEP___64CCF2AE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvProdu__pcEP___64CCF2AE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvProducts] ADD  CONSTRAINT [DF__pcEvProdu__pcEP___64CCF2AE]  DEFAULT ((0)) FOR [pcEP_Qty]
END


END
GO
/****** Object:  Default [DF__pcEvProdu__pcEP___65C116E7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvProdu__pcEP___65C116E7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvProdu__pcEP___65C116E7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvProducts] ADD  CONSTRAINT [DF__pcEvProdu__pcEP___65C116E7]  DEFAULT ((0)) FOR [pcEP_HQty]
END


END
GO
/****** Object:  Default [DF__pcEvProdu__pcEP___66B53B20]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvProdu__pcEP___66B53B20]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvProdu__pcEP___66B53B20]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvProducts] ADD  CONSTRAINT [DF__pcEvProdu__pcEP___66B53B20]  DEFAULT ((0)) FOR [pcEP_IDOptionA]
END


END
GO
/****** Object:  Default [DF__pcEvProdu__pcEP___67A95F59]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvProdu__pcEP___67A95F59]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvProdu__pcEP___67A95F59]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvProducts] ADD  CONSTRAINT [DF__pcEvProdu__pcEP___67A95F59]  DEFAULT ((0)) FOR [pcEP_IDOptionB]
END


END
GO
/****** Object:  Default [DF__pcEvProdu__pcEP___689D8392]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvProdu__pcEP___689D8392]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvProdu__pcEP___689D8392]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvProducts] ADD  CONSTRAINT [DF__pcEvProdu__pcEP___689D8392]  DEFAULT ((0)) FOR [pcEP_GC]
END


END
GO
/****** Object:  Default [DF__pcEvProdu__pcEP___6991A7CB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvProdu__pcEP___6991A7CB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvProdu__pcEP___6991A7CB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvProducts] ADD  CONSTRAINT [DF__pcEvProdu__pcEP___6991A7CB]  DEFAULT ((0)) FOR [pcEP_IDConfig]
END


END
GO
/****** Object:  Default [DF__pcEvProdu__pcEP___6A85CC04]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvProdu__pcEP___6A85CC04]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvProdu__pcEP___6A85CC04]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvProducts] ADD  CONSTRAINT [DF__pcEvProdu__pcEP___6A85CC04]  DEFAULT ((0)) FOR [pcEP_Price]
END


END
GO
/****** Object:  Default [DF__pcEvents__pcEv_I__787EE5A0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvents__pcEv_I__787EE5A0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvents__pcEv_I__787EE5A0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvents] ADD  CONSTRAINT [DF__pcEvents__pcEv_I__787EE5A0]  DEFAULT ((0)) FOR [pcEv_IDCustomer]
END


END
GO
/****** Object:  Default [DF__pcEvents__pcEv_D__797309D9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvents__pcEv_D__797309D9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvents__pcEv_D__797309D9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvents] ADD  CONSTRAINT [DF__pcEvents__pcEv_D__797309D9]  DEFAULT ((0)) FOR [pcEv_Delivery]
END


END
GO
/****** Object:  Default [DF__pcEvents__pcEv_M__7A672E12]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvents__pcEv_M__7A672E12]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvents__pcEv_M__7A672E12]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvents] ADD  CONSTRAINT [DF__pcEvents__pcEv_M__7A672E12]  DEFAULT ((0)) FOR [pcEv_MyAddr]
END


END
GO
/****** Object:  Default [DF__pcEvents__pcEv_H__7B5B524B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvents__pcEv_H__7B5B524B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvents__pcEv_H__7B5B524B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvents] ADD  CONSTRAINT [DF__pcEvents__pcEv_H__7B5B524B]  DEFAULT ((0)) FOR [pcEv_Hide]
END


END
GO
/****** Object:  Default [DF__pcEvents__pcEv_N__7C4F7684]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvents__pcEv_N__7C4F7684]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvents__pcEv_N__7C4F7684]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvents] ADD  CONSTRAINT [DF__pcEvents__pcEv_N__7C4F7684]  DEFAULT ((0)) FOR [pcEv_Notify]
END


END
GO
/****** Object:  Default [DF__pcEvents__pcEv_I__7D439ABD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvents__pcEv_I__7D439ABD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvents__pcEv_I__7D439ABD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvents] ADD  CONSTRAINT [DF__pcEvents__pcEv_I__7D439ABD]  DEFAULT ((0)) FOR [pcEv_IncGcs]
END


END
GO
/****** Object:  Default [DF__pcEvents__pcEv_A__7E37BEF6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEvents__pcEv_A__7E37BEF6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEvents__pcEv_A__7E37BEF6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvents] ADD  CONSTRAINT [DF__pcEvents__pcEv_A__7E37BEF6]  DEFAULT ((0)) FOR [pcEv_Active]
END


END
GO
/****** Object:  Default [DF_pcEvents_pcEv_HideAddress]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcEvents_pcEv_HideAddress]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcEvents_pcEv_HideAddress]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEvents] ADD  CONSTRAINT [DF_pcEvents_pcEv_HideAddress]  DEFAULT ((0)) FOR [pcEv_HideAddress]
END


END
GO
/****** Object:  Default [DF__pcEDCTran__IDOrd__4CE05A84]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCTran__IDOrd__4CE05A84]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCTran__IDOrd__4CE05A84]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCTrans] ADD  CONSTRAINT [DF__pcEDCTran__IDOrd__4CE05A84]  DEFAULT ((0)) FOR [IDOrder]
END


END
GO
/****** Object:  Default [DF__pcEDCTran__pcPac__4DD47EBD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCTran__pcPac__4DD47EBD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCTran__pcPac__4DD47EBD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCTrans] ADD  CONSTRAINT [DF__pcEDCTran__pcPac__4DD47EBD]  DEFAULT ((0)) FOR [pcPackageInfo_ID]
END


END
GO
/****** Object:  Default [DF__pcEDCTran__pcET___4EC8A2F6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCTran__pcET___4EC8A2F6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCTran__pcET___4EC8A2F6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCTrans] ADD  CONSTRAINT [DF__pcEDCTran__pcET___4EC8A2F6]  DEFAULT ((0)) FOR [pcET_Postage]
END


END
GO
/****** Object:  Default [DF__pcEDCTran__pcET___4FBCC72F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCTran__pcET___4FBCC72F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCTran__pcET___4FBCC72F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCTrans] ADD  CONSTRAINT [DF__pcEDCTran__pcET___4FBCC72F]  DEFAULT ((0)) FOR [pcET_RefundID]
END


END
GO
/****** Object:  Default [DF__pcEDCTran__pcET___50B0EB68]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCTran__pcET___50B0EB68]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCTran__pcET___50B0EB68]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCTrans] ADD  CONSTRAINT [DF__pcEDCTran__pcET___50B0EB68]  DEFAULT ((0)) FOR [pcET_Method]
END


END
GO
/****** Object:  Default [DF__pcEDCTran__pcET___51A50FA1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCTran__pcET___51A50FA1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCTran__pcET___51A50FA1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCTrans] ADD  CONSTRAINT [DF__pcEDCTran__pcET___51A50FA1]  DEFAULT ((0)) FOR [pcET_Success]
END


END
GO
/****** Object:  Default [DF__pcEDCTran__pcET___529933DA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCTran__pcET___529933DA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCTran__pcET___529933DA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCTrans] ADD  CONSTRAINT [DF__pcEDCTran__pcET___529933DA]  DEFAULT ((0)) FOR [pcET_Fees]
END


END
GO
/****** Object:  Default [DF__pcEDCTran__pcET___538D5813]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCTran__pcET___538D5813]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCTran__pcET___538D5813]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCTrans] ADD  CONSTRAINT [DF__pcEDCTran__pcET___538D5813]  DEFAULT ((0)) FOR [pcET_subPostage]
END


END
GO
/****** Object:  Default [DF__pcEDCSett__pcES___4356F04A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCSett__pcES___4356F04A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCSett__pcES___4356F04A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCSettings] ADD  CONSTRAINT [DF__pcEDCSett__pcES___4356F04A]  DEFAULT ((0)) FOR [pcES_UserID]
END


END
GO
/****** Object:  Default [DF__pcEDCSett__pcES___444B1483]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCSett__pcES___444B1483]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCSett__pcES___444B1483]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCSettings] ADD  CONSTRAINT [DF__pcEDCSett__pcES___444B1483]  DEFAULT ((0)) FOR [pcES_AutoRefill]
END


END
GO
/****** Object:  Default [DF__pcEDCSett__pcES___453F38BC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCSett__pcES___453F38BC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCSett__pcES___453F38BC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCSettings] ADD  CONSTRAINT [DF__pcEDCSett__pcES___453F38BC]  DEFAULT ((0)) FOR [pcES_TriggerAmount]
END


END
GO
/****** Object:  Default [DF__pcEDCSett__pcES___46335CF5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCSett__pcES___46335CF5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCSett__pcES___46335CF5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCSettings] ADD  CONSTRAINT [DF__pcEDCSett__pcES___46335CF5]  DEFAULT ((0)) FOR [pcES_FillAmount]
END


END
GO
/****** Object:  Default [DF__pcEDCSett__pcES___4727812E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCSett__pcES___4727812E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCSett__pcES___4727812E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCSettings] ADD  CONSTRAINT [DF__pcEDCSett__pcES___4727812E]  DEFAULT ((0)) FOR [pcES_LogTrans]
END


END
GO
/****** Object:  Default [DF__pcEDCSett__pcES___481BA567]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCSett__pcES___481BA567]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCSett__pcES___481BA567]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCSettings] ADD  CONSTRAINT [DF__pcEDCSett__pcES___481BA567]  DEFAULT ((0)) FOR [pcES_Reg]
END


END
GO
/****** Object:  Default [DF__pcEDCSett__pcES___490FC9A0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCSett__pcES___490FC9A0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCSett__pcES___490FC9A0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCSettings] ADD  CONSTRAINT [DF__pcEDCSett__pcES___490FC9A0]  DEFAULT ((0)) FOR [pcES_TestMode]
END


END
GO
/****** Object:  Default [DF_pcEDCSettings_pcES_AutoRmvLogs]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcEDCSettings_pcES_AutoRmvLogs]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcEDCSettings_pcES_AutoRmvLogs]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCSettings] ADD  CONSTRAINT [DF_pcEDCSettings_pcES_AutoRmvLogs]  DEFAULT ((0)) FOR [pcES_AutoRmvLogs]
END


END
GO
/****** Object:  Default [DF__pcEDCLogs__pcET___4AF81212]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcEDCLogs__pcET___4AF81212]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcEDCLogs__pcET___4AF81212]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcEDCLogs] ADD  CONSTRAINT [DF__pcEDCLogs__pcET___4AF81212]  DEFAULT ((0)) FOR [pcET_ID]
END


END
GO
/****** Object:  Default [DF__pcDropShi__idPro__160F4887]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDropShi__idPro__160F4887]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDropShi__idPro__160F4887]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDropShippersSuppliers] ADD  CONSTRAINT [DF__pcDropShi__idPro__160F4887]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__pcDropShi__pcDS___17036CC0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDropShi__pcDS___17036CC0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDropShi__pcDS___17036CC0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDropShippersSuppliers] ADD  CONSTRAINT [DF__pcDropShi__pcDS___17036CC0]  DEFAULT ((0)) FOR [pcDS_IsDropShipper]
END


END
GO
/****** Object:  Default [DF__pcDropShi__pcDro__3F865F66]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDropShi__pcDro__3F865F66]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDropShi__pcDro__3F865F66]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDropShippersOrders] ADD  CONSTRAINT [DF__pcDropShi__pcDro__3F865F66]  DEFAULT ((0)) FOR [pcDropShipO_DropShipper_ID]
END


END
GO
/****** Object:  Default [DF__pcDropShi__pcDro__407A839F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDropShi__pcDro__407A839F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDropShi__pcDro__407A839F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDropShippersOrders] ADD  CONSTRAINT [DF__pcDropShi__pcDro__407A839F]  DEFAULT ((0)) FOR [pcDropShipO_idOrder]
END


END
GO
/****** Object:  Default [DF__pcDropShi__pcDro__416EA7D8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDropShi__pcDro__416EA7D8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDropShi__pcDro__416EA7D8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDropShippersOrders] ADD  CONSTRAINT [DF__pcDropShi__pcDro__416EA7D8]  DEFAULT ((0)) FOR [pcDropShipO_OrderStatus]
END


END
GO
/****** Object:  Default [DF__pcDropshi__pcDro__1AD3FDA4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDropshi__pcDro__1AD3FDA4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDropshi__pcDro__1AD3FDA4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDropshippers] ADD  CONSTRAINT [DF__pcDropshi__pcDro__1AD3FDA4]  DEFAULT ((0)) FOR [pcDropShipper_NoticeType]
END


END
GO
/****** Object:  Default [DF__pcDropshi__pcDro__1BC821DD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDropshi__pcDro__1BC821DD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDropshi__pcDro__1BC821DD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDropshippers] ADD  CONSTRAINT [DF__pcDropshi__pcDro__1BC821DD]  DEFAULT ((0)) FOR [pcDropShipper_NotifyManually]
END


END
GO
/****** Object:  Default [DF__pcDropshi__pcDro__1CBC4616]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDropshi__pcDro__1CBC4616]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDropshi__pcDro__1CBC4616]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDropshippers] ADD  CONSTRAINT [DF__pcDropshi__pcDro__1CBC4616]  DEFAULT ((0)) FOR [pcDropShipper_CustNotifyUpdates]
END


END
GO
/****** Object:  Default [DF__pcDFShip__pcFShi__208CD6FA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFShip__pcFShi__208CD6FA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFShip__pcFShi__208CD6FA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFShip] ADD  CONSTRAINT [DF__pcDFShip__pcFShi__208CD6FA]  DEFAULT ((0)) FOR [pcFShip_IDDiscount]
END


END
GO
/****** Object:  Default [DF__pcDFShip__pcFShi__2180FB33]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFShip__pcFShi__2180FB33]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFShip__pcFShi__2180FB33]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFShip] ADD  CONSTRAINT [DF__pcDFShip__pcFShi__2180FB33]  DEFAULT ((0)) FOR [pcFShip_IDShipOpt]
END


END
GO
/****** Object:  Default [DF__pcDFProds__pcFPr__25518C17]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFProds__pcFPr__25518C17]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFProds__pcFPr__25518C17]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFProds] ADD  CONSTRAINT [DF__pcDFProds__pcFPr__25518C17]  DEFAULT ((0)) FOR [pcFPro_IDDiscount]
END


END
GO
/****** Object:  Default [DF__pcDFProds__pcFPr__2645B050]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFProds__pcFPr__2645B050]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFProds__pcFPr__2645B050]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFProds] ADD  CONSTRAINT [DF__pcDFProds__pcFPr__2645B050]  DEFAULT ((0)) FOR [pcFPro_IDProduct]
END


END
GO
/****** Object:  Default [DF__pcDFCusts__pcFCu__2A164134]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFCusts__pcFCu__2A164134]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFCusts__pcFCu__2A164134]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFCusts] ADD  CONSTRAINT [DF__pcDFCusts__pcFCu__2A164134]  DEFAULT ((0)) FOR [pcFCust_IDDiscount]
END


END
GO
/****** Object:  Default [DF__pcDFCusts__pcFCu__2B0A656D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFCusts__pcFCu__2B0A656D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFCusts__pcFCu__2B0A656D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFCusts] ADD  CONSTRAINT [DF__pcDFCusts__pcFCu__2B0A656D]  DEFAULT ((0)) FOR [pcFCust_IDCustomer]
END


END
GO
/****** Object:  Default [DF__pcDFCustP__pcFCP__1D114BD1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFCustP__pcFCP__1D114BD1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFCustP__pcFCP__1D114BD1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFCustPriceCats] ADD  CONSTRAINT [DF__pcDFCustP__pcFCP__1D114BD1]  DEFAULT ((0)) FOR [pcFCPCat_IDDiscount]
END


END
GO
/****** Object:  Default [DF__pcDFCustP__pcFCP__1E05700A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFCustP__pcFCP__1E05700A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFCustP__pcFCP__1E05700A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFCustPriceCats] ADD  CONSTRAINT [DF__pcDFCustP__pcFCP__1E05700A]  DEFAULT ((0)) FOR [pcFCPCat_IDCategory]
END


END
GO
/****** Object:  Default [DF__pcDFCats__pcFCat__2EDAF651]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFCats__pcFCat__2EDAF651]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFCats__pcFCat__2EDAF651]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFCats] ADD  CONSTRAINT [DF__pcDFCats__pcFCat__2EDAF651]  DEFAULT ((0)) FOR [pcFCat_IDDiscount]
END


END
GO
/****** Object:  Default [DF__pcDFCats__pcFCat__2FCF1A8A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFCats__pcFCat__2FCF1A8A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFCats__pcFCat__2FCF1A8A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFCats] ADD  CONSTRAINT [DF__pcDFCats__pcFCat__2FCF1A8A]  DEFAULT ((0)) FOR [pcFCat_IDCategory]
END


END
GO
/****** Object:  Default [DF__pcDFCats__pcFCat__30C33EC3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcDFCats__pcFCat__30C33EC3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcDFCats__pcFCat__30C33EC3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcDFCats] ADD  CONSTRAINT [DF__pcDFCats__pcFCat__30C33EC3]  DEFAULT ((0)) FOR [pcFCat_SubCats]
END


END
GO
/****** Object:  Default [DF__pcCustome__idCus__0C70CFB4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__idCus__0C70CFB4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__idCus__0C70CFB4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerTermsAgreed] ADD  CONSTRAINT [DF__pcCustome__idCus__0C70CFB4]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcCustome__idOrd__0D64F3ED]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__idOrd__0D64F3ED]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__idOrd__0D64F3ED]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerTermsAgreed] ADD  CONSTRAINT [DF__pcCustome__idOrd__0D64F3ED]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__pcCustome__rando__6DCC4D03]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__rando__6DCC4D03]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__rando__6DCC4D03]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF__pcCustome__rando__6DCC4D03]  DEFAULT ((0)) FOR [randomKey]
END


END
GO
/****** Object:  Default [DF__pcCustome__idCus__6EC0713C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__idCus__6EC0713C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__idCus__6EC0713C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF__pcCustome__idCus__6EC0713C]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCus__6FB49575]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCus__6FB49575]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCus__6FB49575]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF__pcCustome__pcCus__6FB49575]  DEFAULT ((0)) FOR [pcCustSession_IdRefer]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCus__70A8B9AE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCus__70A8B9AE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCus__70A8B9AE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF__pcCustome__pcCus__70A8B9AE]  DEFAULT ((0)) FOR [pcCustSession_UseRewards]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCus__719CDDE7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCus__719CDDE7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCus__719CDDE7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF__pcCustome__pcCus__719CDDE7]  DEFAULT ((0)) FOR [pcCustSession_RewardsBalance]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCus__72910220]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCus__72910220]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCus__72910220]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF__pcCustome__pcCus__72910220]  DEFAULT ((0)) FOR [pcCustSession_IdPayment]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCus__73852659]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCus__73852659]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCus__73852659]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF__pcCustome__pcCus__73852659]  DEFAULT ((0)) FOR [pcCustSession_OrdPackageNumber]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_CartRewards]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_CartRewards]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_CartRewards]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_CartRewards]  DEFAULT ((0)) FOR [pcCustSession_CartRewards]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_VATTotal]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_VATTotal]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_VATTotal]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_VATTotal]  DEFAULT ((0)) FOR [pcCustSession_VATTotal]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_intCodeCnt]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_intCodeCnt]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_intCodeCnt]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_intCodeCnt]  DEFAULT ((0)) FOR [pcCustSession_intCodeCnt]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_total]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_total]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_total]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_total]  DEFAULT ((0)) FOR [pcCustSession_total]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_taxAmount]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_taxAmount]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_taxAmount]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_taxAmount]  DEFAULT ((0)) FOR [pcCustSession_taxAmount]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_GWTotal]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_GWTotal]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_GWTotal]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_GWTotal]  DEFAULT ((0)) FOR [pcCustSession_GWTotal]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_ShowShipAddr]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_ShowShipAddr]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_ShowShipAddr]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_ShowShipAddr]  DEFAULT ((0)) FOR [pcCustSession_ShowShipAddr]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_RewardsDollarValue]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_RewardsDollarValue]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_RewardsDollarValue]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_RewardsDollarValue]  DEFAULT ((0)) FOR [pcCustSession_RewardsDollarValue]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_pSubTotal]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_pSubTotal]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_pSubTotal]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_pSubTotal]  DEFAULT ((0)) FOR [pcCustSession_pSubTotal]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_DiscountCodeTotal]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_DiscountCodeTotal]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_DiscountCodeTotal]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_DiscountCodeTotal]  DEFAULT ((0)) FOR [pcCustSession_DiscountCodeTotal]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_CatDiscTotal]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_CatDiscTotal]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_CatDiscTotal]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_CatDiscTotal]  DEFAULT ((0)) FOR [pcCustSession_CatDiscTotal]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCus__2DD1C37F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCus__2DD1C37F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCus__2DD1C37F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF__pcCustome__pcCus__2DD1C37F]  DEFAULT ((0)) FOR [pcCustSession_GCTotal]
END


END
GO
/****** Object:  Default [DF_pcCustomerSessions_pcCustSession_SB_taxAmount]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerSessions_pcCustSession_SB_taxAmount]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerSessions_pcCustSession_SB_taxAmount]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerSessions] ADD  CONSTRAINT [DF_pcCustomerSessions_pcCustSession_SB_taxAmount]  DEFAULT ((0)) FOR [pcCustSession_SB_taxAmount]
END


END
GO
/****** Object:  Default [DF__pcCustome__idCus__3C34F16F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__idCus__3C34F16F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__idCus__3C34F16F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFieldsValues] ADD  CONSTRAINT [DF__pcCustome__idCus__3C34F16F]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCFi__3D2915A8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCFi__3D2915A8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCFi__3D2915A8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFieldsValues] ADD  CONSTRAINT [DF__pcCustome__pcCFi__3D2915A8]  DEFAULT ((0)) FOR [pcCField_ID]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCFi__06CD04F7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCFi__06CD04F7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCFi__06CD04F7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFields] ADD  CONSTRAINT [DF__pcCustome__pcCFi__06CD04F7]  DEFAULT ((0)) FOR [pcCField_FieldType]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCFi__07C12930]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCFi__07C12930]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCFi__07C12930]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFields] ADD  CONSTRAINT [DF__pcCustome__pcCFi__07C12930]  DEFAULT ((0)) FOR [pcCField_Length]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCFi__08B54D69]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCFi__08B54D69]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCFi__08B54D69]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFields] ADD  CONSTRAINT [DF__pcCustome__pcCFi__08B54D69]  DEFAULT ((0)) FOR [pcCField_Maximum]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCFi__09A971A2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCFi__09A971A2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCFi__09A971A2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFields] ADD  CONSTRAINT [DF__pcCustome__pcCFi__09A971A2]  DEFAULT ((0)) FOR [pcCField_Required]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCFi__0A9D95DB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCFi__0A9D95DB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCFi__0A9D95DB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFields] ADD  CONSTRAINT [DF__pcCustome__pcCFi__0A9D95DB]  DEFAULT ((0)) FOR [pcCField_ShowOnReg]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCFi__0B91BA14]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCFi__0B91BA14]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCFi__0B91BA14]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFields] ADD  CONSTRAINT [DF__pcCustome__pcCFi__0B91BA14]  DEFAULT ((0)) FOR [pcCField_ShowOnCheckout]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCFi__0C85DE4D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCFi__0C85DE4D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCFi__0C85DE4D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFields] ADD  CONSTRAINT [DF__pcCustome__pcCFi__0C85DE4D]  DEFAULT ((0)) FOR [pcCField_PricingCategories]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCFi__16EE5E27]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCFi__16EE5E27]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCFi__16EE5E27]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerFields] ADD  CONSTRAINT [DF__pcCustome__pcCFi__16EE5E27]  DEFAULT ((0)) FOR [pcCField_Order]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCC___46B27FE2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCC___46B27FE2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCC___46B27FE2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerCategories] ADD  CONSTRAINT [DF__pcCustome__pcCC___46B27FE2]  DEFAULT ((0)) FOR [pcCC_WholesalePriv]
END


END
GO
/****** Object:  Default [DF__pcCustome__pcCC___47A6A41B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustome__pcCC___47A6A41B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustome__pcCC___47A6A41B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerCategories] ADD  CONSTRAINT [DF__pcCustome__pcCC___47A6A41B]  DEFAULT ((0)) FOR [pcCC_ATB_Percentage]
END


END
GO
/****** Object:  Default [DF_pcCustomerCategories_pcCC_NFSoverride]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCustomerCategories_pcCC_NFSoverride]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCustomerCategories_pcCC_NFSoverride]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustomerCategories] ADD  CONSTRAINT [DF_pcCustomerCategories_pcCC_NFSoverride]  DEFAULT ((0)) FOR [pcCC_NFSoverride]
END


END
GO
/****** Object:  Default [DF__pcCustFie__pcCFi__4B7734FF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustFie__pcCFi__4B7734FF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustFie__pcCFi__4B7734FF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustFieldsPricingCats] ADD  CONSTRAINT [DF__pcCustFie__pcCFi__4B7734FF]  DEFAULT ((0)) FOR [pcCField_ID]
END


END
GO
/****** Object:  Default [DF__pcCustFie__idCus__4C6B5938]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCustFie__idCus__4C6B5938]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCustFie__idCus__4C6B5938]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCustFieldsPricingCats] ADD  CONSTRAINT [DF__pcCustFie__idCus__4C6B5938]  DEFAULT ((0)) FOR [idCustomerCategory]
END


END
GO
/****** Object:  Default [DF_pcCPFProducts_pcCatPro_id]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCPFProducts_pcCatPro_id]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCPFProducts_pcCatPro_id]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCPFProducts] ADD  CONSTRAINT [DF_pcCPFProducts_pcCatPro_id]  DEFAULT ((0)) FOR [pcCatPro_id]
END


END
GO
/****** Object:  Default [DF_pcCPFProducts_idproduct]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCPFProducts_idproduct]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCPFProducts_idproduct]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCPFProducts] ADD  CONSTRAINT [DF_pcCPFProducts_idproduct]  DEFAULT ((0)) FOR [idproduct]
END


END
GO
/****** Object:  Default [DF_pcCPFCategories_pcCatPro_id]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCPFCategories_pcCatPro_id]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCPFCategories_pcCatPro_id]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCPFCategories] ADD  CONSTRAINT [DF_pcCPFCategories_pcCatPro_id]  DEFAULT ((0)) FOR [pcCatPro_id]
END


END
GO
/****** Object:  Default [DF_pcCPFCategories_idcategory]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCPFCategories_idcategory]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCPFCategories_idcategory]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCPFCategories] ADD  CONSTRAINT [DF_pcCPFCategories_idcategory]  DEFAULT ((0)) FOR [idcategory]
END


END
GO
/****** Object:  Default [DF_pcCPFCategories_pcCPFCats_IncSubCats]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCPFCategories_pcCPFCats_IncSubCats]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCPFCategories_pcCPFCats_IncSubCats]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCPFCategories] ADD  CONSTRAINT [DF_pcCPFCategories_pcCPFCats_IncSubCats]  DEFAULT ((0)) FOR [pcCPFCats_IncSubCats]
END


END
GO
/****** Object:  Default [DF__pcContent__pcCon__51300E55]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcContent__pcCon__51300E55]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcContent__pcCon__51300E55]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcContents] ADD  CONSTRAINT [DF__pcContent__pcCon__51300E55]  DEFAULT ((0)) FOR [pcCont_IncHeader]
END


END
GO
/****** Object:  Default [DF__pcContent__pcCon__5224328E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcContent__pcCon__5224328E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcContent__pcCon__5224328E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcContents] ADD  CONSTRAINT [DF__pcContent__pcCon__5224328E]  DEFAULT ((0)) FOR [pcCont_InActive]
END


END
GO
/****** Object:  Default [DF_pcContents_pcCont_Order]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcContents_pcCont_Order]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcContents_pcCont_Order]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcContents] ADD  CONSTRAINT [DF_pcContents_pcCont_Order]  DEFAULT ((0)) FOR [pcCont_Order]
END


END
GO
/****** Object:  Default [DF_pcContents_pcCont_Parent]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcContents_pcCont_Parent]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcContents_pcCont_Parent]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcContents] ADD  CONSTRAINT [DF_pcContents_pcCont_Parent]  DEFAULT ((0)) FOR [pcCont_Parent]
END


END
GO
/****** Object:  Default [DF_pcContents_pcCont_Published]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcContents_pcCont_Published]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcContents_pcCont_Published]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcContents] ADD  CONSTRAINT [DF_pcContents_pcCont_Published]  DEFAULT ((1)) FOR [pcCont_Published]
END


END
GO
/****** Object:  Default [DF_pcContents_pcCont_MenuExclude]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcContents_pcCont_MenuExclude]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcContents_pcCont_MenuExclude]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcContents] ADD  CONSTRAINT [DF_pcContents_pcCont_MenuExclude]  DEFAULT ((0)) FOR [pcCont_MenuExclude]
END


END
GO
/****** Object:  Default [DF_pcContents_pcES_pcCont_HideBackButton]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcContents_pcES_pcCont_HideBackButton]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcContents_pcES_pcCont_HideBackButton]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcContents] ADD  CONSTRAINT [DF_pcContents_pcES_pcCont_HideBackButton]  DEFAULT ((0)) FOR [pcCont_HideBackButton]
END


END
GO
/****** Object:  Default [DF_pcContents_pcCont_DraftStatus]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcContents_pcCont_DraftStatus]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcContents_pcCont_DraftStatus]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcContents] ADD  CONSTRAINT [DF_pcContents_pcCont_DraftStatus]  DEFAULT ((0)) FOR [pcCont_DraftStatus]
END


END
GO
/****** Object:  Default [DF__pcComment__pcCom__55F4C372]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcComment__pcCom__55F4C372]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcComment__pcCom__55F4C372]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcComments] ADD  CONSTRAINT [DF__pcComment__pcCom__55F4C372]  DEFAULT ((0)) FOR [pcComm_IDOrder]
END


END
GO
/****** Object:  Default [DF__pcComment__pcCom__56E8E7AB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcComment__pcCom__56E8E7AB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcComment__pcCom__56E8E7AB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcComments] ADD  CONSTRAINT [DF__pcComment__pcCom__56E8E7AB]  DEFAULT ((0)) FOR [pcComm_IDParent]
END


END
GO
/****** Object:  Default [DF__pcComment__pcCom__57DD0BE4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcComment__pcCom__57DD0BE4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcComment__pcCom__57DD0BE4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcComments] ADD  CONSTRAINT [DF__pcComment__pcCom__57DD0BE4]  DEFAULT ((0)) FOR [pcComm_IDUser]
END


END
GO
/****** Object:  Default [DF__pcComment__pcCom__58D1301D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcComment__pcCom__58D1301D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcComment__pcCom__58D1301D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcComments] ADD  CONSTRAINT [DF__pcComment__pcCom__58D1301D]  DEFAULT ((0)) FOR [pcComm_FType]
END


END
GO
/****** Object:  Default [DF__pcComment__pcCom__59C55456]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcComment__pcCom__59C55456]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcComment__pcCom__59C55456]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcComments] ADD  CONSTRAINT [DF__pcComment__pcCom__59C55456]  DEFAULT ((0)) FOR [pcComm_FStatus]
END


END
GO
/****** Object:  Default [DF__pcComment__pcCom__5AB9788F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcComment__pcCom__5AB9788F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcComment__pcCom__5AB9788F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcComments] ADD  CONSTRAINT [DF__pcComment__pcCom__5AB9788F]  DEFAULT ((0)) FOR [pcComm_Priority]
END


END
GO
/****** Object:  Default [DF__pcCC_Pric__idcus__4262CC11]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCC_Pric__idcus__4262CC11]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCC_Pric__idcus__4262CC11]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCC_Pricing] ADD  CONSTRAINT [DF__pcCC_Pric__idcus__4262CC11]  DEFAULT ((0)) FOR [idcustomerCategory]
END


END
GO
/****** Object:  Default [DF__pcCC_Pric__idPro__4356F04A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCC_Pric__idPro__4356F04A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCC_Pric__idPro__4356F04A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCC_Pricing] ADD  CONSTRAINT [DF__pcCC_Pric__idPro__4356F04A]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__pcCC_Pric__pcCC___444B1483]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCC_Pric__pcCC___444B1483]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCC_Pric__pcCC___444B1483]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCC_Pricing] ADD  CONSTRAINT [DF__pcCC_Pric__pcCC___444B1483]  DEFAULT ((0)) FOR [pcCC_Price]
END


END
GO
/****** Object:  Default [DF__pcCC_BTO___idcus__3D9E16F4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCC_BTO___idcus__3D9E16F4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCC_BTO___idcus__3D9E16F4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCC_BTO_Pricing] ADD  CONSTRAINT [DF__pcCC_BTO___idcus__3D9E16F4]  DEFAULT ((0)) FOR [idcustomerCategory]
END


END
GO
/****** Object:  Default [DF__pcCC_BTO___idBTO__3E923B2D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCC_BTO___idBTO__3E923B2D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCC_BTO___idBTO__3E923B2D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCC_BTO_Pricing] ADD  CONSTRAINT [DF__pcCC_BTO___idBTO__3E923B2D]  DEFAULT ((0)) FOR [idBTOProduct]
END


END
GO
/****** Object:  Default [DF__pcCC_BTO___idBTO__3F865F66]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCC_BTO___idBTO__3F865F66]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCC_BTO___idBTO__3F865F66]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCC_BTO_Pricing] ADD  CONSTRAINT [DF__pcCC_BTO___idBTO__3F865F66]  DEFAULT ((0)) FOR [idBTOItem]
END


END
GO
/****** Object:  Default [DF__pcCC_BTO___pcCC___407A839F]    Script Date: 1/10/2012 17:12:23 ******/

IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCC_BTO___pcCC___407A839F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCC_BTO___pcCC___407A839F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCC_BTO_Pricing] ADD  CONSTRAINT [DF__pcCC_BTO___pcCC___407A839F]  DEFAULT ((0)) FOR [pcCC_BTO_Price]
END


END
GO
/****** Object:  Default [DF_pcCatPromotions_idcategory]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCatPromotions_idcategory]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCatPromotions_idcategory]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatPromotions] ADD  CONSTRAINT [DF_pcCatPromotions_idcategory]  DEFAULT ((0)) FOR [idcategory]
END


END
GO
/****** Object:  Default [DF_pcCatPromotions_pcCatPro_QtyTrigger]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCatPromotions_pcCatPro_QtyTrigger]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCatPromotions_pcCatPro_QtyTrigger]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatPromotions] ADD  CONSTRAINT [DF_pcCatPromotions_pcCatPro_QtyTrigger]  DEFAULT ((0)) FOR [pcCatPro_QtyTrigger]
END


END
GO
/****** Object:  Default [DF_pcCatPromotions_pcCatPro_DiscountType]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCatPromotions_pcCatPro_DiscountType]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCatPromotions_pcCatPro_DiscountType]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatPromotions] ADD  CONSTRAINT [DF_pcCatPromotions_pcCatPro_DiscountType]  DEFAULT ((0)) FOR [pcCatPro_DiscountType]
END


END
GO
/****** Object:  Default [DF_pcCatPromotions_pcCatPro_DiscountValue]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCatPromotions_pcCatPro_DiscountValue]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCatPromotions_pcCatPro_DiscountValue]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatPromotions] ADD  CONSTRAINT [DF_pcCatPromotions_pcCatPro_DiscountValue]  DEFAULT ((0)) FOR [pcCatPro_DiscountValue]
END


END
GO
/****** Object:  Default [DF_pcCatPromotions_pcCatPro_ApplyUnits]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCatPromotions_pcCatPro_ApplyUnits]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCatPromotions_pcCatPro_ApplyUnits]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatPromotions] ADD  CONSTRAINT [DF_pcCatPromotions_pcCatPro_ApplyUnits]  DEFAULT ((0)) FOR [pcCatPro_ApplyUnits]
END


END
GO
/****** Object:  Default [DF__pcCatDisc__pcCD___5E8A0973]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCatDisc__pcCD___5E8A0973]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCatDisc__pcCD___5E8A0973]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatDiscounts] ADD  CONSTRAINT [DF__pcCatDisc__pcCD___5E8A0973]  DEFAULT ((0)) FOR [pcCD_IDCategory]
END


END
GO
/****** Object:  Default [DF__pcCatDisc__pcCD___5F7E2DAC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCatDisc__pcCD___5F7E2DAC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCatDisc__pcCD___5F7E2DAC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatDiscounts] ADD  CONSTRAINT [DF__pcCatDisc__pcCD___5F7E2DAC]  DEFAULT ((0)) FOR [pcCD_quantityFrom]
END


END
GO
/****** Object:  Default [DF__pcCatDisc__pcCD___607251E5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCatDisc__pcCD___607251E5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCatDisc__pcCD___607251E5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatDiscounts] ADD  CONSTRAINT [DF__pcCatDisc__pcCD___607251E5]  DEFAULT ((0)) FOR [pcCD_quantityUntil]
END


END
GO
/****** Object:  Default [DF__pcCatDisc__pcCD___6166761E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCatDisc__pcCD___6166761E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCatDisc__pcCD___6166761E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatDiscounts] ADD  CONSTRAINT [DF__pcCatDisc__pcCD___6166761E]  DEFAULT ((0)) FOR [pcCD_discountPerUnit]
END


END
GO
/****** Object:  Default [DF__pcCatDisc__pcCD___625A9A57]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCatDisc__pcCD___625A9A57]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCatDisc__pcCD___625A9A57]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatDiscounts] ADD  CONSTRAINT [DF__pcCatDisc__pcCD___625A9A57]  DEFAULT ((0)) FOR [pcCD_num]
END


END
GO
/****** Object:  Default [DF__pcCatDisc__pcCD___634EBE90]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCatDisc__pcCD___634EBE90]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCatDisc__pcCD___634EBE90]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatDiscounts] ADD  CONSTRAINT [DF__pcCatDisc__pcCD___634EBE90]  DEFAULT ((0)) FOR [pcCD_percentage]
END


END
GO
/****** Object:  Default [DF__pcCatDisc__pcCD___6442E2C9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCatDisc__pcCD___6442E2C9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCatDisc__pcCD___6442E2C9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatDiscounts] ADD  CONSTRAINT [DF__pcCatDisc__pcCD___6442E2C9]  DEFAULT ((0)) FOR [pcCD_discountPerWUnit]
END


END
GO
/****** Object:  Default [DF__pcCatDisc__pcCD___65370702]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcCatDisc__pcCD___65370702]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcCatDisc__pcCD___65370702]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCatDiscounts] ADD  CONSTRAINT [DF__pcCatDisc__pcCD___65370702]  DEFAULT ((0)) FOR [pcCD_baseproductonly]
END


END
GO
/****** Object:  Default [DF_pcCartArray_pcCartArray_3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCartArray_pcCartArray_3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCartArray_pcCartArray_3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCartArray] ADD  CONSTRAINT [DF_pcCartArray_pcCartArray_3]  DEFAULT ((0)) FOR [pcCartArray_3]
END


END
GO
/****** Object:  Default [DF_pcCartArray_pcCartArray_5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCartArray_pcCartArray_5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCartArray_pcCartArray_5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCartArray] ADD  CONSTRAINT [DF_pcCartArray_pcCartArray_5]  DEFAULT ((0)) FOR [pcCartArray_5]
END


END
GO
/****** Object:  Default [DF_pcCartArray_pcCartArray_14]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCartArray_pcCartArray_14]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCartArray_pcCartArray_14]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCartArray] ADD  CONSTRAINT [DF_pcCartArray_pcCartArray_14]  DEFAULT ((0)) FOR [pcCartArray_14]
END


END
GO
/****** Object:  Default [DF_pcCartArray_pcCartArray_15]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCartArray_pcCartArray_15]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCartArray_pcCartArray_15]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCartArray] ADD  CONSTRAINT [DF_pcCartArray_pcCartArray_15]  DEFAULT ((0)) FOR [pcCartArray_15]
END


END
GO
/****** Object:  Default [DF_pcCartArray_pcCartArray_17]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCartArray_pcCartArray_17]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCartArray_pcCartArray_17]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCartArray] ADD  CONSTRAINT [DF_pcCartArray_pcCartArray_17]  DEFAULT ((0)) FOR [pcCartArray_17]
END


END
GO
/****** Object:  Default [DF_pcCartArray_pcCartArray_28]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcCartArray_pcCartArray_28]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcCartArray_pcCartArray_28]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcCartArray] ADD  CONSTRAINT [DF_pcCartArray_pcCartArray_28]  DEFAULT ((0)) FOR [pcCartArray_28]
END


END
GO
/****** Object:  Default [DF_pcBTODefaultPriceCats_idProduct]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcBTODefaultPriceCats_idProduct]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcBTODefaultPriceCats_idProduct]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcBTODefaultPriceCats] ADD  CONSTRAINT [DF_pcBTODefaultPriceCats_idProduct]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF_pcBTODefaultPriceCats_idCustomerCategory]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcBTODefaultPriceCats_idCustomerCategory]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcBTODefaultPriceCats_idCustomerCategory]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcBTODefaultPriceCats] ADD  CONSTRAINT [DF_pcBTODefaultPriceCats_idCustomerCategory]  DEFAULT ((0)) FOR [idCustomerCategory]
END


END
GO
/****** Object:  Default [DF_pcBTODefaultPriceCats_pcBDPC_Price]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcBTODefaultPriceCats_pcBDPC_Price]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcBTODefaultPriceCats_pcBDPC_Price]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcBTODefaultPriceCats] ADD  CONSTRAINT [DF_pcBTODefaultPriceCats_pcBDPC_Price]  DEFAULT ((0)) FOR [pcBDPC_Price]
END


END
GO
/****** Object:  Default [DF__pcBestSel__pcBSS__5070F446]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcBestSel__pcBSS__5070F446]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcBestSel__pcBSS__5070F446]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcBestSellerSettings] ADD  CONSTRAINT [DF__pcBestSel__pcBSS__5070F446]  DEFAULT ((0)) FOR [pcBSS_BestSellCount]
END


END
GO
/****** Object:  Default [DF__pcBestSel__pcBSS__5165187F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcBestSel__pcBSS__5165187F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcBestSel__pcBSS__5165187F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcBestSellerSettings] ADD  CONSTRAINT [DF__pcBestSel__pcBSS__5165187F]  DEFAULT ((0)) FOR [pcBSS_NSold]
END


END
GO
/****** Object:  Default [DF__pcBestSel__pcBSS__52593CB8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcBestSel__pcBSS__52593CB8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcBestSel__pcBSS__52593CB8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcBestSellerSettings] ADD  CONSTRAINT [DF__pcBestSel__pcBSS__52593CB8]  DEFAULT ((0)) FOR [pcBSS_NotForSale]
END


END
GO
/****** Object:  Default [DF__pcBestSel__pcBSS__534D60F1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcBestSel__pcBSS__534D60F1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcBestSel__pcBSS__534D60F1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcBestSellerSettings] ADD  CONSTRAINT [DF__pcBestSel__pcBSS__534D60F1]  DEFAULT ((0)) FOR [pcBSS_OutOfStock]
END


END
GO
/****** Object:  Default [DF__pcBestSel__pcBSS__5441852A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcBestSel__pcBSS__5441852A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcBestSel__pcBSS__5441852A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcBestSellerSettings] ADD  CONSTRAINT [DF__pcBestSel__pcBSS__5441852A]  DEFAULT ((0)) FOR [pcBSS_SKU]
END


END
GO
/****** Object:  Default [DF__pcBestSel__pcBSS__5535A963]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcBestSel__pcBSS__5535A963]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcBestSel__pcBSS__5535A963]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcBestSellerSettings] ADD  CONSTRAINT [DF__pcBestSel__pcBSS__5535A963]  DEFAULT ((0)) FOR [pcBSS_ShowImg]
END


END
GO
/****** Object:  Default [DF__pcAmazonS__pcAmz__45A94D10]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAmazonS__pcAmz__45A94D10]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAmazonS__pcAmz__45A94D10]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazonSettings] ADD  CONSTRAINT [DF__pcAmazonS__pcAmz__45A94D10]  DEFAULT ((0)) FOR [pcAmzSet_prdIDType]
END


END
GO
/****** Object:  Default [DF__pcAmazonS__pcAmz__469D7149]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAmazonS__pcAmz__469D7149]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAmazonS__pcAmz__469D7149]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazonSettings] ADD  CONSTRAINT [DF__pcAmazonS__pcAmz__469D7149]  DEFAULT ((0)) FOR [pcAmzSet_icondition]
END


END
GO
/****** Object:  Default [DF__pcAmazonS__pcAmz__47919582]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAmazonS__pcAmz__47919582]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAmazonS__pcAmz__47919582]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazonSettings] ADD  CONSTRAINT [DF__pcAmazonS__pcAmz__47919582]  DEFAULT ((0)) FOR [pcAmzSet_price]
END


END
GO
/****** Object:  Default [DF__pcAmazonS__pcAmz__4885B9BB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAmazonS__pcAmz__4885B9BB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAmazonS__pcAmz__4885B9BB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazonSettings] ADD  CONSTRAINT [DF__pcAmazonS__pcAmz__4885B9BB]  DEFAULT ((0)) FOR [pcAmzSet_willshipout]
END


END
GO
/****** Object:  Default [DF_pcAmazon_idproduct]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcAmazon_idproduct]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcAmazon_idproduct]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazon] ADD  CONSTRAINT [DF_pcAmazon_idproduct]  DEFAULT ((0)) FOR [idproduct]
END


END
GO
/****** Object:  Default [DF_pcAmazon_pcAmz_prdIDType]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcAmazon_pcAmz_prdIDType]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcAmazon_pcAmz_prdIDType]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazon] ADD  CONSTRAINT [DF_pcAmazon_pcAmz_prdIDType]  DEFAULT ((0)) FOR [pcAmz_prdIDType]
END


END
GO
/****** Object:  Default [DF_pcAmazon_pcAmz_icondition]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcAmazon_pcAmz_icondition]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcAmazon_pcAmz_icondition]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazon] ADD  CONSTRAINT [DF_pcAmazon_pcAmz_icondition]  DEFAULT ((0)) FOR [pcAmz_icondition]
END


END
GO
/****** Object:  Default [DF_pcAmazon_pcAmz_price]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcAmazon_pcAmz_price]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcAmazon_pcAmz_price]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazon] ADD  CONSTRAINT [DF_pcAmazon_pcAmz_price]  DEFAULT ((0)) FOR [pcAmz_price]
END


END
GO
/****** Object:  Default [DF_pcAmazon_pcAmz_quantity]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcAmazon_pcAmz_quantity]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcAmazon_pcAmz_quantity]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazon] ADD  CONSTRAINT [DF_pcAmazon_pcAmz_quantity]  DEFAULT ((0)) FOR [pcAmz_quantity]
END


END
GO
/****** Object:  Default [DF_pcAmazon_pcAmz_willshipout]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_pcAmazon_pcAmz_willshipout]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_pcAmazon_pcAmz_willshipout]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAmazon] ADD  CONSTRAINT [DF_pcAmazon_pcAmz_willshipout]  DEFAULT ((0)) FOR [pcAmz_willshipout]
END


END
GO
/****** Object:  Default [DF__pcAffilia__pcAff__41EDCAC5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAffilia__pcAff__41EDCAC5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAffilia__pcAff__41EDCAC5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAffiliatesPayments] ADD  CONSTRAINT [DF__pcAffilia__pcAff__41EDCAC5]  DEFAULT ((0)) FOR [pcAffpay_idAffiliate]
END


END
GO
/****** Object:  Default [DF__pcAffilia__pcAff__42E1EEFE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAffilia__pcAff__42E1EEFE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAffilia__pcAff__42E1EEFE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAffiliatesPayments] ADD  CONSTRAINT [DF__pcAffilia__pcAff__42E1EEFE]  DEFAULT ((0)) FOR [pcAffpay_Amount]
END


END
GO
/****** Object:  Default [DF__pcAdminCo__idOrd__3493CFA7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAdminCo__idOrd__3493CFA7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAdminCo__idOrd__3493CFA7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAdminComments] ADD  CONSTRAINT [DF__pcAdminCo__idOrd__3493CFA7]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__pcAdminCo__pcACO__3587F3E0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAdminCo__pcACO__3587F3E0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAdminCo__pcACO__3587F3E0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAdminComments] ADD  CONSTRAINT [DF__pcAdminCo__pcACO__3587F3E0]  DEFAULT ((0)) FOR [pcACOM_ComType]
END


END
GO
/****** Object:  Default [DF__pcAdminCo__pcDro__367C1819]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAdminCo__pcDro__367C1819]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAdminCo__pcDro__367C1819]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAdminComments] ADD  CONSTRAINT [DF__pcAdminCo__pcDro__367C1819]  DEFAULT ((0)) FOR [pcDropShipper_ID]
END


END
GO
/****** Object:  Default [DF__pcAdminCo__pcACo__37703C52]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAdminCo__pcACo__37703C52]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAdminCo__pcACo__37703C52]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAdminComments] ADD  CONSTRAINT [DF__pcAdminCo__pcACo__37703C52]  DEFAULT ((0)) FOR [pcACom_IsSupplier]
END


END
GO
/****** Object:  Default [DF__pcAdminCo__pcPac__3864608B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__pcAdminCo__pcPac__3864608B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__pcAdminCo__pcPac__3864608B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[pcAdminComments] ADD  CONSTRAINT [DF__pcAdminCo__pcPac__3864608B]  DEFAULT ((0)) FOR [pcPackageInfo_ID]
END


END
GO
/****** Object:  Default [DF__payTypes__gwCode__7D78A4E7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__gwCode__7D78A4E7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__gwCode__7D78A4E7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__gwCode__7D78A4E7]  DEFAULT ((0)) FOR [gwCode]
END


END
GO
/****** Object:  Default [DF__payTypes__priceT__7E6CC920]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__priceT__7E6CC920]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__priceT__7E6CC920]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__priceT__7E6CC920]  DEFAULT ((0)) FOR [priceToAdd]
END


END
GO
/****** Object:  Default [DF__payTypes__percen__7F60ED59]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__percen__7F60ED59]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__percen__7F60ED59]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__percen__7F60ED59]  DEFAULT ((0)) FOR [percentageToAdd]
END


END
GO
/****** Object:  Default [DF__payTypes__ssl__00551192]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__ssl__00551192]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__ssl__00551192]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__ssl__00551192]  DEFAULT ((0)) FOR [ssl]
END


END
GO
/****** Object:  Default [DF__payTypes__quanti__014935CB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__quanti__014935CB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__quanti__014935CB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__quanti__014935CB]  DEFAULT ((0)) FOR [quantityFrom]
END


END
GO
/****** Object:  Default [DF__payTypes__quanti__023D5A04]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__quanti__023D5A04]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__quanti__023D5A04]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__quanti__023D5A04]  DEFAULT ((0)) FOR [quantityUntil]
END


END
GO
/****** Object:  Default [DF__payTypes__weight__03317E3D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__weight__03317E3D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__weight__03317E3D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__weight__03317E3D]  DEFAULT ((0)) FOR [weightFrom]
END


END
GO
/****** Object:  Default [DF__payTypes__weight__0425A276]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__weight__0425A276]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__weight__0425A276]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__weight__0425A276]  DEFAULT ((0)) FOR [weightUntil]
END


END
GO
/****** Object:  Default [DF__payTypes__priceF__0519C6AF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__priceF__0519C6AF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__priceF__0519C6AF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__priceF__0519C6AF]  DEFAULT ((0)) FOR [priceFrom]
END


END
GO
/****** Object:  Default [DF__payTypes__priceU__060DEAE8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__priceU__060DEAE8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__priceU__060DEAE8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__priceU__060DEAE8]  DEFAULT ((0)) FOR [priceUntil]
END


END
GO
/****** Object:  Default [DF__payTypes__active__07020F21]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__active__07020F21]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__active__07020F21]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__active__07020F21]  DEFAULT ((0)) FOR [active]
END


END
GO
/****** Object:  Default [DF__payTypes__Cbtob__07F6335A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__Cbtob__07F6335A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__Cbtob__07F6335A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__Cbtob__07F6335A]  DEFAULT ((0)) FOR [Cbtob]
END


END
GO
/****** Object:  Default [DF__payTypes__CReq__08EA5793]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__CReq__08EA5793]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__CReq__08EA5793]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__CReq__08EA5793]  DEFAULT ((0)) FOR [CReq]
END


END
GO
/****** Object:  Default [DF_cvv_payTypes]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_cvv_payTypes]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_cvv_payTypes]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF_cvv_payTypes]  DEFAULT ((0)) FOR [cvv]
END


END
GO
/****** Object:  Default [DF__payTypes__paymen__09DE7BCC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__paymen__09DE7BCC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__paymen__09DE7BCC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__paymen__09DE7BCC]  DEFAULT ((0)) FOR [paymentPriority]
END


END
GO
/****** Object:  Default [DF__payTypes__pcPayT__0AD2A005]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__pcPayT__0AD2A005]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__pcPayT__0AD2A005]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__pcPayT__0AD2A005]  DEFAULT ((0)) FOR [pcPayTypes_processOrder]
END


END
GO
/****** Object:  Default [DF__payTypes__pcPayT__0BC6C43E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__pcPayT__0BC6C43E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__pcPayT__0BC6C43E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__pcPayT__0BC6C43E]  DEFAULT ((0)) FOR [pcPayTypes_setPayStatus]
END


END
GO
/****** Object:  Default [DF__payTypes__pcPayT__0BC6C46F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__payTypes__pcPayT__0BC6C46F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__payTypes__pcPayT__0BC6C46F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF__payTypes__pcPayT__0BC6C46F]  DEFAULT ((0)) FOR [pcPayTypes_ppab]
END


END
GO
/****** Object:  Default [DF_payTypes_pcPayTypes_Subscription]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_payTypes_pcPayTypes_Subscription]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_payTypes_pcPayTypes_Subscription]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[payTypes] ADD  CONSTRAINT [DF_payTypes_pcPayTypes_Subscription]  DEFAULT ((0)) FOR [pcPayTypes_Subscription]
END


END
GO
/****** Object:  Default [DF__paypal__id__108B795B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__paypal__id__108B795B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__paypal__id__108B795B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[paypal] ADD  CONSTRAINT [DF__paypal__id__108B795B]  DEFAULT ((0)) FOR [id]
END


END
GO
/****** Object:  Default [DF__paypal__PP_Curre__117F9D94]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__paypal__PP_Curre__117F9D94]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__paypal__PP_Curre__117F9D94]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[paypal] ADD  CONSTRAINT [DF__paypal__PP_Curre__117F9D94]  DEFAULT ('USD') FOR [PP_Currency]
END


END
GO
/****** Object:  Default [DF__paypal__PP_Sandb__2145C81B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__paypal__PP_Sandb__2145C81B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__paypal__PP_Sandb__2145C81B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[paypal] ADD  CONSTRAINT [DF__paypal__PP_Sandb__2145C81B]  DEFAULT ((0)) FOR [PP_Sandbox]
END


END
GO
/****** Object:  Default [DF__orders__idCustom__15502E78]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__idCustom__15502E78]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__idCustom__15502E78]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__idCustom__15502E78]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__orders__total__164452B1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__total__164452B1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__total__164452B1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__total__164452B1]  DEFAULT ((0)) FOR [total]
END


END
GO
/****** Object:  Default [DF__orders__taxAmoun__173876EA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__taxAmoun__173876EA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__taxAmoun__173876EA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__taxAmoun__173876EA]  DEFAULT ((0)) FOR [taxAmount]
END


END
GO
/****** Object:  Default [DF__orders__randomNu__182C9B23]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__randomNu__182C9B23]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__randomNu__182C9B23]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__randomNu__182C9B23]  DEFAULT ((0)) FOR [randomNumber]
END


END
GO
/****** Object:  Default [DF__orders__orderSta__1920BF5C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__orderSta__1920BF5C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__orderSta__1920BF5C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__orderSta__1920BF5C]  DEFAULT ((0)) FOR [orderStatus]
END


END
GO
/****** Object:  Default [DF__orders__viewed__1A14E395]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__viewed__1A14E395]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__viewed__1A14E395]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__viewed__1A14E395]  DEFAULT ((0)) FOR [viewed]
END


END
GO
/****** Object:  Default [DF__orders__idAffili__1B0907CE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__idAffili__1B0907CE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__idAffili__1B0907CE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__idAffili__1B0907CE]  DEFAULT ((0)) FOR [idAffiliate]
END


END
GO
/****** Object:  Default [DF__orders__iRewardP__1BFD2C07]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__iRewardP__1BFD2C07]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__iRewardP__1BFD2C07]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__iRewardP__1BFD2C07]  DEFAULT ((0)) FOR [iRewardPoints]
END


END
GO
/****** Object:  Default [DF__orders__iRewardV__1CF15040]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__iRewardV__1CF15040]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__iRewardV__1CF15040]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__iRewardV__1CF15040]  DEFAULT ((0)) FOR [iRewardValue]
END


END
GO
/****** Object:  Default [DF__orders__iRewardR__1DE57479]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__iRewardR__1DE57479]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__iRewardR__1DE57479]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__iRewardR__1DE57479]  DEFAULT ((0)) FOR [iRewardRefId]
END


END
GO
/****** Object:  Default [DF__orders__iRewardP__1ED998B2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__iRewardP__1ED998B2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__iRewardP__1ED998B2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__iRewardP__1ED998B2]  DEFAULT ((0)) FOR [iRewardPointsRef]
END


END
GO
/****** Object:  Default [DF__orders__iRewardP__1FCDBCEB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__iRewardP__1FCDBCEB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__iRewardP__1FCDBCEB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__iRewardP__1FCDBCEB]  DEFAULT ((0)) FOR [iRewardPointsCustAccrued]
END


END
GO
/****** Object:  Default [DF__orders__IDRefer__20C1E124]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__IDRefer__20C1E124]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__IDRefer__20C1E124]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__IDRefer__20C1E124]  DEFAULT ((0)) FOR [IDRefer]
END


END
GO
/****** Object:  Default [DF__orders__rmaCredi__22AA2996]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__rmaCredi__22AA2996]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__rmaCredi__22AA2996]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__rmaCredi__22AA2996]  DEFAULT ((0)) FOR [rmaCredit]
END


END
GO
/****** Object:  Default [DF__orders__DPs__21B6055D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__DPs__21B6055D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__DPs__21B6055D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__DPs__21B6055D]  DEFAULT ((0)) FOR [DPs]
END


END
GO
/****** Object:  Default [DF__orders__SRF__239E4DCF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__SRF__239E4DCF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__SRF__239E4DCF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__SRF__239E4DCF]  DEFAULT ((0)) FOR [SRF]
END


END
GO
/****** Object:  Default [DF__orders__ordShipt__24927208]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__ordShipt__24927208]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__ordShipt__24927208]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__ordShipt__24927208]  DEFAULT ((0)) FOR [ordShiptype]
END


END
GO
/****** Object:  Default [DF__orders__ordPacka__25869641]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__ordPacka__25869641]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__ordPacka__25869641]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__ordPacka__25869641]  DEFAULT ((1)) FOR [ordPackageNum]
END


END
GO
/****** Object:  Default [DF__orders__ord_VAT__267ABA7A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__ord_VAT__267ABA7A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__ord_VAT__267ABA7A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__ord_VAT__267ABA7A]  DEFAULT ((0)) FOR [ord_VAT]
END


END
GO
/****** Object:  Default [DF__orders__pcOrd_Ca__276EDEB3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__pcOrd_Ca__276EDEB3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__pcOrd_Ca__276EDEB3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__pcOrd_Ca__276EDEB3]  DEFAULT ((0)) FOR [pcOrd_CatDiscounts]
END


END
GO
/****** Object:  Default [DF__orders__pcOrd_Pa__286302EC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__pcOrd_Pa__286302EC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__pcOrd_Pa__286302EC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__pcOrd_Pa__286302EC]  DEFAULT ((0)) FOR [pcOrd_PaymentStatus]
END


END
GO
/****** Object:  Default [DF__orders__pcOrd_Cu__29572725]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__pcOrd_Cu__29572725]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__pcOrd_Cu__29572725]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__pcOrd_Cu__29572725]  DEFAULT ((0)) FOR [pcOrd_CustAllowSeparate]
END


END
GO
/****** Object:  Default [DF__orders__pcOrd_Gc__2A4B4B5E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__pcOrd_Gc__2A4B4B5E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__pcOrd_Gc__2A4B4B5E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__pcOrd_Gc__2A4B4B5E]  DEFAULT ((0)) FOR [pcOrd_GcUsed]
END


END
GO
/****** Object:  Default [DF__orders__pcOrd_GC__2B3F6F97]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__pcOrd_GC__2B3F6F97]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__pcOrd_GC__2B3F6F97]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__pcOrd_GC__2B3F6F97]  DEFAULT ((0)) FOR [pcOrd_GCs]
END


END
GO
/****** Object:  Default [DF__orders__pcOrd_ID__2C3393D0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__pcOrd_ID__2C3393D0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__pcOrd_ID__2C3393D0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__pcOrd_ID__2C3393D0]  DEFAULT ((0)) FOR [pcOrd_IDEvent]
END


END
GO
/****** Object:  Default [DF__orders__pcOrd_GW__2D27B809]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__pcOrd_GW__2D27B809]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__pcOrd_GW__2D27B809]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__pcOrd_GW__2D27B809]  DEFAULT ((0)) FOR [pcOrd_GWTotal]
END


END
GO
/****** Object:  Default [DF_Orders_pcOrd_ShipWeight]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_Orders_pcOrd_ShipWeight]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_Orders_pcOrd_ShipWeight]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF_Orders_pcOrd_ShipWeight]  DEFAULT ((0)) FOR [pcOrd_ShipWeight]
END


END
GO
/****** Object:  Default [DF_orders_pcOrd_Archived]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_orders_pcOrd_Archived]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_orders_pcOrd_Archived]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF_orders_pcOrd_Archived]  DEFAULT ((0)) FOR [pcOrd_Archived]
END


END
GO
/****** Object:  Default [DF_orders_pcOrd_MobileSF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_orders_pcOrd_MobileSF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_orders_pcOrd_MobileSF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF_orders_pcOrd_MobileSF]  DEFAULT ((0)) FOR [pcOrd_MobileSF]
END


END
GO
/****** Object:  Default [DF_orders_pcOrd_ShowShipAddr]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_orders_pcOrd_ShowShipAddr]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_orders_pcOrd_ShowShipAddr]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF_orders_pcOrd_ShowShipAddr]  DEFAULT ((1)) FOR [pcOrd_ShowShipAddr]
END


END
GO
/****** Object:  Default [DF__orders__pcOrd_GC__20ACD28B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__orders__pcOrd_GC__20ACD28B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__orders__pcOrd_GC__20ACD28B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF__orders__pcOrd_GC__20ACD28B]  DEFAULT ((0)) FOR [pcOrd_GCAmount]
END


END
GO
/****** Object:  Default [DF_orders_pcOrd_SubTax]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_orders_pcOrd_SubTax]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_orders_pcOrd_SubTax]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF_orders_pcOrd_SubTax]  DEFAULT ((0)) FOR [pcOrd_SubTax]
END


END
GO
/****** Object:  Default [DF_orders_pcOrd_SubTrialTax]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_orders_pcOrd_SubTrialTax]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_orders_pcOrd_SubTrialTax]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF_orders_pcOrd_SubTrialTax]  DEFAULT ((0)) FOR [pcOrd_SubTrialTax]
END


END
GO
/****** Object:  Default [DF_orders_pcOrd_SubShipping]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_orders_pcOrd_SubShipping]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_orders_pcOrd_SubShipping]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF_orders_pcOrd_SubShipping]  DEFAULT ((0)) FOR [pcOrd_SubShipping]
END


END
GO
/****** Object:  Default [DF_orders_pcOrd_SubTrialShipping]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_orders_pcOrd_SubTrialShipping]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_orders_pcOrd_SubTrialShipping]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[orders] ADD  CONSTRAINT [DF_orders_pcOrd_SubTrialShipping]  DEFAULT ((0)) FOR [pcOrd_SubTrialShipping]
END


END
GO
/****** Object:  Default [DF__options_o__idPro__35BCFE0A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__options_o__idPro__35BCFE0A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__options_o__idPro__35BCFE0A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[options_optionsGroups] ADD  CONSTRAINT [DF__options_o__idPro__35BCFE0A]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__options_o__idOpt__36B12243]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__options_o__idOpt__36B12243]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__options_o__idOpt__36B12243]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[options_optionsGroups] ADD  CONSTRAINT [DF__options_o__idOpt__36B12243]  DEFAULT ((0)) FOR [idOptionGroup]
END


END
GO
/****** Object:  Default [DF__options_o__idOpt__37A5467C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__options_o__idOpt__37A5467C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__options_o__idOpt__37A5467C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[options_optionsGroups] ADD  CONSTRAINT [DF__options_o__idOpt__37A5467C]  DEFAULT ((0)) FOR [idOption]
END


END
GO
/****** Object:  Default [DF__options_o__price__38996AB5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__options_o__price__38996AB5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__options_o__price__38996AB5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[options_optionsGroups] ADD  CONSTRAINT [DF__options_o__price__38996AB5]  DEFAULT ((0)) FOR [price]
END


END
GO
/****** Object:  Default [DBX_Wprice_32317]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_Wprice_32317]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_Wprice_32317]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[options_optionsGroups] ADD  CONSTRAINT [DBX_Wprice_32317]  DEFAULT ((0)) FOR [Wprice]
END


END
GO
/****** Object:  Default [DF__options_o__sortO__398D8EEE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__options_o__sortO__398D8EEE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__options_o__sortO__398D8EEE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[options_optionsGroups] ADD  CONSTRAINT [DF__options_o__sortO__398D8EEE]  DEFAULT ((0)) FOR [sortOrder]
END


END
GO
/****** Object:  Default [DF__options_o__InAct__3A81B327]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__options_o__InAct__3A81B327]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__options_o__InAct__3A81B327]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[options_optionsGroups] ADD  CONSTRAINT [DF__options_o__InAct__3A81B327]  DEFAULT ((0)) FOR [InActive]
END


END
GO
/****** Object:  Default [DF__optGrps__idOptio__4316F928]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__optGrps__idOptio__4316F928]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__optGrps__idOptio__4316F928]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[optGrps] ADD  CONSTRAINT [DF__optGrps__idOptio__4316F928]  DEFAULT ((0)) FOR [idOptionGroup]
END


END
GO
/****** Object:  Default [DF__optGrps__idoptio__440B1D61]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__optGrps__idoptio__440B1D61]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__optGrps__idoptio__440B1D61]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[optGrps] ADD  CONSTRAINT [DF__optGrps__idoptio__440B1D61]  DEFAULT ((0)) FOR [idoption]
END


END
GO
/****** Object:  Default [DF__offlinepa__idOrd__02084FDA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__offlinepa__idOrd__02084FDA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__offlinepa__idOrd__02084FDA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[offlinepayments] ADD  CONSTRAINT [DF__offlinepa__idOrd__02084FDA]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__offlinepa__idpay__02FC7413]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__offlinepa__idpay__02FC7413]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__offlinepa__idpay__02FC7413]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[offlinepayments] ADD  CONSTRAINT [DF__offlinepa__idpay__02FC7413]  DEFAULT ((0)) FOR [idpayment]
END


END
GO
/****** Object:  Default [DF__News__msgtype__4AB81AF0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__News__msgtype__4AB81AF0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__News__msgtype__4AB81AF0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[News] ADD  CONSTRAINT [DF__News__msgtype__4AB81AF0]  DEFAULT ((0)) FOR [msgtype]
END


END
GO
/****** Object:  Default [DF__News__custtotal__4BAC3F29]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__News__custtotal__4BAC3F29]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__News__custtotal__4BAC3F29]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[News] ADD  CONSTRAINT [DF__News__custtotal__4BAC3F29]  DEFAULT ((0)) FOR [custtotal]
END


END
GO
/****** Object:  Default [DF__netbillor__idOrd__76CBA758]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__netbillor__idOrd__76CBA758]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__netbillor__idOrd__76CBA758]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[netbillorders] ADD  CONSTRAINT [DF__netbillor__idOrd__76CBA758]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__netbillor__amoun__77BFCB91]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__netbillor__amoun__77BFCB91]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__netbillor__amoun__77BFCB91]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[netbillorders] ADD  CONSTRAINT [DF__netbillor__amoun__77BFCB91]  DEFAULT ((0)) FOR [amount]
END


END
GO
/****** Object:  Default [DF__netbillor__idCus__78B3EFCA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__netbillor__idCus__78B3EFCA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__netbillor__idCus__78B3EFCA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[netbillorders] ADD  CONSTRAINT [DF__netbillor__idCus__78B3EFCA]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__netbillor__captu__79A81403]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__netbillor__captu__79A81403]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__netbillor__captu__79A81403]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[netbillorders] ADD  CONSTRAINT [DF__netbillor__captu__79A81403]  DEFAULT ((0)) FOR [captured]
END


END
GO
/****** Object:  Default [DF_netbillorders_pcSecurityKeyID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_netbillorders_pcSecurityKeyID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_netbillorders_pcSecurityKeyID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[netbillorders] ADD  CONSTRAINT [DF_netbillorders_pcSecurityKeyID]  DEFAULT ((0)) FOR [pcSecurityKeyID]
END


END
GO
/****** Object:  Default [DF__netbill__idNetbi__59063A47]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__netbill__idNetbi__59063A47]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__netbill__idNetbi__59063A47]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[netbill] ADD  CONSTRAINT [DF__netbill__idNetbi__59063A47]  DEFAULT ((0)) FOR [idNetbill]
END


END
GO
/****** Object:  Default [DBX_NBCVVEnabled_20432]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_NBCVVEnabled_20432]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_NBCVVEnabled_20432]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[netbill] ADD  CONSTRAINT [DBX_NBCVVEnabled_20432]  DEFAULT ((0)) FOR [NBCVVEnabled]
END


END
GO
/****** Object:  Default [DF__netbill__NBAVS__59FA5E80]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__netbill__NBAVS__59FA5E80]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__netbill__NBAVS__59FA5E80]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[netbill] ADD  CONSTRAINT [DF__netbill__NBAVS__59FA5E80]  DEFAULT ((0)) FOR [NBAVS]
END


END
GO
/****** Object:  Default [DF__netbill__Netbill__5AEE82B9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__netbill__Netbill__5AEE82B9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__netbill__Netbill__5AEE82B9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[netbill] ADD  CONSTRAINT [DF__netbill__Netbill__5AEE82B9]  DEFAULT ((0)) FOR [NetbillCheck]
END


END
GO
/****** Object:  Default [DF__linkpoint__id__619B8048]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__linkpoint__id__619B8048]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__linkpoint__id__619B8048]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[linkpoint] ADD  CONSTRAINT [DF__linkpoint__id__619B8048]  DEFAULT ((0)) FOR [id]
END


END
GO
/****** Object:  Default [DF__linkpoint__CVM__628FA481]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__linkpoint__CVM__628FA481]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__linkpoint__CVM__628FA481]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[linkpoint] ADD  CONSTRAINT [DF__linkpoint__CVM__628FA481]  DEFAULT ((0)) FOR [CVM]
END


END
GO
/****** Object:  Default [DF__klix__CVV__6A30C649]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__klix__CVV__6A30C649]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__klix__CVV__6A30C649]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[klix] ADD  CONSTRAINT [DF__klix__CVV__6A30C649]  DEFAULT ((0)) FOR [CVV]
END


END
GO
/****** Object:  Default [DF__klix__ssl_avs__6B24EA82]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__klix__ssl_avs__6B24EA82]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__klix__ssl_avs__6B24EA82]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[klix] ADD  CONSTRAINT [DF__klix__ssl_avs__6B24EA82]  DEFAULT ((0)) FOR [ssl_avs]
END


END
GO
/****** Object:  Default [DF__klix__testmode__6C190EBB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__klix__testmode__6C190EBB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__klix__testmode__6C190EBB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[klix] ADD  CONSTRAINT [DF__klix__testmode__6C190EBB]  DEFAULT ((0)) FOR [testmode]
END


END
GO
/****** Object:  Default [DF__ITransact__id__6FE99F9F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ITransact__id__6FE99F9F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ITransact__id__6FE99F9F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ITransact] ADD  CONSTRAINT [DF__ITransact__id__6FE99F9F]  DEFAULT ((0)) FOR [id]
END


END
GO
/****** Object:  Default [DF__ITransact__it_am__70DDC3D8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ITransact__it_am__70DDC3D8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ITransact__it_am__70DDC3D8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ITransact] ADD  CONSTRAINT [DF__ITransact__it_am__70DDC3D8]  DEFAULT ((0)) FOR [it_amex]
END


END
GO
/****** Object:  Default [DF__ITransact__it_di__71D1E811]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ITransact__it_di__71D1E811]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ITransact__it_di__71D1E811]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ITransact] ADD  CONSTRAINT [DF__ITransact__it_di__71D1E811]  DEFAULT ((0)) FOR [it_diner]
END


END
GO
/****** Object:  Default [DF__ITransact__it_di__72C60C4A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ITransact__it_di__72C60C4A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ITransact__it_di__72C60C4A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ITransact] ADD  CONSTRAINT [DF__ITransact__it_di__72C60C4A]  DEFAULT ((0)) FOR [it_disc]
END


END
GO
/****** Object:  Default [DF__ITransact__it_mc__73BA3083]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ITransact__it_mc__73BA3083]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ITransact__it_mc__73BA3083]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ITransact] ADD  CONSTRAINT [DF__ITransact__it_mc__73BA3083]  DEFAULT ((0)) FOR [it_mc]
END


END
GO
/****** Object:  Default [DF__ITransact__it_vi__74AE54BC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__ITransact__it_vi__74AE54BC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__ITransact__it_vi__74AE54BC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ITransact] ADD  CONSTRAINT [DF__ITransact__it_vi__74AE54BC]  DEFAULT ((0)) FOR [it_visa]
END


END
GO
/****** Object:  Default [DBX_ReqCVV_24858]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_ReqCVV_24858]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_ReqCVV_24858]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ITransact] ADD  CONSTRAINT [DBX_ReqCVV_24858]  DEFAULT ((0)) FOR [ReqCVV]
END


END
GO
/****** Object:  Default [DBX_TransType_31671]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_TransType_31671]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_TransType_31671]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[ITransact] ADD  CONSTRAINT [DBX_TransType_31671]  DEFAULT ((0)) FOR [TransType]
END


END
GO
/****** Object:  Default [DBX_id_23731]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_id_23731]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_id_23731]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[icons] ADD  CONSTRAINT [DBX_id_23731]  DEFAULT ((0)) FOR [id]
END


END
GO
/****** Object:  Default [DF_icons_arrowUp]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_icons_arrowUp]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_icons_arrowUp]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[icons] ADD  CONSTRAINT [DF_icons_arrowUp]  DEFAULT ('images/sample/up-arrow.gif') FOR [arrowUp]
END


END
GO
/****** Object:  Default [DF_icons_arrowDown]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_icons_arrowDown]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_icons_arrowDown]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[icons] ADD  CONSTRAINT [DF_icons_arrowDown]  DEFAULT ('images/sample/down-arrow.gif') FOR [arrowDown]
END


END
GO
/****** Object:  Default [DF__FlatShipTyp__WQP__690797E6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__FlatShipTyp__WQP__690797E6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__FlatShipTyp__WQP__690797E6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[FlatShipTypes] ADD  CONSTRAINT [DF__FlatShipTyp__WQP__690797E6]  DEFAULT ('Q') FOR [WQP]
END


END
GO
/****** Object:  Default [DF__FlatShipT__idFla__00AA174D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__FlatShipT__idFla__00AA174D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__FlatShipT__idFla__00AA174D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[FlatShipTypeRules] ADD  CONSTRAINT [DF__FlatShipT__idFla__00AA174D]  DEFAULT ((0)) FOR [idFlatshipType]
END


END
GO
/****** Object:  Default [DF__FlatShipT__quant__019E3B86]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__FlatShipT__quant__019E3B86]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__FlatShipT__quant__019E3B86]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[FlatShipTypeRules] ADD  CONSTRAINT [DF__FlatShipT__quant__019E3B86]  DEFAULT ((0)) FOR [quantityFrom]
END


END
GO
/****** Object:  Default [DF__FlatShipT__quant__02925FBF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__FlatShipT__quant__02925FBF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__FlatShipT__quant__02925FBF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[FlatShipTypeRules] ADD  CONSTRAINT [DF__FlatShipT__quant__02925FBF]  DEFAULT ((0)) FOR [quantityTo]
END


END
GO
/****** Object:  Default [DF__FlatShipT__shipp__038683F8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__FlatShipT__shipp__038683F8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__FlatShipT__shipp__038683F8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[FlatShipTypeRules] ADD  CONSTRAINT [DF__FlatShipT__shipp__038683F8]  DEFAULT ((0)) FOR [shippingPrice]
END


END
GO
/****** Object:  Default [DF__FlatShipTyp__num__047AA831]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__FlatShipTyp__num__047AA831]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__FlatShipTyp__num__047AA831]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[FlatShipTypeRules] ADD  CONSTRAINT [DF__FlatShipTyp__num__047AA831]  DEFAULT ((0)) FOR [num]
END


END
GO
/****** Object:  Default [DBX_CVV2_15956]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_CVV2_15956]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_CVV2_15956]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[fasttransact] ADD  CONSTRAINT [DBX_CVV2_15956]  DEFAULT ((0)) FOR [CVV2]
END


END
GO
/****** Object:  Default [DBX_eWayID_6896]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_eWayID_6896]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_eWayID_6896]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[eWay] ADD  CONSTRAINT [DBX_eWayID_6896]  DEFAULT ((0)) FOR [eWayID]
END


END
GO
/****** Object:  Default [DBX_eWayTestmode_21251]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_eWayTestmode_21251]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_eWayTestmode_21251]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[eWay] ADD  CONSTRAINT [DBX_eWayTestmode_21251]  DEFAULT ((0)) FOR [eWayTestmode]
END


END
GO
/****** Object:  Default [DBX_eWayCVV_3037]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_eWayCVV_3037]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_eWayCVV_3037]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[eWay] ADD  CONSTRAINT [DBX_eWayCVV_3037]  DEFAULT ((0)) FOR [eWayCVV]
END


END
GO
/****** Object:  Default [DBX_eWayBeagleActive_3037]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_eWayBeagleActive_3037]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_eWayBeagleActive_3037]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[eWay] ADD  CONSTRAINT [DBX_eWayBeagleActive_3037]  DEFAULT ((0)) FOR [eWayBeagleActive]
END


END
GO
/****** Object:  Default [DBX_id_23014]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_id_23014]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_id_23014]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[echo] ADD  CONSTRAINT [DBX_id_23014]  DEFAULT ((0)) FOR [id]
END


END
GO
/****** Object:  Default [DBX_cnp_security_9458]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_cnp_security_9458]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_cnp_security_9458]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[echo] ADD  CONSTRAINT [DBX_cnp_security_9458]  DEFAULT ((0)) FOR [cnp_security]
END


END
GO
/****** Object:  Default [DF_idProduct_DProducts]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_idProduct_DProducts]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_idProduct_DProducts]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[DProducts] ADD  CONSTRAINT [DF_idProduct_DProducts]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF_URLExpire_DProducts]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_URLExpire_DProducts]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_URLExpire_DProducts]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[DProducts] ADD  CONSTRAINT [DF_URLExpire_DProducts]  DEFAULT ((0)) FOR [URLExpire]
END


END
GO
/****** Object:  Default [DF_ExpireDays_DProducts]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_ExpireDays_DProducts]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_ExpireDays_DProducts]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[DProducts] ADD  CONSTRAINT [DF_ExpireDays_DProducts]  DEFAULT ((0)) FOR [ExpireDays]
END


END
GO
/****** Object:  Default [DF_License_DProducts]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_License_DProducts]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_License_DProducts]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[DProducts] ADD  CONSTRAINT [DF_License_DProducts]  DEFAULT ((0)) FOR [License]
END


END
GO
/****** Object:  Default [DBX_IdOrder_32071]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_IdOrder_32071]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_IdOrder_32071]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[DPRequests] ADD  CONSTRAINT [DBX_IdOrder_32071]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DBX_IdProduct_12939]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_IdProduct_12939]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_IdProduct_12939]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[DPRequests] ADD  CONSTRAINT [DBX_IdProduct_12939]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DBX_IdCustomer_34]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_IdCustomer_34]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_IdCustomer_34]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[DPRequests] ADD  CONSTRAINT [DBX_IdCustomer_34]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DBX_idorder_14900]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idorder_14900]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idorder_14900]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[DPLicenses] ADD  CONSTRAINT [DBX_idorder_14900]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DBX_idproduct_8882]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idproduct_8882]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idproduct_8882]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[DPLicenses] ADD  CONSTRAINT [DBX_idproduct_8882]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__discounts__idpro__753864A1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__idpro__753864A1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__idpro__753864A1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discountsPerQuantity] ADD  CONSTRAINT [DF__discounts__idpro__753864A1]  DEFAULT ((0)) FOR [idproduct]
END


END
GO
/****** Object:  Default [DF__discounts__idcat__762C88DA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__idcat__762C88DA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__idcat__762C88DA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discountsPerQuantity] ADD  CONSTRAINT [DF__discounts__idcat__762C88DA]  DEFAULT ((0)) FOR [idcategory]
END


END
GO
/****** Object:  Default [DF__discounts__quant__7720AD13]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__quant__7720AD13]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__quant__7720AD13]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discountsPerQuantity] ADD  CONSTRAINT [DF__discounts__quant__7720AD13]  DEFAULT ((0)) FOR [quantityFrom]
END


END
GO
/****** Object:  Default [DF__discounts__quant__7814D14C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__quant__7814D14C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__quant__7814D14C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discountsPerQuantity] ADD  CONSTRAINT [DF__discounts__quant__7814D14C]  DEFAULT ((0)) FOR [quantityUntil]
END


END
GO
/****** Object:  Default [DF__discounts__disco__7908F585]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__disco__7908F585]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__disco__7908F585]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discountsPerQuantity] ADD  CONSTRAINT [DF__discounts__disco__7908F585]  DEFAULT ((0)) FOR [discountPerUnit]
END


END
GO
/****** Object:  Default [DF__discountsPe__num__79FD19BE]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discountsPe__num__79FD19BE]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discountsPe__num__79FD19BE]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discountsPerQuantity] ADD  CONSTRAINT [DF__discountsPe__num__79FD19BE]  DEFAULT ((0)) FOR [num]
END


END
GO
/****** Object:  Default [DF__discounts__perce__7AF13DF7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__perce__7AF13DF7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__perce__7AF13DF7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discountsPerQuantity] ADD  CONSTRAINT [DF__discounts__perce__7AF13DF7]  DEFAULT ((0)) FOR [percentage]
END


END
GO
/****** Object:  Default [DF__discounts__disco__7BE56230]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__disco__7BE56230]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__disco__7BE56230]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discountsPerQuantity] ADD  CONSTRAINT [DF__discounts__disco__7BE56230]  DEFAULT ((0)) FOR [discountPerWUnit]
END


END
GO
/****** Object:  Default [DF_baseproductonly_discountsPerQuantity]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_baseproductonly_discountsPerQuantity]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_baseproductonly_discountsPerQuantity]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discountsPerQuantity] ADD  CONSTRAINT [DF_baseproductonly_discountsPerQuantity]  DEFAULT ((0)) FOR [baseproductonly]
END


END
GO
/****** Object:  Default [DF__discounts__price__25DB9BFC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__price__25DB9BFC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__price__25DB9BFC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__price__25DB9BFC]  DEFAULT ((0)) FOR [pricetodiscount]
END


END
GO
/****** Object:  Default [DF__discounts__perce__26CFC035]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__perce__26CFC035]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__perce__26CFC035]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__perce__26CFC035]  DEFAULT ((0)) FOR [percentagetodiscount]
END


END
GO
/****** Object:  Default [DF__discounts__activ__27C3E46E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__activ__27C3E46E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__activ__27C3E46E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__activ__27C3E46E]  DEFAULT ((0)) FOR [active]
END


END
GO
/****** Object:  Default [DF__discounts__used__28B808A7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__used__28B808A7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__used__28B808A7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__used__28B808A7]  DEFAULT ((0)) FOR [used]
END


END
GO
/****** Object:  Default [DF__discounts__oneti__29AC2CE0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__oneti__29AC2CE0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__oneti__29AC2CE0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__oneti__29AC2CE0]  DEFAULT ((0)) FOR [onetime]
END


END
GO
/****** Object:  Default [DF__discounts__quant__2AA05119]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__quant__2AA05119]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__quant__2AA05119]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__quant__2AA05119]  DEFAULT ((0)) FOR [quantityfrom]
END


END
GO
/****** Object:  Default [DF__discounts__quant__2B947552]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__quant__2B947552]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__quant__2B947552]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__quant__2B947552]  DEFAULT ((0)) FOR [quantityuntil]
END


END
GO
/****** Object:  Default [DF__discounts__weigh__2C88998B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__weigh__2C88998B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__weigh__2C88998B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__weigh__2C88998B]  DEFAULT ((0)) FOR [weightfrom]
END


END
GO
/****** Object:  Default [DF__discounts__weigh__2D7CBDC4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__weigh__2D7CBDC4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__weigh__2D7CBDC4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__weigh__2D7CBDC4]  DEFAULT ((0)) FOR [weightuntil]
END


END
GO
/****** Object:  Default [DF__discounts__price__2E70E1FD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__price__2E70E1FD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__price__2E70E1FD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__price__2E70E1FD]  DEFAULT ((0)) FOR [pricefrom]
END


END
GO
/****** Object:  Default [DF__discounts__price__2F650636]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__price__2F650636]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__price__2F650636]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__price__2F650636]  DEFAULT ((0)) FOR [priceuntil]
END


END
GO
/****** Object:  Default [DF__discounts__idPro__30592A6F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__idPro__30592A6F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__idPro__30592A6F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__idPro__30592A6F]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__discounts__pcSep__314D4EA8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__discounts__pcSep__314D4EA8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__discounts__pcSep__314D4EA8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF__discounts__pcSep__314D4EA8]  DEFAULT ((0)) FOR [pcSeparate]
END


END
GO
/****** Object:  Default [DF_discounts_pcDisc_Auto]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_discounts_pcDisc_Auto]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_discounts_pcDisc_Auto]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF_discounts_pcDisc_Auto]  DEFAULT ((0)) FOR [pcDisc_Auto]
END


END
GO
/****** Object:  Default [DBX_pcDisc_PerToFlatCartTotal_26371]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcDisc_PerToFlatCartTotal_26371]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcDisc_PerToFlatCartTotal_26371]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DBX_pcDisc_PerToFlatCartTotal_26371]  DEFAULT ((0)) FOR [pcDisc_PerToFlatCartTotal]
END


END
GO
/****** Object:  Default [DBX_pcRetailFlag_12139]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcRetailFlag_12139]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcRetailFlag_12139]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DBX_pcRetailFlag_12139]  DEFAULT ((0)) FOR [pcRetailFlag]
END


END
GO
/****** Object:  Default [DBX_pcDisc_PerToFlatDiscount_31801]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcDisc_PerToFlatDiscount_31801]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcDisc_PerToFlatDiscount_31801]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DBX_pcDisc_PerToFlatDiscount_31801]  DEFAULT ((0)) FOR [pcDisc_PerToFlatDiscount]
END


END
GO
/****** Object:  Default [DBX_pcWholesaleFlag_27310]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcWholesaleFlag_27310]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcWholesaleFlag_27310]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DBX_pcWholesaleFlag_27310]  DEFAULT ((0)) FOR [pcWholesaleFlag]
END


END
GO
/****** Object:  Default [DF_discounts_pcDisc_IncExcPrd]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_discounts_pcDisc_IncExcPrd]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_discounts_pcDisc_IncExcPrd]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF_discounts_pcDisc_IncExcPrd]  DEFAULT ((0)) FOR [pcDisc_IncExcPrd]
END


END
GO
/****** Object:  Default [DF_discounts_pcDisc_IncExcCat]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_discounts_pcDisc_IncExcCat]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_discounts_pcDisc_IncExcCat]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF_discounts_pcDisc_IncExcCat]  DEFAULT ((0)) FOR [pcDisc_IncExcCat]
END


END
GO
/****** Object:  Default [DF_discounts_pcDisc_IncExcCust]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_discounts_pcDisc_IncExcCust]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_discounts_pcDisc_IncExcCust]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF_discounts_pcDisc_IncExcCust]  DEFAULT ((0)) FOR [pcDisc_IncExcCust]
END


END
GO
/****** Object:  Default [DF_discounts_pcDisc_IncExcCPrice]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_discounts_pcDisc_IncExcCPrice]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_discounts_pcDisc_IncExcCPrice]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[discounts] ADD  CONSTRAINT [DF_discounts_pcDisc_IncExcCPrice]  DEFAULT ((0)) FOR [pcDisc_IncExcCPrice]
END


END
GO
/****** Object:  Default [DF__customfie__searc__361203C5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customfie__searc__361203C5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customfie__searc__361203C5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customfields] ADD  CONSTRAINT [DF__customfie__searc__361203C5]  DEFAULT ((1)) FOR [searchable]
END


END
GO
/****** Object:  Default [DF__customers__custo__39E294A9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__custo__39E294A9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__custo__39E294A9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__custo__39E294A9]  DEFAULT ((0)) FOR [customerType]
END


END
GO
/****** Object:  Default [DF__customers__Total__3AD6B8E2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__Total__3AD6B8E2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__Total__3AD6B8E2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__Total__3AD6B8E2]  DEFAULT ((0)) FOR [TotalOrders]
END


END
GO
/****** Object:  Default [DF__customers__Total__3BCADD1B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__Total__3BCADD1B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__Total__3BCADD1B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__Total__3BCADD1B]  DEFAULT ((0)) FOR [TotalSales]
END


END
GO
/****** Object:  Default [DF__customers__iRewa__3CBF0154]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__iRewa__3CBF0154]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__iRewa__3CBF0154]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__iRewa__3CBF0154]  DEFAULT ((0)) FOR [iRewardPointsAccrued]
END


END
GO
/****** Object:  Default [DF__customers__iRewa__3DB3258D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__iRewa__3DB3258D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__iRewa__3DB3258D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__iRewa__3DB3258D]  DEFAULT ((0)) FOR [iRewardPointsUsed]
END


END
GO
/****** Object:  Default [DF__customers__iRewa__3EA749C6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__iRewa__3EA749C6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__iRewa__3EA749C6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__iRewa__3EA749C6]  DEFAULT ((0)) FOR [iRewardPointsHistory]
END


END
GO
/****** Object:  Default [DF__customers__iRewa__3F9B6DFF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__iRewa__3F9B6DFF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__iRewa__3F9B6DFF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__iRewa__3F9B6DFF]  DEFAULT ((0)) FOR [iRewardPointsHistoryUsed]
END


END
GO
/****** Object:  Default [DF__customers__RecvN__408F9238]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__RecvN__408F9238]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__RecvN__408F9238]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__RecvN__408F9238]  DEFAULT ((0)) FOR [RecvNews]
END


END
GO
/****** Object:  Default [DF__customers__suspe__4183B671]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__suspe__4183B671]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__suspe__4183B671]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__suspe__4183B671]  DEFAULT ((0)) FOR [suspend]
END


END
GO
/****** Object:  Default [DF__customers__idCus__4277DAAA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customers__idCus__4277DAAA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customers__idCus__4277DAAA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF__customers__idCus__4277DAAA]  DEFAULT ((0)) FOR [idCustomerCategory]
END


END
GO
/****** Object:  Default [DBX_pcCust_Locked_4570]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcCust_Locked_4570]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcCust_Locked_4570]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DBX_pcCust_Locked_4570]  DEFAULT ((0)) FOR [pcCust_Locked]
END


END
GO
/****** Object:  Default [DF_customers_pcCust_Guest]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_customers_pcCust_Guest]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_customers_pcCust_Guest]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF_customers_pcCust_Guest]  DEFAULT ((0)) FOR [pcCust_Guest]
END


END
GO
/****** Object:  Default [DF_customers_pcCust_Residential]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_customers_pcCust_Residential]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_customers_pcCust_Residential]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF_customers_pcCust_Residential]  DEFAULT ((1)) FOR [pcCust_Residential]
END


END
GO
/****** Object:  Default [DF_customers_pcCust_AgreeTerms]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_customers_pcCust_AgreeTerms]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_customers_pcCust_AgreeTerms]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF_customers_pcCust_AgreeTerms]  DEFAULT ((0)) FOR [pcCust_AgreeTerms]
END


END
GO
/****** Object:  Default [DF_customers_pcCust_AllowReviewEmails]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_customers_pcCust_AllowReviewEmails]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_customers_pcCust_AllowReviewEmails]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customers] ADD  CONSTRAINT [DF_customers_pcCust_AllowReviewEmails]  DEFAULT ((1)) FOR [pcCust_AllowReviewEmails]
END


END
GO
/****** Object:  Default [DF__customCar__idCus__4B0D20AB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customCar__idCus__4B0D20AB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customCar__idCus__4B0D20AB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customCardRules] ADD  CONSTRAINT [DF__customCar__idCus__4B0D20AB]  DEFAULT ((0)) FOR [idCustomCardType]
END


END
GO
/****** Object:  Default [DF__customCar__intru__4C0144E4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customCar__intru__4C0144E4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customCar__intru__4C0144E4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customCardRules] ADD  CONSTRAINT [DF__customCar__intru__4C0144E4]  DEFAULT ((0)) FOR [intruleRequired]
END


END
GO
/****** Object:  Default [DF__customCar__intle__4CF5691D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customCar__intle__4CF5691D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customCar__intle__4CF5691D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customCardRules] ADD  CONSTRAINT [DF__customCar__intle__4CF5691D]  DEFAULT ((0)) FOR [intlengthOfField]
END


END
GO
/****** Object:  Default [DF__customCar__intma__4DE98D56]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customCar__intma__4DE98D56]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customCar__intma__4DE98D56]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customCardRules] ADD  CONSTRAINT [DF__customCar__intma__4DE98D56]  DEFAULT ((0)) FOR [intmaxInput]
END


END
GO
/****** Object:  Default [DF__customCar__intOr__4EDDB18F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customCar__intOr__4EDDB18F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customCar__intOr__4EDDB18F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customCardRules] ADD  CONSTRAINT [DF__customCar__intOr__4EDDB18F]  DEFAULT ((0)) FOR [intOrder]
END


END
GO
/****** Object:  Default [DF__customCar__idOrd__53A266AC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customCar__idOrd__53A266AC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customCar__idOrd__53A266AC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customCardOrders] ADD  CONSTRAINT [DF__customCar__idOrd__53A266AC]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__customCar__idcus__54968AE5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customCar__idcus__54968AE5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customCar__idcus__54968AE5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customCardOrders] ADD  CONSTRAINT [DF__customCar__idcus__54968AE5]  DEFAULT ((0)) FOR [idcustomCardType]
END


END
GO
/****** Object:  Default [DF__customCar__idCus__558AAF1E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customCar__idCus__558AAF1E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customCar__idCus__558AAF1E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customCardOrders] ADD  CONSTRAINT [DF__customCar__idCus__558AAF1E]  DEFAULT ((0)) FOR [idCustomCardRules]
END


END
GO
/****** Object:  Default [DF__customCar__intOr__567ED357]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__customCar__intOr__567ED357]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__customCar__intOr__567ED357]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[customCardOrders] ADD  CONSTRAINT [DF__customCar__intOr__567ED357]  DEFAULT ((0)) FOR [intOrderTotal]
END


END
GO
/****** Object:  Default [DBX_idproduct_1794]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idproduct_1794]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idproduct_1794]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[cs_relationships] ADD  CONSTRAINT [DBX_idproduct_1794]  DEFAULT ((0)) FOR [idproduct]
END


END
GO
/****** Object:  Default [DBX_idrelation_26496]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idrelation_26496]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idrelation_26496]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[cs_relationships] ADD  CONSTRAINT [DBX_idrelation_26496]  DEFAULT ((0)) FOR [idrelation]
END


END
GO
/****** Object:  Default [DF__cs_relation__num__5B438874]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__cs_relation__num__5B438874]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__cs_relation__num__5B438874]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[cs_relationships] ADD  CONSTRAINT [DF__cs_relation__num__5B438874]  DEFAULT ((0)) FOR [num]
END


END
GO
/****** Object:  Default [DF__cs_relati__disco__5C37ACAD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__cs_relati__disco__5C37ACAD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__cs_relati__disco__5C37ACAD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[cs_relationships] ADD  CONSTRAINT [DF__cs_relati__disco__5C37ACAD]  DEFAULT ((0)) FOR [discount]
END


END
GO
/****** Object:  Default [DF__cs_relati__isPer__5D2BD0E6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__cs_relati__isPer__5D2BD0E6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__cs_relati__isPer__5D2BD0E6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[cs_relationships] ADD  CONSTRAINT [DF__cs_relati__isPer__5D2BD0E6]  DEFAULT ((0)) FOR [isPercent]
END


END
GO
/****** Object:  Default [DF__cs_relati__isReq__5E1FF51F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__cs_relati__isReq__5E1FF51F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__cs_relati__isReq__5E1FF51F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[cs_relationships] ADD  CONSTRAINT [DF__cs_relati__isReq__5E1FF51F]  DEFAULT ((0)) FOR [isRequired]
END


END
GO
/****** Object:  Default [DF__cs_relati__cs_ty__5F141958]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__cs_relati__cs_ty__5F141958]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__cs_relati__cs_ty__5F141958]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[cs_relationships] ADD  CONSTRAINT [DF__cs_relati__cs_ty__5F141958]  DEFAULT ('Accessory') FOR [cs_type]
END


END
GO
/****** Object:  Default [DBX_id_5105]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_id_5105]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_id_5105]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[crossSelldata] ADD  CONSTRAINT [DBX_id_5105]  DEFAULT ((0)) FOR [id]
END


END
GO
/****** Object:  Default [DF__crossSell__cs_st__41B8C09B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__crossSell__cs_st__41B8C09B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__crossSell__cs_st__41B8C09B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[crossSelldata] ADD  CONSTRAINT [DF__crossSell__cs_st__41B8C09B]  DEFAULT ((0)) FOR [cs_status]
END


END
GO
/****** Object:  Default [DF__crossSell__cs_sh__42ACE4D4]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__crossSell__cs_sh__42ACE4D4]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__crossSell__cs_sh__42ACE4D4]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[crossSelldata] ADD  CONSTRAINT [DF__crossSell__cs_sh__42ACE4D4]  DEFAULT ((0)) FOR [cs_showprod]
END


END
GO
/****** Object:  Default [DF__crossSell__cs_sh__43A1090D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__crossSell__cs_sh__43A1090D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__crossSell__cs_sh__43A1090D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[crossSelldata] ADD  CONSTRAINT [DF__crossSell__cs_sh__43A1090D]  DEFAULT ((0)) FOR [cs_showcart]
END


END
GO
/****** Object:  Default [DF__crossSell__cs_sh__44952D46]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__crossSell__cs_sh__44952D46]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__crossSell__cs_sh__44952D46]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[crossSelldata] ADD  CONSTRAINT [DF__crossSell__cs_sh__44952D46]  DEFAULT ((0)) FOR [cs_showimage]
END


END
GO
/****** Object:  Default [DF__crossSell__cs_Pr__4589517F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__crossSell__cs_Pr__4589517F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__crossSell__cs_Pr__4589517F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[crossSelldata] ADD  CONSTRAINT [DF__crossSell__cs_Pr__4589517F]  DEFAULT ((0)) FOR [cs_ProductViewCnt]
END


END
GO
/****** Object:  Default [DF__crossSell__cs_Ca__467D75B8]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__crossSell__cs_Ca__467D75B8]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__crossSell__cs_Ca__467D75B8]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[crossSelldata] ADD  CONSTRAINT [DF__crossSell__cs_Ca__467D75B8]  DEFAULT ((0)) FOR [cs_CartViewCnt]
END


END
GO
/****** Object:  Default [DF__crossSell__cs_Im__477199F1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__crossSell__cs_Im__477199F1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__crossSell__cs_Im__477199F1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[crossSelldata] ADD  CONSTRAINT [DF__crossSell__cs_Im__477199F1]  DEFAULT ((0)) FOR [cs_ImageHeight]
END


END
GO
/****** Object:  Default [DF__crossSell__cs_Im__4865BE2A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__crossSell__cs_Im__4865BE2A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__crossSell__cs_Im__4865BE2A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[crossSelldata] ADD  CONSTRAINT [DF__crossSell__cs_Im__4865BE2A]  DEFAULT ((0)) FOR [cs_ImageWidth]
END


END
GO
/****** Object:  Default [DF__creditCar__idOrd__220B0B18]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__creditCar__idOrd__220B0B18]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__creditCar__idOrd__220B0B18]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[creditCards] ADD  CONSTRAINT [DF__creditCar__idOrd__220B0B18]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF_creditCards_pcSecurityKeyID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_creditCards_pcSecurityKeyID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_creditCards_pcSecurityKeyID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[creditCards] ADD  CONSTRAINT [DF_creditCards_pcSecurityKeyID]  DEFAULT ((0)) FOR [pcSecurityKeyID]
END


END
GO
/****** Object:  Default [DBX_pcSubDivisionID_20378]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcSubDivisionID_20378]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcSubDivisionID_20378]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[countries] ADD  CONSTRAINT [DBX_pcSubDivisionID_20378]  DEFAULT ((0)) FOR [pcSubDivisionID]
END


END
GO
/****** Object:  Default [DF__configWis__dPric__7C1A6C5A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configWis__dPric__7C1A6C5A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configWis__dPric__7C1A6C5A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configWishlistSessions] ADD  CONSTRAINT [DF__configWis__dPric__7C1A6C5A]  DEFAULT ((0)) FOR [dPrice]
END


END
GO
/****** Object:  Default [DF__configWis__pccon__7D0E9093]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configWis__pccon__7D0E9093]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configWis__pccon__7D0E9093]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configWishlistSessions] ADD  CONSTRAINT [DF__configWis__pccon__7D0E9093]  DEFAULT ((0)) FOR [pcconf_Quantity]
END


END
GO
/****** Object:  Default [DF__configWis__pccon__7E02B4CC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configWis__pccon__7E02B4CC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configWis__pccon__7E02B4CC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configWishlistSessions] ADD  CONSTRAINT [DF__configWis__pccon__7E02B4CC]  DEFAULT ((0)) FOR [pcconf_QDiscount]
END


END
GO
/****** Object:  Default [DBX_specProduct_17660]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_specProduct_17660]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_specProduct_17660]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DBX_specProduct_17660]  DEFAULT ((0)) FOR [specProduct]
END


END
GO
/****** Object:  Default [DBX_configProduct_22248]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_configProduct_22248]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_configProduct_22248]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DBX_configProduct_22248]  DEFAULT ((0)) FOR [configProduct]
END


END
GO
/****** Object:  Default [DBX_price_15772]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_price_15772]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_price_15772]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DBX_price_15772]  DEFAULT ((0)) FOR [price]
END


END
GO
/****** Object:  Default [DBX_Wprice_4972]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_Wprice_4972]') AND type = 'D')

BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_Wprice_4972]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DBX_Wprice_4972]  DEFAULT ((0)) FOR [Wprice]
END


END
GO
/****** Object:  Default [DF__configSpe__cdefa__01D345B0]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__cdefa__01D345B0]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__cdefa__01D345B0]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__cdefa__01D345B0]  DEFAULT ((0)) FOR [cdefault]
END


END
GO
/****** Object:  Default [DF__configSpe__showI__02C769E9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__showI__02C769E9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__showI__02C769E9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__showI__02C769E9]  DEFAULT ((0)) FOR [showInfo]
END


END
GO
/****** Object:  Default [DF__configSpe__requi__03BB8E22]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__requi__03BB8E22]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__requi__03BB8E22]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__requi__03BB8E22]  DEFAULT ((0)) FOR [requiredCategory]
END


END
GO
/****** Object:  Default [DF__configSpe__multi__04AFB25B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__multi__04AFB25B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__multi__04AFB25B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__multi__04AFB25B]  DEFAULT ((0)) FOR [multiSelect]
END


END
GO
/****** Object:  Default [DF__configSpe__prdSo__05A3D694]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__prdSo__05A3D694]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__prdSo__05A3D694]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__prdSo__05A3D694]  DEFAULT ((0)) FOR [prdSort]
END


END
GO
/****** Object:  Default [DF__configSpe__catSo__0697FACD]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__catSo__0697FACD]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__catSo__0697FACD]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__catSo__0697FACD]  DEFAULT ((0)) FOR [catSort]
END


END
GO
/****** Object:  Default [DF__configSpe__confi__078C1F06]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__confi__078C1F06]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__confi__078C1F06]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__confi__078C1F06]  DEFAULT ((0)) FOR [configProductCategory]
END


END
GO
/****** Object:  Default [DF__configSpe__displ__0880433F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__displ__0880433F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__displ__0880433F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__displ__0880433F]  DEFAULT ((0)) FOR [displayQF]
END


END
GO
/****** Object:  Default [DF__configSpe__pcCon__09746778]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__pcCon__09746778]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__pcCon__09746778]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__pcCon__09746778]  DEFAULT ((0)) FOR [pcConfPro_ShowImg]
END


END
GO
/****** Object:  Default [DF__configSpe__pcCon__0A688BB1]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__pcCon__0A688BB1]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__pcCon__0A688BB1]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__pcCon__0A688BB1]  DEFAULT ((0)) FOR [pcConfPro_ImgWidth]
END


END
GO
/****** Object:  Default [DF__configSpe__pcCon__0B5CAFEA]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__pcCon__0B5CAFEA]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__pcCon__0B5CAFEA]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__pcCon__0B5CAFEA]  DEFAULT ((0)) FOR [pcConfPro_ShowSKU]
END


END
GO
/****** Object:  Default [DF__configSpe__pcCon__0C50D423]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__pcCon__0C50D423]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__pcCon__0C50D423]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DF__configSpe__pcCon__0C50D423]  DEFAULT ((0)) FOR [pcConfPro_ShowDesc]
END


END
GO
/****** Object:  Default [DBX_pcConfPro_UseRadio_18108]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_pcConfPro_UseRadio_18108]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_pcConfPro_UseRadio_18108]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_products] ADD  CONSTRAINT [DBX_pcConfPro_UseRadio_18108]  DEFAULT ((0)) FOR [pcConfPro_UseRadio]
END


END
GO
/****** Object:  Default [DF__configSpe__cdefa__10216507]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__cdefa__10216507]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__cdefa__10216507]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__cdefa__10216507]  DEFAULT ((0)) FOR [cdefault]
END


END
GO
/****** Object:  Default [DF__configSpe__showI__11158940]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__showI__11158940]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__showI__11158940]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__showI__11158940]  DEFAULT ((0)) FOR [showInfo]
END


END
GO
/****** Object:  Default [DF__configSpe__requi__1209AD79]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__requi__1209AD79]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__requi__1209AD79]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__requi__1209AD79]  DEFAULT ((0)) FOR [requiredCategory]
END


END
GO
/****** Object:  Default [DF__configSpe__multi__12FDD1B2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__multi__12FDD1B2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__multi__12FDD1B2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__multi__12FDD1B2]  DEFAULT ((0)) FOR [multiSelect]
END


END
GO
/****** Object:  Default [DF__configSpe__displ__13F1F5EB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__displ__13F1F5EB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__displ__13F1F5EB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__displ__13F1F5EB]  DEFAULT ((0)) FOR [displayQF]
END


END
GO
/****** Object:  Default [DF__configSpe__pcCon__14E61A24]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__pcCon__14E61A24]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__pcCon__14E61A24]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__pcCon__14E61A24]  DEFAULT ((0)) FOR [pcConfCha_ShowImg]
END


END
GO
/****** Object:  Default [DF__configSpe__pcCon__15DA3E5D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__pcCon__15DA3E5D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__pcCon__15DA3E5D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__pcCon__15DA3E5D]  DEFAULT ((0)) FOR [pcConfCha_ImgWidth]
END


END
GO
/****** Object:  Default [DF__configSpe__pcCon__16CE6296]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__pcCon__16CE6296]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__pcCon__16CE6296]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__pcCon__16CE6296]  DEFAULT ((0)) FOR [pcConfCha_ShowSKU]
END


END
GO
/****** Object:  Default [DF__configSpe__pcCon__17C286CF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__pcCon__17C286CF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__pcCon__17C286CF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__pcCon__17C286CF]  DEFAULT ((0)) FOR [pcConfCha_ShowDesc]
END


END
GO
/****** Object:  Default [DF__configSpe__pcCon__17C286CG]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__pcCon__17C286CG]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__pcCon__17C286CG]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_Charges] ADD  CONSTRAINT [DF__configSpe__pcCon__17C286CG]  DEFAULT ((0)) FOR [pcConfCha_UseRadio]
END


END
GO
/****** Object:  Default [DBX_idProduct_30896]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idProduct_30896]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idProduct_30896]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_categories] ADD  CONSTRAINT [DBX_idProduct_30896]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DBX_idCategory_16291]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idCategory_16291]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idCategory_16291]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_categories] ADD  CONSTRAINT [DBX_idCategory_16291]  DEFAULT ((0)) FOR [idCategory]
END


END
GO
/****** Object:  Default [DF__configSpe__catOr__1B9317B3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__configSpe__catOr__1B9317B3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__configSpe__catOr__1B9317B3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSpec_categories] ADD  CONSTRAINT [DF__configSpe__catOr__1B9317B3]  DEFAULT ((0)) FOR [catOrder]
END


END
GO
/****** Object:  Default [DBX_configKey_30283]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_configKey_30283]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_configKey_30283]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSessions] ADD  CONSTRAINT [DBX_configKey_30283]  DEFAULT ((0)) FOR [configKey]
END


END
GO
/****** Object:  Default [DBX_idproduct_18726]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idproduct_18726]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idproduct_18726]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[configSessions] ADD  CONSTRAINT [DBX_idproduct_18726]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__concord__CVV__22401542]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__concord__CVV__22401542]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__concord__CVV__22401542]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[concord] ADD  CONSTRAINT [DF__concord__CVV__22401542]  DEFAULT ((0)) FOR [CVV]
END


END
GO
/****** Object:  Default [DF__concord__testmod__2334397B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__concord__testmod__2334397B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__concord__testmod__2334397B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[concord] ADD  CONSTRAINT [DF__concord__testmod__2334397B]  DEFAULT ((0)) FOR [testmode]
END


END
GO
/****** Object:  Default [DF__CCTypes__active__2704CA5F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__CCTypes__active__2704CA5F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__CCTypes__active__2704CA5F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[CCTypes] ADD  CONSTRAINT [DF__CCTypes__active__2704CA5F]  DEFAULT ((0)) FOR [active]
END


END
GO
/****** Object:  Default [DF__categorie__idPro__6E8B6712]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__idPro__6E8B6712]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__idPro__6E8B6712]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories_products] ADD  CONSTRAINT [DF__categorie__idPro__6E8B6712]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__categorie__idCat__6F7F8B4B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__idCat__6F7F8B4B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__idCat__6F7F8B4B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories_products] ADD  CONSTRAINT [DF__categorie__idCat__6F7F8B4B]  DEFAULT ((0)) FOR [idCategory]
END


END
GO
/****** Object:  Default [DF__categorie__POrde__7073AF84]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__POrde__7073AF84]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__POrde__7073AF84]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories_products] ADD  CONSTRAINT [DF__categorie__POrde__7073AF84]  DEFAULT ((0)) FOR [POrder]
END


END
GO
/****** Object:  Default [DF__categorie__idPar__2EA5EC27]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__idPar__2EA5EC27]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__idPar__2EA5EC27]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__idPar__2EA5EC27]  DEFAULT ((0)) FOR [idParentCategory]
END


END
GO
/****** Object:  Default [DF__categories__tier__2F9A1060]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categories__tier__2F9A1060]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categories__tier__2F9A1060]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categories__tier__2F9A1060]  DEFAULT ((0)) FOR [tier]
END


END
GO
/****** Object:  Default [DF__categorie__servi__308E3499]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__servi__308E3499]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__servi__308E3499]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__servi__308E3499]  DEFAULT ((0)) FOR [serviceSpec]
END


END
GO
/****** Object:  Default [DF__categorie__requi__318258D2]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__requi__318258D2]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__requi__318258D2]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__requi__318258D2]  DEFAULT ((0)) FOR [required]
END


END
GO
/****** Object:  Default [DF__categorie__defin__32767D0B]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__defin__32767D0B]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__defin__32767D0B]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__defin__32767D0B]  DEFAULT ((0)) FOR [definePrd]
END


END
GO
/****** Object:  Default [DF__categorie__prior__336AA144]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__prior__336AA144]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__prior__336AA144]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__prior__336AA144]  DEFAULT ((0)) FOR [priority]
END


END
GO
/****** Object:  Default [DF__categorie__multi__345EC57D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__multi__345EC57D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__multi__345EC57D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__multi__345EC57D]  DEFAULT ((0)) FOR [multi]
END


END
GO
/****** Object:  Default [DBX_basePrice_6379]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_basePrice_6379]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_basePrice_6379]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DBX_basePrice_6379]  DEFAULT ((0)) FOR [basePrice]
END


END
GO
/****** Object:  Default [DF__categorie__iBTOh__3552E9B6]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__iBTOh__3552E9B6]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__iBTOh__3552E9B6]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__iBTOh__3552E9B6]  DEFAULT ((0)) FOR [iBTOhide]
END


END
GO
/****** Object:  Default [DBX_HideDesc_12021]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_HideDesc_12021]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_HideDesc_12021]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DBX_HideDesc_12021]  DEFAULT ((0)) FOR [HideDesc]
END


END
GO
/****** Object:  Default [DF__categorie__pcCat__36470DEF]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__pcCat__36470DEF]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__pcCat__36470DEF]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__pcCat__36470DEF]  DEFAULT ((0)) FOR [pcCats_RetailHide]
END


END
GO
/****** Object:  Default [DF__categorie__pcCat__373B3228]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__pcCat__373B3228]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__pcCat__373B3228]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__pcCat__373B3228]  DEFAULT ((0)) FOR [pcCats_SubCategoryView]
END


END
GO
/****** Object:  Default [DF__categorie__pcCat__382F5661]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__pcCat__382F5661]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__pcCat__382F5661]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__pcCat__382F5661]  DEFAULT ((0)) FOR [pcCats_CategoryColumns]
END


END
GO
/****** Object:  Default [DF__categorie__pcCat__39237A9A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__pcCat__39237A9A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__pcCat__39237A9A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__pcCat__39237A9A]  DEFAULT ((0)) FOR [pcCats_CategoryRows]
END


END
GO
/****** Object:  Default [DF__categorie__pcCat__3A179ED3]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__pcCat__3A179ED3]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__pcCat__3A179ED3]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__pcCat__3A179ED3]  DEFAULT ((0)) FOR [pcCats_ProductColumns]
END


END
GO
/****** Object:  Default [DF__categorie__pcCat__3B0BC30C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__pcCat__3B0BC30C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__pcCat__3B0BC30C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__pcCat__3B0BC30C]  DEFAULT ((0)) FOR [pcCats_ProductRows]
END


END
GO
/****** Object:  Default [DF__categorie__pcCat__3BFFE745]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__pcCat__3BFFE745]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__pcCat__3BFFE745]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__pcCat__3BFFE745]  DEFAULT ((0)) FOR [pcCats_FeaturedCategory]
END


END
GO
/****** Object:  Default [DF__categorie__pcCat__3CF40B7E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__categorie__pcCat__3CF40B7E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__categorie__pcCat__3CF40B7E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[categories] ADD  CONSTRAINT [DF__categorie__pcCat__3CF40B7E]  DEFAULT ((0)) FOR [pcCats_FeaturedCategoryImage]
END


END
GO
/****** Object:  Default [DF_Brands_pcBrands_SubBrandsView]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_Brands_pcBrands_SubBrandsView]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_Brands_pcBrands_SubBrandsView]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[Brands] ADD  CONSTRAINT [DF_Brands_pcBrands_SubBrandsView]  DEFAULT ((0)) FOR [pcBrands_SubBrandsView]
END


END
GO
/****** Object:  Default [DF_Brands_pcBrands_Active]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_Brands_pcBrands_Active]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_Brands_pcBrands_Active]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[Brands] ADD  CONSTRAINT [DF_Brands_pcBrands_Active]  DEFAULT ((1)) FOR [pcBrands_Active]
END


END
GO
/****** Object:  Default [DF_Brands_pcBrands_Order]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_Brands_pcBrands_Order]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_Brands_pcBrands_Order]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[Brands] ADD  CONSTRAINT [DF_Brands_pcBrands_Order]  DEFAULT ((0)) FOR [pcBrands_Order]
END


END
GO
/****** Object:  Default [DF_Brands_pcBrands_Parent]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_Brands_pcBrands_Parent]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_Brands_pcBrands_Parent]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[Brands] ADD  CONSTRAINT [DF_Brands_pcBrands_Parent]  DEFAULT ((0)) FOR [pcBrands_Parent]
END


END
GO
/****** Object:  Default [DBX_idBluePay_12384]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idBluePay_12384]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idBluePay_12384]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[BluePay] ADD  CONSTRAINT [DBX_idBluePay_12384]  DEFAULT ((0)) FOR [idBluePay]
END


END
GO
/****** Object:  Default [DBX_BPTestmode_5712]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_BPTestmode_5712]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_BPTestmode_5712]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[BluePay] ADD  CONSTRAINT [DBX_BPTestmode_5712]  DEFAULT ((0)) FOR [BPTestmode]
END


END
GO
/****** Object:  Default [DF__authorizeNet__id__59904A2C]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__authorizeNet__id__59904A2C]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__authorizeNet__id__59904A2C]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorizeNet] ADD  CONSTRAINT [DF__authorizeNet__id__59904A2C]  DEFAULT ((0)) FOR [id]
END


END
GO
/****** Object:  Default [DBX_x_version_134]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_x_version_134]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_x_version_134]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorizeNet] ADD  CONSTRAINT [DBX_x_version_134]  DEFAULT ((3.1)) FOR [x_version]
END


END
GO
/****** Object:  Default [DF__authorize__x_CVV__5A846E65]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__authorize__x_CVV__5A846E65]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__authorize__x_CVV__5A846E65]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorizeNet] ADD  CONSTRAINT [DF__authorize__x_CVV__5A846E65]  DEFAULT ((0)) FOR [x_CVV]
END


END
GO
/****** Object:  Default [DF__authorize__x_tes__5B78929E]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__authorize__x_tes__5B78929E]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__authorize__x_tes__5B78929E]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorizeNet] ADD  CONSTRAINT [DF__authorize__x_tes__5B78929E]  DEFAULT ((0)) FOR [x_testmode]
END


END
GO
/****** Object:  Default [DF__authorize__x_eCh__5C6CB6D7]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__authorize__x_eCh__5C6CB6D7]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__authorize__x_eCh__5C6CB6D7]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorizeNet] ADD  CONSTRAINT [DF__authorize__x_eCh__5C6CB6D7]  DEFAULT ((0)) FOR [x_eCheck]
END


END
GO
/****** Object:  Default [DF__authorize__x_sec__5D60DB10]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__authorize__x_sec__5D60DB10]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__authorize__x_sec__5D60DB10]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorizeNet] ADD  CONSTRAINT [DF__authorize__x_sec__5D60DB10]  DEFAULT ((0)) FOR [x_secureSource]
END


END
GO

/****** Object:  Default [DF__authorize__x_eCh__5E54FF49]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__authorize__x_eCh__5E54FF49]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__authorize__x_eCh__5E54FF49]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorizeNet] ADD  CONSTRAINT [DF__authorize__x_eCh__5E54FF49]  DEFAULT ((0)) FOR [x_eCheckPending]
END


END
GO
/****** Object:  Default [DF__authorder__idOrd__6225902D]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__authorder__idOrd__6225902D]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__authorder__idOrd__6225902D]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorders] ADD  CONSTRAINT [DF__authorder__idOrd__6225902D]  DEFAULT ((0)) FOR [idOrder]
END


END
GO
/****** Object:  Default [DF__authorder__amoun__6319B466]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__authorder__amoun__6319B466]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__authorder__amoun__6319B466]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorders] ADD  CONSTRAINT [DF__authorder__amoun__6319B466]  DEFAULT ((0)) FOR [amount]
END


END
GO
/****** Object:  Default [DBX_idCustomer_18379]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_idCustomer_18379]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_idCustomer_18379]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorders] ADD  CONSTRAINT [DBX_idCustomer_18379]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__authorder__captu__640DD89F]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__authorder__captu__640DD89F]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__authorder__captu__640DD89F]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorders] ADD  CONSTRAINT [DF__authorder__captu__640DD89F]  DEFAULT ((0)) FOR [captured]
END


END
GO
/****** Object:  Default [DF_authorders_pcSecurityKeyID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_authorders_pcSecurityKeyID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_authorders_pcSecurityKeyID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[authorders] ADD  CONSTRAINT [DF_authorders_pcSecurityKeyID]  DEFAULT ((0)) FOR [pcSecurityKeyID]
END


END
GO
/****** Object:  Default [DF__affiliate__commi__68D28DBC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__affiliate__commi__68D28DBC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__affiliate__commi__68D28DBC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[affiliates] ADD  CONSTRAINT [DF__affiliate__commi__68D28DBC]  DEFAULT ((0)) FOR [commission]
END


END
GO
/****** Object:  Default [DF__affiliate__pcaff__69C6B1F5]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__affiliate__pcaff__69C6B1F5]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__affiliate__pcaff__69C6B1F5]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[affiliates] ADD  CONSTRAINT [DF__affiliate__pcaff__69C6B1F5]  DEFAULT ((0)) FOR [pcaff_Active]
END


END
GO
/****** Object:  Default [DF__admins__idadmin__114A936A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__admins__idadmin__114A936A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__admins__idadmin__114A936A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[admins] ADD  CONSTRAINT [DF__admins__idadmin__114A936A]  DEFAULT ((0)) FOR [idadmin]
END


END
GO
/****** Object:  Default [DF_admins_pcSecurityKeyID]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_admins_pcSecurityKeyID]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_admins_pcSecurityKeyID]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[admins] ADD  CONSTRAINT [DF_admins_pcSecurityKeyID]  DEFAULT ((0)) FOR [pcSecurityKeyID]
END


END
GO
/****** Object:  Default [DBX_textarea_16175]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_textarea_16175]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_textarea_16175]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[xfields] ADD  CONSTRAINT [DBX_textarea_16175]  DEFAULT ((0)) FOR [textarea]
END


END
GO
/****** Object:  Default [DBX_widthoffield_30648]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_widthoffield_30648]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_widthoffield_30648]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[xfields] ADD  CONSTRAINT [DBX_widthoffield_30648]  DEFAULT ((0)) FOR [widthoffield]
END


END
GO
/****** Object:  Default [DBX_maxlength_19072]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_maxlength_19072]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_maxlength_19072]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[xfields] ADD  CONSTRAINT [DBX_maxlength_19072]  DEFAULT ((0)) FOR [maxlength]
END


END
GO
/****** Object:  Default [DBX_rowlength_13679]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_rowlength_13679]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_rowlength_13679]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[xfields] ADD  CONSTRAINT [DBX_rowlength_13679]  DEFAULT ((0)) FOR [rowlength]
END


END
GO
/****** Object:  Default [DBX_randnum_29557]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DBX_randnum_29557]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DBX_randnum_29557]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[xfields] ADD  CONSTRAINT [DBX_randnum_29557]  DEFAULT ((0)) FOR [randnum]
END


END
GO
/****** Object:  Default [DF__wishList__idCust__6462DE5A]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__wishList__idCust__6462DE5A]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__wishList__idCust__6462DE5A]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[wishList] ADD  CONSTRAINT [DF__wishList__idCust__6462DE5A]  DEFAULT ((0)) FOR [idCustomer]
END


END
GO
/****** Object:  Default [DF__wishList__idProd__65570293]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__wishList__idProd__65570293]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__wishList__idProd__65570293]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[wishList] ADD  CONSTRAINT [DF__wishList__idProd__65570293]  DEFAULT ((0)) FOR [idProduct]
END


END
GO
/****** Object:  Default [DF__wishList__idconf__664B26CC]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__wishList__idconf__664B26CC]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__wishList__idconf__664B26CC]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[wishList] ADD  CONSTRAINT [DF__wishList__idconf__664B26CC]  DEFAULT ((0)) FOR [idconfigWishlistSession]
END


END
GO
/****** Object:  Default [DF__wishList__QSubmi__673F4B05]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__wishList__QSubmi__673F4B05]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__wishList__QSubmi__673F4B05]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[wishList] ADD  CONSTRAINT [DF__wishList__QSubmi__673F4B05]  DEFAULT ((0)) FOR [QSubmit]
END


END
GO
/****** Object:  Default [DF_idOptionA_wishList]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_idOptionA_wishList]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_idOptionA_wishList]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[wishList] ADD  CONSTRAINT [DF_idOptionA_wishList]  DEFAULT ((0)) FOR [idOptionA]
END


END
GO
/****** Object:  Default [DF_wishList_idOptionB]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_wishList_idOptionB]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF_wishList_idOptionB]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[wishList] ADD  CONSTRAINT [DF_wishList_idOptionB]  DEFAULT ((0)) FOR [idOptionB]
END


END
GO


IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[verisign_pfp]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
INSERT INTO affiliates (affiliateName, commission) VALUES ('None',0)

INSERT INTO authorizeNet (id, x_Type, x_Login, x_Password, x_version,x_Curcode,x_Method,x_AIMType,x_CVV,x_testmode) VALUES (1,'AUTH_ONLY', 'testdriver', 'testdriver', '3.1','USD','AIM','PASS',0,1)

INSERT INTO categories (categoryDesc, idParentCategory, tier, serviceSpec, required, definePrd, priority, multi,details,image,basePrice) VALUES ('< No Parent - Top Level Category >',1,0,0,0,0,0,0,'','',0)

INSERT INTO CCTypes (CCType, active, CCcode) VALUES ('MasterCard',0,'M')
INSERT INTO CCTypes (CCType, active, CCcode) VALUES ('Visa',0,'V')
INSERT INTO CCTypes (CCType, active, CCcode) VALUES ('American Express',0,'A')
INSERT INTO CCTypes (CCType, active, CCcode) VALUES ('Discover',0,'D')
INSERT INTO CCTypes (CCType, active, CCcode) VALUES ('Diners Card',0,'DC')

INSERT INTO countries (countryName, countryCode) VALUES ('Afghanistan','AF')
INSERT INTO countries (countryName, countryCode) VALUES ('Åland Islands','AX')
INSERT INTO countries (countryName, countryCode) VALUES ('Albania','AL')
INSERT INTO countries (countryName, countryCode) VALUES ('Algeria','DZ')
INSERT INTO countries (countryName, countryCode) VALUES ('American Samoa','AS')
INSERT INTO countries (countryName, countryCode) VALUES ('Andorra','AD')
INSERT INTO countries (countryName, countryCode) VALUES ('Angola','AO')
INSERT INTO countries (countryName, countryCode) VALUES ('Anguilla','AI')
INSERT INTO countries (countryName, countryCode) VALUES ('Antarctica','AQ')
INSERT INTO countries (countryName, countryCode) VALUES ('Antigua and Barbuda','AG')
INSERT INTO countries (countryName, countryCode) VALUES ('Argentina', 'AR')
INSERT INTO countries (countryName, countryCode) VALUES ('Armenia','AM')
INSERT INTO countries (countryName, countryCode) VALUES ('Aruba', 'AW')
INSERT INTO countries (countryName, countryCode, pcSubDivisionID) VALUES ('Australia', 'AU', 1)
INSERT INTO countries (countryName, countryCode) VALUES ('Austria', 'AT')
INSERT INTO countries (countryName, countryCode) VALUES ('Azerbaijan','AZ')
INSERT INTO countries (countryName, countryCode) VALUES ('Bahamas', 'BS')
INSERT INTO countries (countryName, countryCode) VALUES ('Bahrain', 'BH')
INSERT INTO countries (countryName, countryCode) VALUES ('Bangladesh', 'BD')
INSERT INTO countries (countryName, countryCode) VALUES ('Barbados', 'BB')
INSERT INTO countries (countryName, countryCode) VALUES ('Belarus', 'BY')
INSERT INTO countries (countryName, countryCode) VALUES ('Belgium', 'BE')
INSERT INTO countries (countryName, countryCode) VALUES ('Belize', 'BZ')
INSERT INTO countries (countryName, countryCode) VALUES ('Benin', 'BJ')
INSERT INTO countries (countryName, countryCode) VALUES ('Bermuda', 'BM')
INSERT INTO countries (countryName, countryCode) VALUES ('Bhutan','BT')
INSERT INTO countries (countryName, countryCode) VALUES ('Bolivia', 'BO')
INSERT INTO countries (countryName, countryCode) VALUES ('Bosnia and Herzegovina','BA')
INSERT INTO countries (countryName, countryCode) VALUES ('Botswana', 'BW')
INSERT INTO countries (countryName, countryCode) VALUES ('Bouvet Island','BV')
INSERT INTO countries (countryName, countryCode) VALUES ('Brazil', 'BR')
INSERT INTO countries (countryName, countryCode) VALUES ('British Indian Ocean Territory','IO')
INSERT INTO countries (countryName, countryCode) VALUES ('Brunei Darussalam','BN')
INSERT INTO countries (countryName, countryCode) VALUES ('Bulgaria', 'BG')
INSERT INTO countries (countryName, countryCode) VALUES ('Burkina Faso', 'BF')
INSERT INTO countries (countryName, countryCode) VALUES ('Burundi', 'BI')
INSERT INTO countries (countryName, countryCode) VALUES ('Cambodia', 'KH')
INSERT INTO countries (countryName, countryCode) VALUES ('Cameroon', 'CM')
INSERT INTO countries (countryName, countryCode, pcSubDivisionID) VALUES ('Canada', 'CA', 1)
INSERT INTO countries (countryName, countryCode) VALUES ('Cape Verde', 'CV')
INSERT INTO countries (countryName, countryCode) VALUES ('Cayman Islands', 'KY')
INSERT INTO countries (countryName, countryCode) VALUES ('Central African Republic', 'CF')
INSERT INTO countries (countryName, countryCode) VALUES ('Chad', 'TD')
INSERT INTO countries (countryName, countryCode) VALUES ('Chile', 'CL')
INSERT INTO countries (countryName, countryCode) VALUES ('China - Peoples Republic of', 'CN')
INSERT INTO countries (countryName, countryCode) VALUES ('Christmas Island', 'CX')
INSERT INTO countries (countryName, countryCode) VALUES ('Cocos (Keeling) Islands', 'CC')
INSERT INTO countries (countryName, countryCode) VALUES ('Colombia', 'CO')
INSERT INTO countries (countryName, countryCode) VALUES ('Comoros','KM')
INSERT INTO countries (countryName, countryCode) VALUES ('Congo', 'CG')
INSERT INTO countries (countryName, countryCode) VALUES ('Congo - The Democratic Republic of the','CD')
INSERT INTO countries (countryName, countryCode) VALUES ('Cook Islands', 'CK')
INSERT INTO countries (countryName, countryCode) VALUES ('Costa Rica', 'CR')
INSERT INTO countries (countryName, countryCode) VALUES ('C?te D''Ivoire','CI')
INSERT INTO countries (countryName, countryCode) VALUES ('Croatia', 'HR')
INSERT INTO countries (countryName, countryCode) VALUES ('Cuba','CU')
INSERT INTO countries (countryName, countryCode) VALUES ('Cyprus', 'CY')
INSERT INTO countries (countryName, countryCode) VALUES ('Czech Republic', 'CZ')
INSERT INTO countries (countryName, countryCode) VALUES ('Denmark', 'DK')
INSERT INTO countries (countryName, countryCode) VALUES ('Djibouti', 'DJ')
INSERT INTO countries (countryName, countryCode) VALUES ('Dominica', 'DM')
INSERT INTO countries (countryName, countryCode) VALUES ('Dominican Republic', 'DO')
INSERT INTO countries (countryName, countryCode) VALUES ('Ecuador', 'EC')
INSERT INTO countries (countryName, countryCode) VALUES ('Egypt', 'EG')
INSERT INTO countries (countryName, countryCode) VALUES ('El Salvador', 'SV')
INSERT INTO countries (countryName, countryCode) VALUES ('Equatorial Guinea', 'GQ')
INSERT INTO countries (countryName, countryCode) VALUES ('Eritrea', 'ER')
INSERT INTO countries (countryName, countryCode) VALUES ('Estonia', 'EE')
INSERT INTO countries (countryName, countryCode) VALUES ('Ethiopia', 'ET')
INSERT INTO countries (countryName, countryCode) VALUES ('Falkland Islands (Malvinas)','FK')
INSERT INTO countries (countryName, countryCode) VALUES ('Faroe Islands', 'FO')
INSERT INTO countries (countryName, countryCode) VALUES ('Fiji', 'FJ')
INSERT INTO countries (countryName, countryCode) VALUES ('Finland', 'FI')
INSERT INTO countries (countryName, countryCode) VALUES ('France', 'FR')
INSERT INTO countries (countryName, countryCode) VALUES ('French Guiana', 'GF')
INSERT INTO countries (countryName, countryCode) VALUES ('French Polynesia', 'PF')
INSERT INTO countries (countryName, countryCode) VALUES ('French Southern Territories','TF')
INSERT INTO countries (countryName, countryCode) VALUES ('Gabon', 'GA')
INSERT INTO countries (countryName, countryCode) VALUES ('Gambia', 'GM')
INSERT INTO countries (countryName, countryCode) VALUES ('Georgia','GE')
INSERT INTO countries (countryName, countryCode) VALUES ('Germany', 'DE')
INSERT INTO countries (countryName, countryCode) VALUES ('Ghana', 'GH')
INSERT INTO countries (countryName, countryCode) VALUES ('Gibraltar', 'GI')
INSERT INTO countries (countryName, countryCode) VALUES ('Greece', 'GR')
INSERT INTO countries (countryName, countryCode) VALUES ('Greenland', 'GL')
INSERT INTO countries (countryName, countryCode) VALUES ('Grenada', 'GD')
INSERT INTO countries (countryName, countryCode) VALUES ('Guadeloupe', 'GP')
INSERT INTO countries (countryName, countryCode) VALUES ('Guam', 'GU')
INSERT INTO countries (countryName, countryCode) VALUES ('Guatemala', 'GT')
INSERT INTO countries (countryName, countryCode) VALUES ('Guernsey','GG')
INSERT INTO countries (countryName, countryCode) VALUES ('Guinea', 'GN')
INSERT INTO countries (countryName, countryCode) VALUES ('Guinea-Bissau', 'GW')
INSERT INTO countries (countryName, countryCode) VALUES ('Guyana', 'GY')
INSERT INTO countries (countryName, countryCode) VALUES ('Haiti', 'HT')
INSERT INTO countries (countryName, countryCode) VALUES ('Heard Island and McDonald Islands','HM')
INSERT INTO countries (countryName, countryCode) VALUES ('Honduras', 'HN')
INSERT INTO countries (countryName, countryCode) VALUES ('Hong Kong', 'HK')
INSERT INTO countries (countryName, countryCode) VALUES ('Hungary', 'HU')
INSERT INTO countries (countryName, countryCode) VALUES ('Iceland', 'IS')
INSERT INTO countries (countryName, countryCode) VALUES ('India', 'IN')
INSERT INTO countries (countryName, countryCode) VALUES ('Indonesia', 'ID')
INSERT INTO countries (countryName, countryCode) VALUES ('Iran - Islamic Republic Of', 'IR')
INSERT INTO countries (countryName, countryCode) VALUES ('Iraq','IQ')
INSERT INTO countries (countryName, countryCode) VALUES ('Ireland', 'IE')
INSERT INTO countries (countryName, countryCode) VALUES ('Isle of Man','IM')
INSERT INTO countries (countryName, countryCode) VALUES ('Israel', 'IL')
INSERT INTO countries (countryName, countryCode) VALUES ('Italy', 'IT')
INSERT INTO countries (countryName, countryCode) VALUES ('Jamaica', 'JM')
INSERT INTO countries (countryName, countryCode) VALUES ('Japan', 'JP')
INSERT INTO countries (countryName, countryCode) VALUES ('Jersey','JE')
INSERT INTO countries (countryName, countryCode) VALUES ('Jordan', 'JO')
INSERT INTO countries (countryName, countryCode) VALUES ('Kazakhstan', 'KZ')
INSERT INTO countries (countryName, countryCode) VALUES ('Kenya', 'KE')
INSERT INTO countries (countryName, countryCode) VALUES ('Kiribati', 'KI')
INSERT INTO countries (countryName, countryCode) VALUES ('Korea - Democratic People''s Republic Of','KP')
INSERT INTO countries (countryName, countryCode) VALUES ('Korea - Republic of','KR')
INSERT INTO countries (countryName, countryCode) VALUES ('Kuwait', 'KW')
INSERT INTO countries (countryName, countryCode) VALUES ('Kyrgyzstan', 'KG')
INSERT INTO countries (countryName, countryCode) VALUES ('Lao People''s Democratic Republic', 'LA')
INSERT INTO countries (countryName, countryCode) VALUES ('Latvia', 'LV')
INSERT INTO countries (countryName, countryCode) VALUES ('Lebanon', 'LB')
INSERT INTO countries (countryName, countryCode) VALUES ('Lesotho', 'LS')
INSERT INTO countries (countryName, countryCode) VALUES ('Liberia', 'LR')
INSERT INTO countries (countryName, countryCode) VALUES ('Libyan Arab Jamahiriya','LY')
INSERT INTO countries (countryName, countryCode) VALUES ('Liechtenstein', 'LI')
INSERT INTO countries (countryName, countryCode) VALUES ('Lithuania', 'LT')
INSERT INTO countries (countryName, countryCode) VALUES ('Luxembourg', 'LU')
INSERT INTO countries (countryName, countryCode) VALUES ('Macao', 'MO')
INSERT INTO countries (countryName, countryCode) VALUES ('Macedonia', 'MK')
INSERT INTO countries (countryName, countryCode) VALUES ('Madagascar', 'MG')
INSERT INTO countries (countryName, countryCode) VALUES ('Malawi', 'MW')
INSERT INTO countries (countryName, countryCode) VALUES ('Malaysia', 'MY')
INSERT INTO countries (countryName, countryCode) VALUES ('Maldives', 'MV')
INSERT INTO countries (countryName, countryCode) VALUES ('Mali', 'ML')
INSERT INTO countries (countryName, countryCode) VALUES ('Malta', 'MT')
INSERT INTO countries (countryName, countryCode) VALUES ('Marshall Islands', 'MH')
INSERT INTO countries (countryName, countryCode) VALUES ('Martinique', 'MQ')
INSERT INTO countries (countryName, countryCode) VALUES ('Mauritania', 'MR')
INSERT INTO countries (countryName, countryCode) VALUES ('Mauritius', 'MU')
INSERT INTO countries (countryName, countryCode) VALUES ('Mayotte','YT')
INSERT INTO countries (countryName, countryCode) VALUES ('Mexico', 'MX')
INSERT INTO countries (countryName, countryCode) VALUES ('Micronesia - Federated States of','FM')
INSERT INTO countries (countryName, countryCode) VALUES ('Moldova - Republic of','MD')
INSERT INTO countries (countryName, countryCode) VALUES ('Monaco','MC')
INSERT INTO countries (countryName, countryCode) VALUES ('Mongolia','MN')
INSERT INTO countries (countryName, countryCode) VALUES ('Montenegro','ME')
INSERT INTO countries (countryName, countryCode) VALUES ('Montserrat', 'MS')
INSERT INTO countries (countryName, countryCode) VALUES ('Morocco', 'MA')
INSERT INTO countries (countryName, countryCode) VALUES ('Mozambique', 'MZ')
INSERT INTO countries (countryName, countryCode) VALUES ('Myanmar', 'MM')
INSERT INTO countries (countryName, countryCode) VALUES ('Namibia', 'NA')
INSERT INTO countries (countryName, countryCode) VALUES ('Nauru','NR')
INSERT INTO countries (countryName, countryCode) VALUES ('Nepal', 'NP')
INSERT INTO countries (countryName, countryCode) VALUES ('Netherlands', 'NL')
INSERT INTO countries (countryName, countryCode) VALUES ('Netherlands Antilles', 'AN')
INSERT INTO countries (countryName, countryCode) VALUES ('New Caledonia','NC')
INSERT INTO countries (countryName, countryCode, pcSubDivisionID) VALUES ('New Zealand', 'NZ', 1)
INSERT INTO countries (countryName, countryCode) VALUES ('Nicaragua', 'NI')
INSERT INTO countries (countryName, countryCode) VALUES ('Niger', 'NE')
INSERT INTO countries (countryName, countryCode) VALUES ('Nigeria', 'NG')
INSERT INTO countries (countryName, countryCode) VALUES ('Niue', 'NU')
INSERT INTO countries (countryName, countryCode) VALUES ('Norfolk Island', 'NF')
INSERT INTO countries (countryName, countryCode) VALUES ('Nothern Mariana Islands','MP')
INSERT INTO countries (countryName, countryCode) VALUES ('Norway', 'NO')
INSERT INTO countries (countryName, countryCode) VALUES ('Oman', 'OM')
INSERT INTO countries (countryName, countryCode) VALUES ('Pakistan', 'PK')
INSERT INTO countries (countryName, countryCode) VALUES ('Palau', 'PW')
INSERT INTO countries (countryName, countryCode) VALUES ('Palestinian Territory - Occupied','PS')
INSERT INTO countries (countryName, countryCode) VALUES ('Panama', 'PA')
INSERT INTO countries (countryName, countryCode) VALUES ('Papua New Guinea', 'PG')
INSERT INTO countries (countryName, countryCode) VALUES ('Paraguay', 'PY')
INSERT INTO countries (countryName, countryCode) VALUES ('Peru', 'PE')
INSERT INTO countries (countryName, countryCode) VALUES ('Philippines', 'PH')
INSERT INTO countries (countryName, countryCode) VALUES ('Pitcairn','PN')
INSERT INTO countries (countryName, countryCode) VALUES ('Poland','PL')
INSERT INTO countries (countryName, countryCode) VALUES ('Portugal', 'PT')
INSERT INTO countries (countryName, countryCode) VALUES ('Puerto Rico', 'PR')
INSERT INTO countries (countryName, countryCode) VALUES ('Qatar', 'QA')
INSERT INTO countries (countryName, countryCode) VALUES ('Réunion', 'RE')
INSERT INTO countries (countryName, countryCode) VALUES ('Romania', 'RO')
INSERT INTO countries (countryName, countryCode) VALUES ('Russian Federation', 'RU')
INSERT INTO countries (countryName, countryCode) VALUES ('Rwanda', 'RW')
INSERT INTO countries (countryName, countryCode) VALUES ('Saint Helena','SH')
INSERT INTO countries (countryName, countryCode) VALUES ('Saint Kitts and Nevis','KN')
INSERT INTO countries (countryName, countryCode) VALUES ('Saint Lucia','LC')
INSERT INTO countries (countryName, countryCode) VALUES ('Saint Pierre and Miquelon','PM')
INSERT INTO countries (countryName, countryCode) VALUES ('Saint Vincent and the Grenadines','VC')
INSERT INTO countries (countryName, countryCode) VALUES ('Samoa','WS')
INSERT INTO countries (countryName, countryCode) VALUES ('San Marino','SM')
INSERT INTO countries (countryName, countryCode) VALUES ('Sao Tome and Principe','ST')
INSERT INTO countries (countryName, countryCode) VALUES ('Saudi Arabia','SA')
INSERT INTO countries (countryName, countryCode) VALUES ('Senegal','SN')
INSERT INTO countries (countryName, countryCode) VALUES ('Serbia','RS')
INSERT INTO countries (countryName, countryCode) VALUES ('Seychelles','SC')
INSERT INTO countries (countryName, countryCode) VALUES ('Sierra Leone','SL')
INSERT INTO countries (countryName, countryCode) VALUES ('Singapore','SG')
INSERT INTO countries (countryName, countryCode) VALUES ('Slovakia','SK')
INSERT INTO countries (countryName, countryCode) VALUES ('Slovenia','SI')
INSERT INTO countries (countryName, countryCode) VALUES ('Solomon Islands','SB')
INSERT INTO countries (countryName, countryCode) VALUES ('Somalia','SO')
INSERT INTO countries (countryName, countryCode) VALUES ('South Africa','ZA')
INSERT INTO countries (countryName, countryCode) VALUES ('South Georgia and the South Sandwich Islands','GS')
INSERT INTO countries (countryName, countryCode) VALUES ('Spain', 'ES')
INSERT INTO countries (countryName, countryCode) VALUES ('Sri Lanka', 'LK')
INSERT INTO countries (countryName, countryCode) VALUES ('Sudan', 'SD')
INSERT INTO countries (countryName, countryCode) VALUES ('Suriname', 'SR')
INSERT INTO countries (countryName, countryCode) VALUES ('Svalbard and Jan Mayen','SJ')
INSERT INTO countries (countryName, countryCode) VALUES ('Swaziland', 'SZ')
INSERT INTO countries (countryName, countryCode) VALUES ('Sweden', 'SE')
INSERT INTO countries (countryName, countryCode) VALUES ('Switzerland', 'CH')
INSERT INTO countries (countryName, countryCode) VALUES ('Syrian Arab Republic', 'SY')
INSERT INTO countries (countryName, countryCode) VALUES ('Taiwan - Province Of China', 'TW')
INSERT INTO countries (countryName, countryCode) VALUES ('Tajikistan', 'TJ')
INSERT INTO countries (countryName, countryCode) VALUES ('Tanzania - United Republic Of', 'TZ')
INSERT INTO countries (countryName, countryCode) VALUES ('Thailand', 'TH')
INSERT INTO countries (countryName, countryCode) VALUES ('Timor-Leste','TL')
INSERT INTO countries (countryName, countryCode) VALUES ('Togo', 'TG')
INSERT INTO countries (countryName, countryCode) VALUES ('Tokelau','TK')
INSERT INTO countries (countryName, countryCode) VALUES ('Tonga','TO')
INSERT INTO countries (countryName, countryCode) VALUES ('Trinidad And Tobago', 'TT')
INSERT INTO countries (countryName, countryCode) VALUES ('Tunisia', 'TN')
INSERT INTO countries (countryName, countryCode) VALUES ('Turkey', 'TR')
INSERT INTO countries (countryName, countryCode) VALUES ('Turkmenistan','TM')
INSERT INTO countries (countryName, countryCode) VALUES ('Turks and Caicos Islands', 'TC')
INSERT INTO countries (countryName, countryCode) VALUES ('Tuvalu', 'TV')
INSERT INTO countries (countryName, countryCode) VALUES ('Uganda', 'UG')
INSERT INTO countries (countryName, countryCode) VALUES ('Ukraine', 'UA')
INSERT INTO countries (countryName, countryCode) VALUES ('United Arab Emirates', 'AE')
INSERT INTO countries (countryName, countryCode) VALUES ('United Kingdom', 'GB')
INSERT INTO countries (countryName, countryCode, pcSubDivisionID) VALUES ('United States', 'US', 1)
INSERT INTO countries (countryName, countryCode) VALUES ('United States Minor Outlying Islands','UM')
INSERT INTO countries (countryName, countryCode) VALUES ('Uruguay', 'UY')
INSERT INTO countries (countryName, countryCode) VALUES ('Uzbekistan', 'UZ')
INSERT INTO countries (countryName, countryCode) VALUES ('Vanuatu', 'VU')
INSERT INTO countries (countryName, countryCode) VALUES ('Vatican City','VA')
INSERT INTO countries (countryName, countryCode) VALUES ('Venezuela', 'VE')
INSERT INTO countries (countryName, countryCode) VALUES ('VietNam', 'VN')
INSERT INTO countries (countryName, countryCode) VALUES ('Virgin Islands - British','VG')
INSERT INTO countries (countryName, countryCode) VALUES ('Virgin Islands - U.S.','VI')
INSERT INTO countries (countryName, countryCode) VALUES ('Wallis and Futuna Islands', 'WF')
INSERT INTO countries (countryName, countryCode) VALUES ('Western Sahara','EH')
INSERT INTO countries (countryName, countryCode) VALUES ('Yemen', 'YE')
INSERT INTO countries (countryName, countryCode) VALUES ('Zambia', 'ZM')
INSERT INTO countries (countryName, countryCode) VALUES ('Zimbabwe', 'ZW')

INSERT INTO emailSettings (id,ownerEmail, frmEmail, ConfirmEmail,PayPalEmail,ReceivedEmail,ShippedEmail,CancelledEmail) VALUES (1,'email@yourdomain.com','email@yourdomain.com','Dear <CUSTOMER_NAME><br><br>We wanted to let you know that order number <ORDER_ID> that you placed on <TODAY_DATE> has been processed and will be shipped soon.<br><br>This is your order confirmation. Order details are listed below.<br><br>If you have any questions, please do not hesitate to contact us.','We have received your order and we are awaiting payment confirmation from PayPal, the payment option that you selected. As soon as payment confirmation is received, your order will be processed and you will receive an order receipt at this email address.','Dear <CUSTOMER_NAME><br><br>Thank you for shopping at <COMPANY>.<br><br>We received your order on <TODAY_DATE>. Your order number is <ORDER_ID>.<br><br>Note that this is not an order confirmation. You will receive a detailed confirmation message once your order has been processed. You can check the status of your order by logging into your account at <COMPANY_URL>/productcart/pc/custpref.asp<br><br>If you have any questions, please do not hesitate to contact us.<br><br>Thank you for being a <COMPANY> customer.<br><br>Best Regards,<br><COMPANY>','Dear <CUSTOMER_NAME><br><br>We thought you may like to know that your order number <ORDER_ID> has been shipped. Shipping details are listed below.<br><br>If you have any questions, please do not hesitate to contact us.','This message is to inform you that order number <ORDER_ID> that you submitted in this store on <ORDER_DATE> has been cancelled.')

INSERT INTO ITransact (Gateway_ID, URL, id,it_amex,it_diner,it_disc,it_mc,it_visa) VALUES ('00000','',1,0,0,0,0,1)

INSERT INTO layout (ID, headerid, recalculate, continueshop, checkout, submit, morebtn, viewcartbtn, checkoutbtn, addtocart, addtowl, register, cancel, [remove], add2, login, login_checkout, back, register_checkout, customize, [reconfigure], resetdefault, savequote, RevOrder, SubmitQuote, pcLO_requestQuote, pcLO_placeOrder, pcLO_checkoutWR, pcLO_processShip, pcLO_finalShip, pcLO_backtoOrder, pcLO_previous, pcLO_next, CreRegistry, DelRegistry, AddToRegistry, UpdRegistry, SendMsgs, RetRegistry, pcLO_Update, pcLO_Savecart) VALUES (2, '1', 'images/sample/pc_button_recalculate.gif', 'images/sample/pc_button_continue_shop.gif', 'images/sample/pc_button_checkout.gif', 'images/sample/pc_button_continue.gif', 'images/sample/pc_button_details.gif', 'images/sample/pc_button_viewcart.gif', 'images/sample/pc_button_tellafriend.gif', 'images/sample/pc_button_add.gif', 'images/sample/pc_button_wishlist.gif', 'images/sample/pc_button_register.gif', 'images/sample/pc_button_cancel.gif', 'images/sample/pc_button_remove.gif', 'images/sample/pc_button_addSmall.gif', 'images/sample/pc_button_login.gif', 'images/sample/pc_button_login_checkout.gif', 'images/sample/pc_button_back.gif', 'images/sample/pc_button_register_checkout.gif', 'images/sample/pc_button_customize.gif', 'images/sample/pc_button_reconfig.gif', 'images/sample/pc_button_reset.gif', 'images/sample/pc_button_quote.gif', 'images/sample/pc_button_review_order.gif', 'images/sample/pc_button_submit_quote.gif', 'images/sample/pc_button_request_quote.gif', 'images/sample/pc_button_placeOrder.gif', 'images/sample/pc_button_checkoutwr.gif', 'images/sample/pc_button_shipment_process.gif', 'images/sample/pc_button_shipment_finalize.gif', 'images/sample/pc_button_back_to_order_details.gif', 'images/sample/pc_button_previous.gif', 'images/sample/pc_button_next.gif', 'images/sample/ggg_button_create.gif', 'images/sample/ggg_button_delete.gif', 'images/sample/ggg_button_add.gif', 'images/sample/ggg_button_update.gif', 'images/sample/ggg_button_send.gif', 'images/sample/ggg_button_return.gif', 'images/sample/pc_button_update.gif', 'images/sample/pc_save_cart.gif')

INSERT INTO linkpoint (id, storeName,transType,lp_testmode,lp_cards,CVM,lp_yourpay) VALUES (1, '000000','sale','YES','V',0,'NO')

INSERT INTO pcPay_PayPalAdvanced (pcPay_PayPalAd_ID) VALUES (1)

INSERT INTO optionsGroups (OptionGroupDesc) VALUES ('No Assignments')

INSERT INTO paypal (id, Pay_To, URL,PP_Currency) VALUES (1, 'sales@yourdomain.com','https://www.paypal.com/cgi-bin/webscr','USD')

INSERT INTO payTypes (gwCode, paymentDesc, priceToAdd, percentageToAdd, ssl, sslUrl,quantityFrom, quantityUntil, weightFrom, weightuntil, priceFrom, priceuntil, active, Cbtob, CReq, Type, paymentPriority,pcPayTypes_processOrder,pcPayTypes_setPayStatus) VALUES (6,'Credit Card',0, 0, -1,'paymnta_o.asp',0,9999,0,9999,0,9999,0,0,0,'A',0,0,0)

INSERT INTO PSIGate (Config_File_Name,Config_File_Name_Full,Host,Port,Userid,Mode,id,psi_post,psi_testmode) VALUES ('teststore','teststore.pem','secure.psigate.com','1139','TestPsi','1',1,'HTML','YES')


INSERT INTO protx (idProtx, Protxid, ProtxPassword, CVV, ProtxTestmode, ProtxCurcode, TxType, avs, ProtxMethod, ProtxCardTypes, ProtxApply3DSecure) VALUES (1,'testvendor','', 1,1,'GBP','PAYMENT',1,'FORM','VISA',3)

INSERT INTO pcPay_CBN (pcPay_CBN_id, pcPay_CBN_merchant, pcPay_CBN_test, pcPay_CBN_status) VALUES (1,'',0,0)

INSERT INTO pcPay_CyberSource (pcPay_Cys_Id, pcPay_Cys_MerchantID, pcPay_Cys_TransType, pcPay_Cys_CardType, pcPay_Cys_CVV, pcPay_Cys_TestMode) VALUES (1,'',0,'V,M,A,D,E',0,0)

INSERT INTO pcPay_USAePay (pcPay_Uep_Id,pcPay_Uep_SourceKey,pcPay_Uep_TransType,pcPay_Uep_TestMode,pcPay_Uep_Checking,pcPay_Uep_CheckPending) VALUES (1,'',0,1,0,0)

INSERT INTO ShipmentTypes (idShipment, shipmentDesc, priceToAdd, active, international, ipriceToAdd, shipserver) VALUES (1,'FedEx',0,0,0,0,'GR1')
INSERT INTO ShipmentTypes (idShipment, shipmentDesc, priceToAdd, active, international, ipriceToAdd, shipserver) VALUES (3,'UPS',0,0,0,0,'https://www.ups.com/ups.app/xml/Rate')
INSERT INTO ShipmentTypes (idShipment, shipmentDesc, priceToAdd, active, international, ipriceToAdd, shipserver) VALUES (4,'USPS',0,0,0,0,'http://production.shippingapis.com/ShippingAPI.dll')
INSERT INTO ShipmentTypes (idShipment, shipmentDesc, priceToAdd, active, international, ipriceToAdd, shipserver) VALUES (7,'Canada Post',0,0,0,0,'http://206.191.4.228:30000')
INSERT INTO ShipmentTypes (idShipment, shipmentDesc, priceToAdd, active, international, ipriceToAdd, shipserver) VALUES (8,'Custom',0,0,0,0,'')
INSERT INTO ShipmentTypes (idShipment, shipmentDesc, priceToAdd, active, international, ipriceToAdd, shipserver, AccessLicense) VALUES (9, 'FedExWS', 0, 0, 0, 0, '', 'LIVE')

INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('AL', 'Alabama', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('AK', 'Alaska', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('AR', 'Arkansas', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('AZ', 'Arizona', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('BVI', 'British Virgin Islands', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('CA', 'California', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('CO', 'Colorado', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('CT', 'Connecticut', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('DE', 'Delaware', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('FL', 'Florida', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('GA', 'Georgia', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('GU', 'Guam', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('HI', 'Hawaii', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('IA', 'Iowa', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('ID', 'Idaho', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('IL', 'Illinois', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('IN', 'Indiana', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('KS', 'Kansas', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('KY', 'Kentucky', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('LA', 'Louisiana', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MA', 'Massachusetts', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MD', 'Maryland', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('ME', 'Maine', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MI', 'Michigan', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MN', 'Minnesota', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MO', 'Missouri', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MP', 'Northern Mariana Islands', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MPI', 'Mariana Islands (Pacific)', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MS', 'Mississippi', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MT', 'Montana', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NC', 'North Carolina', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('ND', 'North Dakota', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NE', 'Nebraska', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NH', 'New Hampshire', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NJ', 'New Jersey', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NM', 'New Mexico', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NV', 'Nevada', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NY', 'New York', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('OH', 'Ohio', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('OK', 'Oklahoma', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('OR', 'Oregon', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('PA', 'Pennsylvania', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('PR', 'Puerto Rico', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('RI', 'Rhode Island', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('SC', 'South Carolina', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('SD', 'South Dakota', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('TN', 'Tennessee', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('TX', 'Texas', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('VI', 'U.S. Virgin Islands', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('UT', 'Utah', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('VA', 'Virginia', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('VT', 'Vermont', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('WA', 'Washington', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('WI', 'Wisconsin', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('WV', 'West Virginia', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('WY', 'Wyoming', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('ACT', 'Australian Capital Territory', 'AU')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NSW', 'New South Wales', 'AU')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NT', 'Northern Territory', 'AU')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('QLD', 'Queensland', 'AU')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('SA', 'South Australia', 'AU')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('TAS', 'Tasmania', 'AU')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('VIC', 'Victoria', 'AU')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('WA', 'Western Australia', 'AU')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('AUK', 'Auckland N', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('BOP', 'Bay of Plenty N', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('CAN', 'Canterbury S', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('CI', 'Chatham Islands', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('GIS', 'Gisborne N', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('HKB', 'Hawke''s Bay', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MBH', 'Marlborough S', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MWT', 'Manawatu-Wanganui N', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NSN', 'Nelson S', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NTL', 'Northland N', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('OTA', 'Otago S', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('STL', 'Southland S', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('TAS', 'Tasman S', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('TKI', 'Taranaki N', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('WKO', 'Waikato N', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('WGN', 'Wellington N', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('WTC', 'West Coast S', 'NZ')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('AB', 'Alberta', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('BC', 'British Columbia', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('MB', 'Manitoba', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NB', 'New Brunswick', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NL', 'Newfoundland and Labrador', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NT', 'Northwest Territories', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NS', 'Nova Scotia', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('NU', 'Nunavut', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('ON', 'Ontario', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('PE', 'Prince Edward Island', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('QC', 'Quebec', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('SK', 'Saskatchewan', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('YT', 'Yukon Territory', 'CA')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('AS', 'American Samoa', 'US')
INSERT INTO states (stateCode, stateName, pcCountryCode) VALUES ('DC', 'District of Columbia', 'US')


INSERT INTO suppliers (idSupplier, supplierName, receiveSellEmail, receiveUnderStockAlert) VALUES (10, 'None',0,0)

INSERT INTO icons (id, erroricon, requiredicon, errorfieldicon, previousicon, nexticon, discount, zoom) VALUES (1, 'images/sample/pc_icon_error.gif', 'images/sample/pc_icon_required.gif', 'images/sample/pc_icon_errorfield.gif', 'images/sample/pc_icon_prev.gif', 'images/sample/pc_icon_next.gif','images/sample/pc_icon_discount.gif','images/sample/pc_icon_zoom.gif') 

INSERT INTO crossSelldata (id, cs_status, cs_showProd, cs_showCart, cs_showImage, crossSellText, cs_ProductViewCnt, cs_CartViewCnt,cs_ImageHeight,cs_ImageWidth) Values (1,0,0,0,0,'May we also suggest...',9,9,500,300)

INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '01', 0, 'UPS Next Day Air<sup>®</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '02', 0, 'UPS 2<sup>nd</sup> Day Air<sup>®</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '03', 0, 'UPS Ground<sup>®</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '07', 0, 'UPS Worldwide Express<sup>SM</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '08', 0, 'UPS Worldwide Expedited<sup>SM</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '11', 0, 'UPS Standard To Canada<sup>®</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '12', 0, 'UPS 3 Day Select<sup>®</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '13', 0, 'UPS Next Day Air Saver<sup>®</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '14', 0, 'UPS Next Day Air<sup>®</sup> Early A.M.<sup>®</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '54', 0, 'UPS Worldwide Express Plus<sup>SM</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '59', 0, 'UPS 2<sup>nd</sup> Day Air A.M.<sup>®</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '65', 0, 'UPS Express Saver<sup>®</sup>', 0, 0, 0, 0, 0, 0)

INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9901', 0, 'USPS Priority', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9902', 0, 'USPS Express', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9903', 0, 'USPS Parcel', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9904', 0, 'USPS First Class', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9905', 0, 'Global Express Guaranteed<sup>&reg;</sup> Non-Document Rectangular', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9906', 0, 'Express Mail<sup>&reg;</sup> International (EMS)', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9907', 0, 'Priority Mail<sup>&reg;</sup> International', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9908', 0, 'Priority Mail<sup>&reg;</sup> International Flat Rate Envelope', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9909', 0, 'Priority Mail<sup>&reg;</sup> International Flat Rate Box', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9910', 0, 'Global Express Guaranteed<sup>&reg;</sup> Non-Document Non-Rectangular', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9911', 0, 'Express Mail<sup>&reg;</sup> International (EMS) Flat Rate Envelope', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9912', 0, 'First-Class Mail<sup>&reg;</sup> International', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9913', 0, 'USPS Economy (Surface) Standard Post', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9914', 0, 'Global Express Guaranteed<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9915', 0, 'USPS Bound Printed Matter', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9916', 0, 'USPS Media Mail', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '9917', 0, 'USPS Library Mail', 0, 0, 0, 0, 0, 0)

INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '1010', 0, 'Canada Post Regular', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '1020', 0, 'Canada Post Expedited', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '1130', 0, 'Canada Post Xpresspost Evening', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '1030', 0, 'Canada Post Xpresspost', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '1040', 0, 'Canada Post Priority Courier', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '1120', 0, 'Canada Post Expedited Evening', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '1220', 0, 'Canada Post Expedited Saturday', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '1230', 0, 'Canada Post Xpresspost Saturday', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '2010', 0, 'Canada Post Surface US', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '2020', 0, 'Canada Post Air US', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '2030', 0, 'Canada Post Xpresspost US', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '2040', 0, 'Canada Post Puroloator US', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '2050', 0, 'Canada Post Puropak US', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '3010', 0, 'Canada Post Surface International', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '3020', 0, 'Canada Post Air International', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, '3040', 0, 'Canada Post Puroloator International', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '2005', 0, 'Canada Post Small Packets Surface US', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '2015', 0, 'Canada Post Small Packets Air US', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '2025', 0, 'Canada Post Expedited Commercial US', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '3005', 0, 'Canada Post Small Packets Surface International', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '3015', 0, 'Canada Post Small Packets Air International', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '3025', 0, 'Canada Post Xpresspost International', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee,  serviceLimitation) VALUES (0, '3050', 0, 'Canada Post Puropak International', 0, 0, 0, 0, 0, 0) 

INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FIRSTOVERNIGHT', 0, 'FedEx First Overnight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'PRIORITYOVERNIGHT', 0, 'FedEx Priority Overnight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'STANDARDOVERNIGHT', 0, 'FedEx Standard Overnight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX2DAY', 0, 'FedEx 2Day', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEXEXPRESSSAVER', 0, 'FedEx Express Saver', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEXGROUND', 0, 'FedEx Ground', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'GROUNDHOMEDELIVERY', 0, 'FedEx Home Delivery', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONALFIRST', 0, 'FedEx International First', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONALPRIORITY', 0, 'FedEx International Priority', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONALECONOMY', 0, 'FedEx International Economy', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX1DAYFREIGHT', 0, 'FedEx 1Day Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX2DAYFREIGHT', 0, 'FedEx 2Day Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX3DAYFREIGHT', 0, 'FedEx 3Day Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONALPRIORITYFREIGHT', 0, 'FedEx International Priority Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONALECONOMYFREIGHT', 0, 'FedEx International Economy Freight', 0, 0, 0, 0, 0, 0)

INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FIRST_OVERNIGHT', 0, 'FedEx First Overnight<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'PRIORITY_OVERNIGHT', 0, 'FedEx Priority Overnight<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'STANDARD_OVERNIGHT', 0, 'FedEx Standard Overnight<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX_2_DAY', 0, 'FedEx 2Day<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX_EXPRESS_SAVER', 0, 'FedEx Express Saver<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX_GROUND', 0, 'FedEx Ground<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'GROUND_HOME_DELIVERY', 0, 'FedEx Home Delivery<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONAL_FIRST', 0, 'FedEx International First<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONAL_PRIORITY', 0, 'FedEx International Priority<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONAL_ECONOMY', 0, 'FedEx International Economy<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX_1_DAY_FREIGHT', 0, 'FedEx 1Day<sup>&reg;</sup> Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX_2_DAY_FREIGHT', 0, 'FedEx 2Day<sup>&reg;</sup> Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX_3_DAY_FREIGHT', 0, 'FedEx 3Day<sup>&reg;</sup> Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONAL_PRIORITY_FREIGHT', 0, 'FedEx International Priority<sup>&reg;</sup> Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONAL_ECONOMY_FREIGHT', 0, 'FedEx International Economy<sup>&reg;</sup> Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'INTERNATIONAL_GROUND', 0, 'FedEx International Ground<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX_FREIGHT', 0, 'FedEx<sup>&reg;</sup> Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'FEDEX_NATIONAL_FREIGHT', 0, 'FedEx National<sup>&reg;</sup> Freight', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'SMART_POST', 0, 'FedEx SmartPost<sup>&reg;</sup>', 0, 0, 0, 0, 0, 0)
INSERT INTO shipService (serviceActive, serviceCode, servicePriority, serviceDescription, serviceFree, serviceFreeOverAmt, serviceHandlingFee, serviceHandlingIntFee, serviceShowHandlingFee, serviceLimitation) VALUES (0, 'EUROPE_FIRST_INTERNATIONAL_PRIORITY', 0, 'FedEx Europe First - Int''l Priority', 0, 0, 0, 0, 0, 0)

INSERT INTO moneris (moneris_id, store_id) VALUES (1,'store_id')

INSERT INTO Permissions (IDPm, PmName) VALUES (1,'Settings')  
INSERT INTO Permissions (IDPm, PmName) VALUES (2,'Products')  
INSERT INTO Permissions (IDPm, PmName) VALUES (3,'Marketing')  
INSERT INTO Permissions (IDPm, PmName) VALUES (4,'Shipping')  
INSERT INTO Permissions (IDPm, PmName) VALUES (5,'Payments')  
INSERT INTO Permissions (IDPm, PmName) VALUES (6,'Tax Settings')  
INSERT INTO Permissions (IDPm, PmName) VALUES (7,'Customers')  
INSERT INTO Permissions (IDPm, PmName) VALUES (8,'Affiliates')  
INSERT INTO Permissions (IDPm, PmName) VALUES (9,'Orders')  
INSERT INTO Permissions (IDPm, PmName) VALUES (10,'Reports') 
INSERT INTO Permissions (IDPm, PmName) VALUES (11,'Manage Pages (CMS)')


INSERT INTO ups_license (idUPS,ups_UserId, ups_Password, ups_AccessLicense) VALUES (1,'','','')

INSERT INTO twoCheckout (store_id,v2co) VALUES (0,0)

INSERT INTO WorldPay (wp_id,WP_Currency,WP_instID,WP_testmode) VALUES (1,'USD','0','NO')


INSERT INTO eWay (eWayID, eWayCustomerId, eWayPostMethod, eWayTestmode) VALUES (1, '87654321', 'XML', 1)

INSERT INTO fasttransact (AccountID, SiteTag, tran_type, card_types, CVV2) VALUES ('','','SALE','V',0)

INSERT INTO InternetSecure (IsID, IsLanguage, IsCurrency, IsTestmode) VALUES (1, 'EN', 'CDN', 1)

INSERT INTO pcPriority (pcPri_Name,pcPri_Img,pcPri_ShowImg) VALUES ('High','priority_high.gif',1)
INSERT INTO pcPriority (pcPri_Name,pcPri_Img,pcPri_ShowImg) VALUES ('Medium','priority_med.gif',1)
INSERT INTO pcPriority (pcPri_Name,pcPri_Img,pcPri_ShowImg) VALUES ('Low','priority_low.gif',1)

INSERT INTO pcFStatus (pcFStat_Name,pcFStat_Img,pcFStat_BGColor,pcFStat_ShowImg) VALUES ('Open','','#FFFF99',0)
INSERT INTO pcFStatus (pcFStat_Name,pcFStat_Img,pcFStat_BGColor,pcFStat_ShowImg) VALUES ('Closed','','#99FF99',0)

INSERT INTO pcFTypes (pcFType_Name,pcFType_Img,pcFType_ShowImg) VALUES ('Comment','',0)
INSERT INTO pcFTypes (pcFType_Name,pcFType_Img,pcFType_ShowImg) VALUES ('Suggestion','',0)

INSERT INTO pcTaskManager (pcTaskVersion,pcTaskNum,pcTaskComplete) VALUES ('2.6',1,1)
INSERT INTO pcTaskManager (pcTaskVersion,pcTaskNum,pcTaskComplete) VALUES ('2.6',2,1)
INSERT INTO pcTaskManager (pcTaskVersion,pcTaskNum,pcTaskComplete) VALUES ('2.6',3,1)
INSERT INTO pcTaskManager (pcTaskVersion,pcTaskNum,pcTaskComplete) VALUES ('2.6',4,1)
INSERT INTO pcTaskManager (pcTaskVersion,pcTaskNum,pcTaskComplete) VALUES ('2.6',5,1)
INSERT INTO pcTaskManager (pcTaskVersion,pcTaskNum,pcTaskComplete) VALUES ('2.6',6,1)

INSERT INTO echo (id, transaction_type) VALUES (1,'AS')

INSERT INTO concord (StoreID, StoreKey, CVV,testmode,Curcode,MethodName) VALUES ('','',1,1,'USD','Authorize')

INSERT INTO klix (ssl_merchant_id, ssl_pin, CVV, ssl_avs, testmode,ssl_user_id) VALUES ('','',0,1,0,'')

INSERT INTO tclink (idTCLink, CVV, TCLinkCheck, TCLinkCheckPending, TCLinkecheck, TCCurcode, TranType) VALUES (1, 1, 1, 1, 'TCLink Check', 'usd', 'sale')

INSERT INTO BluePay (idBluePay, BPTestmode, BPTransType, BPInterfaceType, BPCVC) VALUES (1,1,'SALE','API','YES')

INSERT INTO netbill (idNetbill,NBCVVEnabled,NBAVS,NBTranType,NetbillCheck) VALUES (1,0,0,'S',0)

INSERT INTO pcRevSettings (pcRS_RatingType,pcRS_MainRateTxt1,pcRS_MainRateTxt2,pcRS_MainRateTxt3,pcRS_SubRateTxt1,pcRS_SubRateTxt2,pcRS_MaxRating,pcRS_Img1,pcRS_Img2,pcRS_Img3,pcRS_Img4,pcRS_Img5,pcRS_Active,pcRS_ShowRatSum,pcRS_RevCount,pcRS_NeedCheck,pcRS_LockPost,pcRS_PostCount,pcRS_CalMain) VALUES (0, 'liked this product','Like it','Don''t like it','Good','Bad',5,'smileygreen.gif','smileyred.gif','fullstar.gif','halfstar.gif','emptystar.gif',0,1,5,1,2,1,0)

INSERT INTO pcRevFields (pcRF_Name,pcRF_Type,pcRF_Active,pcRF_Required,pcRF_Order) VALUES ('Customer name',0,1,1,0)
INSERT INTO pcRevFields (pcRF_Name,pcRF_Type,pcRF_Active,pcRF_Required,pcRF_Order) VALUES ('Title',0,1,1,0)

INSERT INTO shipAlert (shipExists) VALUES (1)

INSERT INTO pcPay_FastCharge (pcPay_FAC_ID, pcPay_FAC_ATSID, pcPay_FAC_TransType, pcPay_FAC_CVV,pcPay_FAC_Checking,pcPay_FAC_CheckPending) VALUES (1,'',0,0,0,0)

INSERT INTO pcPay_ACHDirect (pcPay_ACH_Id, pcPay_ACH_MerchantID, pcPay_ACH_PWD, pcPay_ACH_TransType, pcPay_ACH_TestMode, pcPay_ACH_CVV, pcPay_ACH_CardTypes) VALUES (1, '', '', 'AUTH', 1, 0, '')

INSERT INTO pcPay_NETOne (pcPay_NETOne_ID, pcPay_NETOne_MID, pcPay_NETOne_Mkey, pcPay_NETOne_Tcode, pcPay_NETOne_CVV, pcPay_NETOne_CardTypes) VALUES (1, '', '', 'AUTH', 1, '')

INSERT INTO pcPay_Centinel (pcPay_Cent_ID) VALUES (1)

INSERT INTO pcPay_GestPay (pcPay_GestPay_Id) VALUES (1)

INSERT INTO pcPay_EPN (pcPay_EPN_ID) VALUES (1)

INSERT INTO pcPay_TripleDeal (pcPay_TD_ID,pcPay_TD_Profile,pcPay_TD_ClientLang,pcPay_TD_PayPeriod,pcPay_TD_TestMode) VALUES (1,'standard','en',1,1)

INSERT INTO pcPay_ParaData (pcPay_ParaData_ID,pcPay_ParaData_TransType) VALUES (1,'SALE')

INSERT INTO pcPay_Moneris (pcPay_Moneris_ID) VALUES (1)

INSERT INTO pcPay_HSBC (pcPay_HSBC_ID) VALUES (1)

INSERT INTO pcPay_PayPal (pcPay_PayPal_ID) VALUES (1)

INSERT INTO pcPay_PaymentExpress (pcPay_PaymentExpress_ID, pcPay_PaymentExpress_TransType, pcPay_PaymentExpress_Username, pcPay_PaymentExpress_Password,pcPay_PaymentExpress_TestMode,pcPay_PaymentExpress_Cvc2,pcPay_PaymentExpress_ReceiptEmail,pcPay_PaymentExpress_TestUsername,pcPay_PaymentExpress_AVS) VALUES (1, 'SALE', ' ', ' ',1,1,' ',' ',1)

INSERT INTO pcNewArrivalsSettings (pcNAS_NewArrCount, pcNAS_Style, pcNAS_PageDesc, pcNAS_NDays, pcNAS_NotForSale, pcNAS_OutOfStock, pcNAS_SKU, pcNAS_ShowImg) VALUES (9, 'h', '', 30, 0, 0, 0, 0)

INSERT INTO pcStoreSettings ( pcStoreSettings_CompanyName, pcStoreSettings_CompanyAddress, pcStoreSettings_CompanyZip, pcStoreSettings_CompanyCity, pcStoreSettings_CompanyState, pcStoreSettings_CompanyCountry, pcStoreSettings_CompanyLogo, pcStoreSettings_QtyLimit, pcStoreSettings_AddLimit, pcStoreSettings_Pre, pcStoreSettings_CustPre, pcStoreSettings_CatImages, pcStoreSettings_ShowStockLmt, pcStoreSettings_OutOfStockPurchase, pcStoreSettings_Cursign, pcStoreSettings_DecSign, pcStoreSettings_DivSign, pcStoreSettings_DateFrmt, pcStoreSettings_MinPurchase, pcStoreSettings_WholesaleMinPurchase, pcStoreSettings_URLredirect, pcStoreSettings_SSL, pcStoreSettings_SSLUrl, pcStoreSettings_IntSSLPage, pcStoreSettings_PrdRow, pcStoreSettings_PrdRowsPerPage, pcStoreSettings_CatRow, pcStoreSettings_CatRowsPerPage, pcStoreSettings_BType, pcStoreSettings_StoreOff, pcStoreSettings_StoreMsg, pcStoreSettings_WL, pcStoreSettings_TF, pcStoreSettings_orderLevel, pcStoreSettings_DisplayStock, pcStoreSettings_HideCategory, pcStoreSettings_AllowNews, pcStoreSettings_NewsCheckOut, pcStoreSettings_NewsReg, pcStoreSettings_NewsLabel, pcStoreSettings_PCOrd, pcStoreSettings_HideSortPro, pcStoreSettings_DFLabel, pcStoreSettings_DFShow, pcStoreSettings_DFReq, pcStoreSettings_TFLabel, pcStoreSettings_TFShow, pcStoreSettings_TFReq, pcStoreSettings_DTCheck, pcStoreSettings_DeliveryZip, pcStoreSettings_OrderName, pcStoreSettings_HideDiscField, pcStoreSettings_AllowSeparate, pcStoreSettings_ReferLabel, pcStoreSettings_ViewRefer, pcStoreSettings_RefNewCheckout, pcStoreSettings_RefNewReg, pcStoreSettings_BrandLogo, pcStoreSettings_BrandPro, pcStoreSettings_RewardsActive, pcStoreSettings_RewardsIncludeWholesale, pcStoreSettings_RewardsPercent, pcStoreSettings_RewardsLabel, pcStoreSettings_RewardsReferral, pcStoreSettings_RewardsFlat, pcStoreSettings_RewardsFlatValue, pcStoreSettings_RewardsPerc, pcStoreSettings_RewardsPercValue, pcStoreSettings_XML, pcStoreSettings_QDiscountType, pcStoreSettings_BTOdisplayType, pcStoreSettings_BTOOutofStockPurchase, pcStoreSettings_BTOShowImage, pcStoreSettings_BTOQuote, pcStoreSettings_BTOQuoteSubmit, pcStoreSettings_BTOQuoteSubmitOnly, pcStoreSettings_BTODetLinkType, pcStoreSettings_BTODetTxt, pcStoreSettings_BTOPopWidth, pcStoreSettings_BTOPopHeight, pcStoreSettings_BTOPopImage, pcStoreSettings_ConfigPurchaseOnly, pcStoreSettings_ShowSKU, pcStoreSettings_ShowSmallImg, pcStoreSettings_Terms, pcStoreSettings_TermsLabel, pcStoreSettings_TermsCopy, pcStoreSettings_HideRMA, pcStoreSettings_ShowHD, pcStoreSettings_StoreUseToolTip, pcStoreSettings_ErrorHandler, pcStoreSettings_AllowCheckoutWR, pcStoreSettings_ViewPrdStyle, pcStoreSettings_TermsShown, pcStoreSettings_DisableGiftRegistry, pcStoreSettings_DisableDiscountCodes,pcStoreSettings_PinterestDisplay,pcStoreSettings_PinterestCounter)	VALUES ('','','','','','AL','',0,0,0,0,0,0,0,'$','.',',','MM/DD/YY',0,0,'','','','0',3,3,3,3,'h','0','',-1,-1,'0',0,0,0,0,0,'',0,0,'','','','','','','','','0','0','','',0,0,0,0,0,0,0,0,'',0,0,0,0,0,'',0,0,0,0,0,0,0,0,'',0,0,0,0,-1,-1,0,'','',0,1,1,1,1,'c',0,0,1,0,'none')

INSERT INTO pcPay_SkipJack (pcPay_SkipJack_ID, pcPay_SkipJack_SerialNumber, pcPay_SkipJack_TestMode, pcPay_SkipJack_Cvc2) VALUES (1,'',1,0)

INSERT INTO pcPay_SecPay (pcPay_SecPay_ID, pcPay_SecPay_TransType, pcPay_SecPay_Username, pcPay_SecPay_Password,pcPay_SecPay_TestMode,pcPay_SecPay_Cvc2,pcPay_SecPay_AVS) VALUES (1, 'SALE', ' ', ' ',1,1,1)

INSERT INTO pcPay_Chronopay (CP_id, CP_ProdID, CP_Currency, CP_testmode) VALUES (1, '', '', 'YES')

INSERT INTO pcPay_eMerchant (pcPay_eMerch_ID, pcPay_eMerch_MerchantID, pcPay_eMerch_PaymentKey, pcPay_eMerch_CVD,	pcPay_eMerch_TransType, pcPay_eMerch_CardType, pcPay_eMerch_TestMode) VALUES (1, '', '', 0, 'P', '', 0)

INSERT INTO pcAmazonSettings (pcAmzSet_prdIDType,pcAmzSet_icondition,pcAmzSet_price,pcAmzSet_willshipout,pcAmzSet_expship,pcAmzSet_marketplace) VALUES (1,11,0,1,'y','y')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('AT','Austria')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('BE','Belgium')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('BG','Bulgaria')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('CY','Cyprus')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('CZ','Czech Republic')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('DK','Denmark')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('EE','Estonia')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('FI','Finland')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('FR','France')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('DE','Germany')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('GR','Greece')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('HU','Hungary')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('IE','Ireland')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('IT','Italy')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('LV','Latvia')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('LT','Lithuania')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('LU','Luxembourg')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('MT','Malta')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('NL','Netherlands')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('PL','Poland')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('PT','Portugal')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('RO','Romania')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('SK','Slovakia')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('SI','Slovenia')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('ES','Spain')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('SE','Sweden')

INSERT INTO pcVATCountries (pcVATCountry_Code,pcVATCountry_State) VALUES ('GB','United Kingdom')


INSERT INTO pcPay_EIG (pcPay_EIG_ID, pcPay_EIG_CVV, pcPay_EIG_TestMode) VALUES (1,0,1)
End
Go

/****** Object:  Table [dbo].[verisign_pfp]    Script Date: 1/10/2012 17:12:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[verisign_pfp]') AND OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
CREATE TABLE [dbo].[verisign_pfp](
	[id] [int] NOT NULL,
	[v_URL] [nvarchar](150) NULL,
	[v_Type] [nvarchar](150) NULL,
	[v_User] [nvarchar](150) NULL,
	[v_Partner] [nvarchar](150) NULL,
	[v_Password] [nvarchar](50) NULL,
	[v_Vendor] [nvarchar](50) NULL,
	[v_Tender] [nvarchar](150) NULL,
	[pfl_testmode] [nvarchar](50) NULL,
	[pfl_transtype] [nvarchar](50) NULL,
	[pfl_CSC] [nvarchar](50) NULL,
 CONSTRAINT [aaaaaverisign_pfp_PK] PRIMARY KEY NONCLUSTERED 
(
	[id] ASC
) ON [PRIMARY]
) ON [PRIMARY]
END
GO
/****** Object:  Default [DF__verisign_pfp__id__6B0FDBE9]    Script Date: 1/10/2012 17:12:23 ******/
IF Not EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF__verisign_pfp__id__6B0FDBE9]') AND type = 'D')
BEGIN
IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[DF__verisign_pfp__id__6B0FDBE9]') AND type = 'D')
BEGIN
ALTER TABLE [dbo].[verisign_pfp] ADD  CONSTRAINT [DF__verisign_pfp__id__6B0FDBE9]  DEFAULT ((0)) FOR [id]
END


END
GO
INSERT INTO verisign_pfp (id, v_URL,v_Type,v_User,v_Partner,v_Password,v_Vendor,v_Tender,pfl_testmode,pfl_transtype,pfl_CSC) VALUES (1,'https://test-payflow.verisign.com','A','user','VeriSign','password','vendor','TENDER=C&ZIP=12345&COMMENT1=ASP/COM Test','YES','A','YES')
