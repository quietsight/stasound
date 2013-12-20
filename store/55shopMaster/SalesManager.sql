/****** Object:  Table [dbo].[pcSales_Pending]    Script Date: 06/06/2011 09:59:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[pcSales_Pending](
	[pcSP_ID] [int] IDENTITY(1,1) NOT NULL,
	[idProduct] [int] NULL,
	[pcSales_ID] [int] NULL,
 CONSTRAINT [PK_pcSales_Pending] PRIMARY KEY CLUSTERED 
(
	[pcSP_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[pcSales_Completed]    Script Date: 04/23/2011 16:03:48 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[pcSales_Completed](
	[pcSC_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcSales_ID] [int] NULL CONSTRAINT [DF_pcSales_Completed_pcSales_ID]  DEFAULT ((0)),
	[pcSC_Status] [int] NULL CONSTRAINT [DF_pcSales_Completed_pcSC_Status]  DEFAULT ((0)),
	[pcSC_StartedDate] [datetime] NULL,
	[pcSC_BUStartedDate] [datetime] NULL,
	[pcSC_BUComDate] [datetime] NULL,
	[pcSC_BUTotal] [int] NULL CONSTRAINT [DF_pcSales_Completed_pcSC_BUTotal]  DEFAULT ((0)),
	[pcSC_StoppedDate] [datetime] NULL,
	[pcSC_REStartedDate] [datetime] NULL,
	[pcSC_REComDate] [datetime] NULL,
	[pcSC_RETotal] [int] NULL CONSTRAINT [DF_pcSales_Completed_pcSC_RETotal]  DEFAULT ((0)),
	[pcSC_ComDate] [datetime] NULL,
	[pcSC_SaveName] [nvarchar](250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSC_SaveDesc] [nvarchar](1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSC_SaveTech] [nvarchar](2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSC_SaveIcon] [nvarchar](250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSC_Archived] [int] NULL CONSTRAINT [DF_pcSales_Completed_pcSC_Archived]  DEFAULT ((0)),
 CONSTRAINT [PK_pcSales_Completed] PRIMARY KEY CLUSTERED 
(
	[pcSC_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[pcSales_BackUp]    Script Date: 04/23/2011 16:03:41 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[pcSales_BackUp](
	[pcSB_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcSC_ID] [int] NULL CONSTRAINT [DF_pcSales_BackUp_pcSC_ID]  DEFAULT ((0)),
	[pcSales_ID] [int] NULL CONSTRAINT [DF_pcSales_BackUp_pcSales_ID]  DEFAULT ((0)),
	[IDProduct] [int] NULL CONSTRAINT [DF_pcSales_BackUp_IDProduct]  DEFAULT ((0)),
	[pcSales_TargetPrice] [int] NULL CONSTRAINT [DF_pcSales_BackUp_pcSales_Type]  DEFAULT ((0)),
	[pcSB_Price] [float] NULL CONSTRAINT [DF_pcSales_BackUp_pcSB_Price]  DEFAULT ((0)),
 CONSTRAINT [PK_pcSales_BackUp] PRIMARY KEY CLUSTERED 
(
	[pcSB_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
GO

/****** Object:  Table [dbo].[pcSales]    Script Date: 04/23/2011 16:03:38 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[pcSales](
	[pcSales_ID] [int] IDENTITY(1,1) NOT NULL,
	[pcSales_TargetPrice] [int] NULL CONSTRAINT [DF_pcSales_pcSales_TargetPrice]  DEFAULT ((0)),
	[pcSales_Type] [int] NULL CONSTRAINT [DF_pcSales_pcSales_Type]  DEFAULT ((0)),
	[pcSales_Relative] [int] NULL CONSTRAINT [DF_Table_1_pcSales_]  DEFAULT ((0)),
	[pcSales_Amount] [float] NULL CONSTRAINT [DF_pcSales_pcSales_Amount]  DEFAULT ((0)),
	[pcSales_Round] [int] NULL CONSTRAINT [DF_pcSales_pcSales_Round]  DEFAULT ((0)),
	[pcSales_Name] [nvarchar](250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSales_ImgURL] [nvarchar](250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSales_Desc] [nvarchar](1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSales_CreatedDate] [datetime] NULL,
	[pcSales_EditedDate] [datetime] NULL,
	[pcSales_Param1] [nvarchar](1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSales_Param2] [nvarchar](1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSales_Tech] [nvarchar](2000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[pcSales_Removed] [int] NULL CONSTRAINT [DF_pcSales_pcSales_Removed]  DEFAULT ((0)),
 CONSTRAINT [PK_pcSales] PRIMARY KEY CLUSTERED 
(
	[pcSales_ID] ASC
) ON [PRIMARY]
) ON [PRIMARY]
GO

/****** Object:  StoredProcedure [dbo].[uspActivatePrds]    Script Date: 04/23/2011 15:52:42 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspActivatePrds]    Script Date: 02/25/2011 22:48:15 ******/
CREATE PROCEDURE [dbo].[uspActivatePrds]
@SCID nvarchar(10) ,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query='UPDATE Products SET Products.active=Products.pcSC_ID,Products.pcSC_ID=0 FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ';'
	EXEC(@query)	
	
	SET @SMCount=@@ROWCOUNT

END
GO

/****** Object:  StoredProcedure [dbo].[uspAddCatPrices]    Script Date: 04/23/2011 15:52:42 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspAddCatPrices]    Script Date: 02/25/2011 22:48:41 ******/
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
		SET @query='INSERT INTO pcCC_Pricing (idcustomerCategory,IDProduct,pcCC_Price) SELECT ' + @IDCat + ',Products.idProduct,Round(Products.Price*' + @CAmount + ',2) FROM ' + @Param1 + ' WHERE (Products.idProduct NOT IN (SELECT idProduct FROM pcCC_Pricing WHERE idcustomerCategory=' + @IDCat + ')) AND ' + @Param2 + ';'
	ELSE
		SET @query='INSERT INTO pcCC_Pricing(idcustomerCategory,IDProduct,pcCC_Price) SELECT ' + @IDCat + ',Products.idProduct,WPrice = CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @CAmount + ',2) ELSE Round(Products.bToBPrice*' + @CAmount + ',2) END FROM ' + @Param1 + ' WHERE (Products.idProduct NOT IN (SELECT idProduct FROM pcCC_Pricing WHERE idcustomerCategory=' + @IDCat + ')) AND ' + @Param2 + ';'
		
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT

END
GO


/****** Object:  StoredProcedure [dbo].[uspBackUpPrices]    Script Date: 04/23/2011 15:52:43 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspBackUpPrices]    Script Date: 02/25/2011 22:49:18 ******/
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
	
	IF @TPrice='0'
		SET @query='INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT ' + @SCID + ',' + @SalesID + ',' + @TPrice + ',Products.idProduct,Products.Price FROM ' + @Param1 + ' WHERE ' + @Param2 + ';'
	ELSE
		BEGIN
		IF @TPrice='-1'
			SET @query='INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT ' + @SCID + ',' + @SalesID + ',' + @TPrice + ',Products.idProduct,Products.bToBPrice FROM ' + @Param1 + ' WHERE ' + @Param2 + ';'
		ELSE
			SET @query='INSERT INTO pcSales_BackUp (pcSC_ID,pcSales_ID,pcSales_TargetPrice,IDProduct,pcSB_Price) SELECT ' + @SCID + ',' + @SalesID + ',' + @TPrice + ',pcCC_Pricing.idProduct,pcCC_Pricing.pcCC_Price FROM pcCC_Pricing WHERE pcCC_Pricing.idcustomerCategory=' + @TPrice + ' AND (pcCC_Pricing.IdProduct IN (SELECT Products.IdProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + '));'
		END
		
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT

END
GO

/****** Object:  StoredProcedure [dbo].[uspChangePrices]    Script Date: 04/23/2011 15:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspChangePrices]    Script Date: 02/25/2011 22:49:56 ******/
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
	
	SET @query=''
	
	IF @CType='0'
	BEGIN
	
		IF @TPrice='0'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price*' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price*' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF @TPrice='-1'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',2) ELSE Round(Products.bToBPrice*' + @Amount + ',2) END'
			ELSE
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',0) ELSE Round(Products.bToBPrice*' + @Amount + ',0) END'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF (@TPrice<>'-1') AND (@TPrice<>'0')
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + '));'
		END
		
	END
	
	IF @CType='1'
	BEGIN
	
		IF @TPrice='0'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price-' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE Products SET Products.Price=Round(Products.Price-' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF @TPrice='-1'
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price-' + @Amount + ',2) ELSE Round(Products.bToBPrice-' + @Amount + ',2) END'
			ELSE
				SET @query='UPDATE Products SET Products.bToBPrice=CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price-' + @Amount + ',0) ELSE Round(Products.bToBPrice-' + @Amount + ',0) END'
			
			SET @query=@query + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF (@TPrice<>'-1') AND (@TPrice<>'0')
		BEGIN
			IF @CRound='0'
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price-' + @Amount + ',2)'
			ELSE
				SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=Round(pcCC_Pricing.pcCC_Price-' + @Amount + ',0)'
			
			SET @query=@query + ' WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + '));'
		END
		
	END
		
	IF @query<>''
	BEGIN
		EXEC(@query)
		SET @SMCount=@@ROWCOUNT
		
		SET @query='UPDATE Products SET Products.pcSC_ID=' + @SCID + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');' 
		EXEC(@query)
	END
	
	IF @CType='2'
	BEGIN
		
		SET @HasTmp=0
		
		IF @Relative='0'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				SET @query='SELECT Products.idProduct,NewPrice = Round(Products.Price*' + @Amount + ',2) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');' 
			ELSE
				SET @query='SELECT Products.idProduct,NewPrice = Round(Products.Price*' + @Amount + ',0) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF @Relative='-1'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				SET @query='SELECT Products.idProduct,NewPrice = CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',2) ELSE Round(Products.bToBPrice*' + @Amount + ',2) END INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');' 
			ELSE
				SET @query='SELECT Products.idProduct,NewPrice = CASE Products.bToBPrice WHEN 0 THEN Round(Products.Price*' + @Amount + ',0) ELSE Round(Products.bToBPrice*' + @Amount + ',0) END INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF @Relative='-2'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				SET @query='SELECT Products.idProduct,NewPrice = Round(Products.listPrice*' + @Amount + ',2) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');' 
			ELSE
				SET @query='SELECT Products.idProduct,NewPrice = Round(Products.listPrice*' + @Amount + ',0) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF @Relative='-3'
		BEGIN
			SET @HasTmp=1
			IF @CRound='0'
				SET @query='SELECT Products.idProduct,NewPrice = Round(Products.cost*' + @Amount + ',2) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');' 
			ELSE
				SET @query='SELECT Products.idProduct,NewPrice = Round(Products.cost*' + @Amount + ',0) INTO tmpSale1 FROM Products WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');'
		END
		
		IF @HasTmp=0
		BEGIN
			IF @CRound='0'
				SET @query='SELECT pcCC_Pricing.idProduct,NewPrice = Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',2) INTO tmpSale1 FROM pcCC_Pricing WHERE (pcCC_Pricing.idcustomerCategory=' + @Relative + ') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + '));' 
			ELSE
				SET @query='SELECT pcCC_Pricing.idProduct,NewPrice = Round(pcCC_Pricing.pcCC_Price*' + @Amount + ',0) INTO tmpSale1 FROM pcCC_Pricing WHERE (pcCC_Pricing.idcustomerCategory=' + @Relative + ') AND (pcCC_Pricing.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + '));'
		END
		
		EXEC(@query)
		
		IF @TPrice='0'
			SET @query='UPDATE Products SET Products.Price=tmpSale1.NewPrice FROM Products,tmpSale1 WHERE Products.IdProduct=tmpSale1.IDProduct;'

		
		IF @TPrice='-1'
			SET @query='UPDATE Products SET Products.bToBPrice=tmpSale1.NewPrice FROM Products,tmpSale1 WHERE Products.IdProduct=tmpSale1.IDProduct;'
		
		IF (@TPrice<>'-1') AND (@TPrice<>'0')
			SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=tmpSale1.NewPrice FROM pcCC_Pricing,tmpSale1 WHERE (pcCC_Pricing.idcustomerCategory=' + @TPrice + ') AND (pcCC_Pricing.IdProduct=tmpSale1.IDProduct);'
		
		EXEC(@query)
		SET @SMCount=@@ROWCOUNT
		
		DROP TABLE tmpSale1
		
		SET @query='UPDATE Products SET Products.pcSC_ID=' + @SCID + ' WHERE Products.IdProduct IN (SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ');' 
		EXEC(@query)
		
	END
	
END
GO

/****** Object:  StoredProcedure [dbo].[uspGetPrdCount]    Script Date: 04/23/2011 15:52:44 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspGetPrdCount]    Script Date: 02/25/2011 22:50:27 ******/
CREATE PROCEDURE [dbo].[uspGetPrdCount]
@Param1 nvarchar(1000) ,
@Param2 nvarchar(1000),
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query='SELECT Products.idProduct FROM ' + @Param1 + ' WHERE ' + @Param2 + ';'
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT

END
GO

/****** Object:  StoredProcedure [dbo].[uspGetUpdatedPrdCount]    Script Date: 04/23/2011 15:52:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspGetUpdatedPrdCount]    Script Date: 02/25/2011 22:51:04 ******/
CREATE PROCEDURE [dbo].[uspGetUpdatedPrdCount]
@SCID nvarchar(10) ,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query='SELECT idProduct FROM Products WHERE pcSC_ID=' + @SCID + ';'
	EXEC(@query)
	
	SET @SMCount=@@ROWCOUNT

END
GO

/****** Object:  StoredProcedure [dbo].[uspInActivatePrds]    Script Date: 04/23/2011 15:52:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspInActivatePrds]    Script Date: 02/25/2011 22:51:46 ******/
CREATE PROCEDURE [dbo].[uspInActivatePrds]
@SCID nvarchar(10) ,
@SMCount int Output
AS
BEGIN
	DECLARE @query varchar(8000)
	
	SET @SMCount=0
	
	SET NOCOUNT ON
	
	SET @query='UPDATE Products SET Products.pcSC_ID=Products.active,Products.active=0 FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ';'
	EXEC(@query)
		
	SET @SMCount=@@ROWCOUNT

END
GO

/****** Object:  StoredProcedure [dbo].[uspRestorePrices]    Script Date: 04/23/2011 15:52:45 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspRestorePrices]    Script Date: 02/25/2011 22:53:19 ******/
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
		SET @query='UPDATE Products SET Products.Price=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ';'
	
	IF @TPrice=-1
		SET @query='UPDATE Products SET Products.bToBPrice=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ';'
		
	IF @TPrice>0
		SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=pcSales_BackUp.pcSB_Price FROM pcCC_Pricing, pcSales_BackUp WHERE pcCC_Pricing.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcCC_Pricing.idcustomerCategory=' + @TPrice + ';'
		
	EXEC(@query)
	SET @SMCount=@@ROWCOUNT

END
GO

/****** Object:  StoredProcedure [dbo].[uspRmvBackedUpRecords]    Script Date: 04/23/2011 15:52:46 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspRmvBackedUpRecords]    Script Date: 02/25/2011 22:53:44 ******/
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
GO

/****** Object:  StoredProcedure [dbo].[uspRmvPrdFromSale]    Script Date: 04/23/2011 15:52:46 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/****** Object:  StoredProcedure [dbo].[uspRmvPrdFromSale]    Script Date: 02/25/2011 22:54:20 ******/
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
	
	SET @query='UPDATE Products SET Products.pcSC_ID=Products.active,Products.active=0 WHERE Products.IDProduct =' + @IDPrd + ';'
	EXEC(@query)
	
	SELECT TOP 1 @TPrice=pcSales_TargetPrice FROM pcSales_BackUp WHERE pcSC_ID=@SCID
	
	IF @TPrice=0
		SET @query='UPDATE Products SET Products.Price=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcSales_BackUp.IDProduct=' + @IDPrd + ';'
	
	IF @TPrice=-1
		SET @query='UPDATE Products SET Products.bToBPrice=pcSales_BackUp.pcSB_Price FROM Products, pcSales_BackUp WHERE Products.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcSales_BackUp.IDProduct=' + @IDPrd + ';'
		
	IF @TPrice>0
		SET @query='UPDATE pcCC_Pricing SET pcCC_Pricing.pcCC_Price=pcSales_BackUp.pcSB_Price FROM pcCC_Pricing, pcSales_BackUp WHERE pcCC_Pricing.IDProduct = pcSales_BackUp.IDProduct AND pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcCC_Pricing.idcustomerCategory=' + @TPrice + ' AND pcSales_BackUp.IDProduct=' + @IDPrd + ';'
		
	EXEC(@query)
	
	SET @query='UPDATE Products SET Products.active=Products.pcSC_ID,Products.pcSC_ID=0 WHERE Products.IDProduct =' + @IDPrd + ';'
	EXEC(@query)
	
	SET @query='DELETE FROM pcSales_BackUp WHERE pcSales_BackUp.pcSC_ID=' + @SCID + ' AND pcSales_BackUp.IDProduct=' + @IDPrd + ';'
	EXEC(@query)
	
	SET @query='UPDATE pcSales_Completed SET pcSC_BUTotal=(SELECT Count(*) FROM Products WHERE Products.pcSC_ID=' + @SCID + ') WHERE pcSales_Completed.pcSC_ID=' + @SCID + ';'
	EXEC(@query)
	
END
GO
