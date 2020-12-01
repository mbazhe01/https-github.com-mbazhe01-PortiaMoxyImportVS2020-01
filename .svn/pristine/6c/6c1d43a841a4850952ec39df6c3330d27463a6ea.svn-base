USE [Moxy86]
GO
/****** Object:  StoredProcedure [dbo].[usp_GetCrossRate]    Script Date: 10/22/2015 15:00:04 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER proc [dbo].[usp_GetCrossRate]
( @tradematchid int,
   @portfolio varchar(50))
as

	select OrderUserDef2, m.SecType,CONVERT(varchar(10), a.TradeDate, 101) as TradeDate
	from dbo.MoxyOrders o, dbo.MoxyAllocation a,	MxSec.SecMaster m   
	where o.OrderID = a.OrderID and
			a.TradeMatchID = @tradematchid and
			m.SecKey  = a.SecKey and
			a.PortID = @portfolio
			
-- Testing:			usp_GetCrossRate	1011927

