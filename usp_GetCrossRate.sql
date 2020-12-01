alter proc usp_GetCrossRate
( @tradematchid int)
as

	select OrderUserDef2
	from dbo.MoxyOrders o, dbo.MoxyAllocation a
	where o.OrderID = a.OrderID and
			a.TradeMatchID = @tradematchid
