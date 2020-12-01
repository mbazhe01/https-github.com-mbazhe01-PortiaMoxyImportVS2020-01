/****** Script for SelectTopNRows command from SSMS  ******/
use portia_test
go

alter proc usp_GetLastCrossRate 
(
	@cur1 varchar(10),
	@cur2 varchar(10),
	@tradedate datetime
)

as

	
	DECLARE @lastdate datetime
	select @lastdate = MAX(rate_date) from hfx where rate_date <= @tradedate

	SELECT  h1.[xrate]/h2.xrate AS CrossRate
	FROM [hfx] h1, cur c1, hfx h2, cur c2
	WHERE h1.rate_type = 1 and h1.currency = c1.ID and h1.rate_date = @lastdate
				    and c1.currency = @cur1  and
				h2.rate_type = 1 and h2.currency = c2.ID and h2.rate_date = @lastdate 
				and c2.currency = @cur2
				 
				
	
  
  -- Testing:   usp_GetLastCrossRate 'CAD', 'EUR', '7/05/14'