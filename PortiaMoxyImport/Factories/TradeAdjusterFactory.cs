using PortiaMoxyImport.Entities;
using PortiaMoxyImport.Enum;
using PortiaMoxyImport.Services;
using System;
using System.Collections.Generic;

public class TradeAdjusterFactory
{
    private readonly ITradeAdjuster _buyUsdNonUsdAdjuster;
    private readonly ITradeAdjuster _buyNonUsdUsdAdjuster;
    private readonly ITradeAdjuster _sellUsdNonUsdAdjuster;
    private readonly ITradeAdjuster _sellNonUsdUsdAdjuster;
    private readonly ITradeAdjuster _buyNonUsdNonUsdAdjuster;
    private readonly ITradeAdjuster _sellNonUsdNonUsdAdjuster;

    public TradeAdjusterFactory(
        ITradeAdjuster buyUsdNonUsdAdjuster,
        ITradeAdjuster buyNonUsdUsdAdjuster,
        ITradeAdjuster sellUsdNonUsdAdjuster,
        ITradeAdjuster sellNonUsdUsdAdjuster,
        ITradeAdjuster buyNonUsdNonUsdAdjuster,
        ITradeAdjuster sellNonUsdNonUsdAdjuster)
    {
        _buyUsdNonUsdAdjuster = buyUsdNonUsdAdjuster;
        _buyNonUsdUsdAdjuster = buyNonUsdUsdAdjuster;
        _sellUsdNonUsdAdjuster = sellUsdNonUsdAdjuster;
        _sellNonUsdUsdAdjuster = sellNonUsdUsdAdjuster;
        _buyNonUsdNonUsdAdjuster = buyNonUsdNonUsdAdjuster;
        _sellNonUsdNonUsdAdjuster = sellNonUsdNonUsdAdjuster;
    }

    public ITradeAdjuster GetAdjuster(NTFXTradeDTO trade)
    {
        var side = ParseSide(trade.BuySell);
        var baseBucket = BucketCurrency(trade.Currency);
        var otherBucket = BucketCurrency(trade.OtherCurrency);

        if (side == TradeSide.Buy && baseBucket == CurrencyBucket.Usd && otherBucket == CurrencyBucket.NonUsd)
            return _buyUsdNonUsdAdjuster;

        if (side == TradeSide.Buy && baseBucket == CurrencyBucket.NonUsd && otherBucket == CurrencyBucket.Usd)
            return _buyNonUsdUsdAdjuster;

        if (side == TradeSide.Sell && baseBucket == CurrencyBucket.Usd && otherBucket == CurrencyBucket.NonUsd)
            return _sellUsdNonUsdAdjuster;

        if (side == TradeSide.Sell && baseBucket == CurrencyBucket.NonUsd && otherBucket == CurrencyBucket.Usd)
            return _sellNonUsdUsdAdjuster;

        if (side == TradeSide.Buy && baseBucket == CurrencyBucket.NonUsd && otherBucket == CurrencyBucket.NonUsd)
            return _buyNonUsdNonUsdAdjuster;

        if (side == TradeSide.Sell && baseBucket == CurrencyBucket.NonUsd && otherBucket == CurrencyBucket.NonUsd)
            return _sellNonUsdNonUsdAdjuster;

        throw new ApplicationException(
            $"No adjuster implemented for combination: Side={side}, Base={baseBucket}, Other={otherBucket}");
    }

    private static TradeSide ParseSide(string buySell)
    {
        if (string.Equals(buySell, "B", StringComparison.OrdinalIgnoreCase))
            return TradeSide.Buy;
        if (string.Equals(buySell, "S", StringComparison.OrdinalIgnoreCase))
            return TradeSide.Sell;

        throw new ApplicationException("Unknown BuySell flag: " + buySell);
    }

    private static CurrencyBucket BucketCurrency(string currency)
    {
        if (currency == null)
            throw new ArgumentNullException(nameof(currency));

        return currency.Equals("USD", StringComparison.OrdinalIgnoreCase)
            ? CurrencyBucket.Usd
            : CurrencyBucket.NonUsd;
    }
}
