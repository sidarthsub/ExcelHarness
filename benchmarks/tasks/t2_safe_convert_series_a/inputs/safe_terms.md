# SAFE notes outstanding — Rivet Labs

Two SAFEs are outstanding and must convert at the Series A close. Conversion
mechanics are described below.

## SAFE 1 — Harbor Fund
- Investment: **$2,000,000**
- Valuation cap: **$20,000,000** (pre-money cap)
- Discount: **0%** (none)

## SAFE 2 — Archer Angels
- Investment: **$1,500,000**
- Valuation cap: **$25,000,000** (pre-money cap)
- Discount: **20%**

## Conversion mechanics (canonical rule used for this model)

For each SAFE, compute three prices and pick the lowest to determine the
conversion share price:

- **cap price** = valuation_cap / pre_round_fully_diluted_shares
- **headline price** = round_pre_money / pre_round_fully_diluted_shares
- **discount price** = (1 − discount) × headline_price

SAFE conversion price = min(cap_price, discount_price).

SAFE shares issued = SAFE_investment / SAFE_conversion_price.

The SAFE investors receive Series A Preferred at their respective conversion
prices. New Series A investors buy at the headline price.

Pre-round fully diluted shares are the share count **before** any SAFE
conversion or option pool top-up. These are the 10,500,000 shares in the
input cap table.
