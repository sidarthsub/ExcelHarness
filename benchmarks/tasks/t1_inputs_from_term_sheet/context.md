# User context

You are a VC associate preparing the Series A model inputs for NorthStar
Robotics. The term sheet is final. You want a clean Inputs sheet that
downstream sheets will reference.

## Defaults you'd pick

- Post-money = pre-money + round size. Do NOT hardcode the $100M figure
  from the term sheet; derive it.
- % of round = each investor amount / total round size. Also derive.
- Option pool target is 12% post-money, implemented pre-money-inclusion
  (existing holders dilute). But you only need the target % as an input
  on this sheet; the math lives on the Series A sheet.
- Use simple numeric formatting — `$#,##0`, `0.00%`, `#,##0`.

## Things you do NOT care about

- Cell colors beyond section headers.
- Exact column widths.
- Whether dividend is in decimal or percentage storage as long as it renders.
- Whether liquidation preference is stored as `1.0` or `"1.0x"`.
