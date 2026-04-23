# User context

VC associate modeling a fresh Series A. You want the share-price math laid
out cleanly so downstream sheets can reference it.

## Defaults

- Share price uses the simple pre-money / pre-round FD formula.
- No option pool expansion in this calculation — the term sheet explicitly
  defers pool work to Series B.
- Formula-only for derived values. Do not hardcode derived values — use
  references so the sheet stays live.

## Things you do NOT care about

- Cell formatting, number display precision.
- Whether the formula uses direct cell refs or named ranges.
- Whether intermediate cells sit near the final answer.
