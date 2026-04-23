# User context

Associate running a quick DCF.

## Defaults

- Excel's `NPV` function excludes Y0 by convention, so use
  `=B10 + NPV(B3, B5:B9)`.
- Don't round — full floating-point precision is fine.

## Things you do NOT care about

- Cell formatting beyond a sensible currency format.
- Whether the formula uses ranges or individual cell refs, as long as it
  references the inputs rather than hardcoding numbers.
