# User context

You are an associate modeling a liquidation waterfall for a preferred
equity holder at exit.

## Defaults you would pick

- 1x non-participating means the LP picks `MAX(preference, conversion)`
  — whichever gives them more. They do not get both.
- Common holders get the residual of the exit value after the LP is paid.
- Don't round. Fractional dollars are fine.
- Formula-only. Every answer cell must be a formula referencing the
  input cells — no pasted literals.

## Things you do NOT know or care about

- Cell formatting or number precision beyond a handful of decimals.
- Whether formulas use direct cell refs or named ranges — either is fine.
- Additional preferred tranches, liquidation stacks, or participation
  caps — this is a single-class 1x non-participating.
