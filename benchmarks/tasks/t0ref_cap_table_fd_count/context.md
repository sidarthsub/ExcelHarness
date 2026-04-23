# User context

You are an associate verifying a cap table before a board meeting. You want
a simple total fully diluted share count, counting every line item in the
cap table: common, all preferred classes, options (granted and unallocated),
and warrants.

## Defaults you would pick

- Sum **ALL** rows — common + all preferred classes + options (granted +
  unallocated) + warrants.
- Use a **formula**, not a hardcoded literal.
- Importing the reference Cap Table sheet into the model workbook is fine;
  referencing shared data via a copied sheet is acceptable.

## Things you do NOT know or care about

- Cell formatting beyond a standard integer share count.
- Whether the sum uses a full-column reference (`C:C`) or a bounded range
  (`C2:C10`) — either is fine.
- Whether there are helper cells or intermediate rows alongside the answer.
