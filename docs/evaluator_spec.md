# Deterministic Evaluator Spec

## Summary

This document specifies a deterministic machine evaluator for workbook logic and completeness. The evaluator runs from dump artifacts produced by `dump.py`, covers whole sheets and whole workbooks, and serves as the primary logic/completeness reviewer. The existing prompt evaluator remains in place for visual review and as a fallback for unsupported logic/features.

This is a v1 implementation spec. It is intended for an engineer or agent building the evaluator, not for end users.

Goals:
- Evaluate full modeled workbooks quickly enough to run every harness cycle
- Compute supported formulas deterministically from dump artifacts only
- Detect broken references, wrong sheet sets, unsupported formulas, and non-converged iterative logic
- Reuse the current JSON verdict shape so the harness can consume the result without a second translation layer

Non-goals:
- Full Excel compatibility
- Pixel-perfect visual grading
- Arbitrary workbook automation

## Inputs And Interfaces

### Required inputs

The evaluator must consume the following artifacts:
- `model_spec.json`
- `evals/formulas_json/workbook.json`
- `evals/formulas_json/<sheet>.json` for each sheet
- `evals/formulas/<sheet>.txt` as optional human-readable supporting evidence
- `evals/styles/<sheet>.txt` and screenshots only for fallback/debugging or future extensions

The evaluator must not open Excel and must not depend on a live spreadsheet runtime.

### Structured formula dump contract

The evaluator relies on `formulas_json` as the stable machine interface.

`workbook.json` must provide:
- workbook name
- ordered sheet list
- function inventory (`functions_used`)
- iterative component inventory (`iterative_components`)

Each per-sheet JSON file must provide:
- `sheet`
- `formula_cells[]`

Each `formula_cells[]` entry must include:
- `cell`
- `formula`
- `cached_value`
- `functions[]`
- `references[]`
- `dependencies[]`
- `tokens[]`
- `iterative_component`

The evaluator may use `cached_value` as:
- initialization seed for iterative solving
- advisory comparison for unsupported formulas

It must not treat cached values as authoritative computed truth for supported logic.

## Evaluator Scope And Boundaries

### Machine evaluator responsibilities

The deterministic evaluator is the primary source of truth for:
- sheet presence and required output set
- dependency graph integrity
- formula parsing and function support classification
- deterministic value computation for supported formulas
- convergence status for iterative components
- hardcoded-vs-linked source checks where the spec requires live linkage
- placeholder/zero-row hygiene checks
- totals, percentages, and consistency checks across full modeled sheets

### Prompt evaluator responsibilities

The prompt evaluator remains responsible for:
- visual fidelity
- unsupported formulas or unsupported workbook features
- semantic finance/accounting/legal judgment that cannot be derived from artifacts

### Architectural rule

The machine evaluator is not an Excel clone. If a workbook feature falls outside the supported surface, the evaluator must classify it explicitly and route it to fallback rather than approximate it silently.

## Whole-Workbook Evaluation Strategy

### Evaluation unit

The evaluator runs across the entire workbook and every formula-bearing sheet. There is no region authoring system in v1.

Why full-sheet evaluation is preferred:
- more generalizable
- avoids manual maintenance of watch regions
- catches drift anywhere in modeled sheets
- fast-path supported functions are common enough to justify whole-sheet execution

### Sheet classes

The evaluator must treat sheets differently by role:

#### Copied source sheets

Checks are lighter:
- required sheet exists
- copied sheet/tab name matches spec
- no unintended rebuild when the spec says copy
- no unexpected modeled formulas or formatting drift where the copy contract is strict

#### Modeled sheets

Checks are full-sheet:
- all formula cells parsed or explicitly classified unsupported
- dependency graph is valid
- supported formulas are executed
- iterative components converge or are reported as non-converged
- totals and percentage checks run
- placeholder rows and visible zero-only junk rows are flagged
- unsupported cell inventory is recorded

## Engine Design

## High-level architecture

The evaluator has two layers:
- expression evaluator
- workbook dependency executor

### Expression evaluator

Responsibilities:
- parse tokenized formulas from `tokens[]`
- normalize refs, ranges, literals, and function calls
- evaluate supported formulas into deterministic numeric/text/date results

### Workbook dependency executor

Responsibilities:
- build the formula dependency graph from `dependencies[]`
- topologically evaluate acyclic formulas
- solve SCCs/iterative components separately
- classify every formula cell as supported-converged, supported-nonconverged, unsupported, or parse/graph-failure

## Supported Formula Surface (v1)

The v1 supported surface is intentionally broad but finite.

### Core language
- arithmetic: `+`, `-`, `*`, `/`, `^`
- comparisons: `=`, `<>`, `<`, `<=`, `>`, `>=`
- unary sign
- parentheses
- string concatenation where represented in the dump
- direct refs and cross-sheet refs
- rectangular ranges

### Logical functions
- `IF`
- `AND`
- `OR`
- `NOT`

### Aggregates and conditional aggregates
- `SUM`
- `SUMIF`
- `SUMIFS`
- `COUNTIF`
- `COUNTIFS`
- `MIN`
- `MAX`

### Numeric transforms
- `ROUND`
- `ROUNDUP`
- `ROUNDDOWN`
- `ABS`

### Date functions
- `DATE`
- `DAY`
- `MONTH`
- `YEAR`
- `EDATE`
- `EOMONTH`
- `YEARFRAC`

### Lookup functions
- `INDEX`
- `MATCH`
- `XMATCH`
- `XLOOKUP`
- `VLOOKUP`
- `HLOOKUP`

### Text assembly
- `CONCATENATE`
- `CONCAT`
- `TEXTJOIN`

### Cash-flow / finance functions
- `IRR`
- `XIRR`
- `NPV`
- `XNPV`

## Unsupported Surface (v1)

The evaluator must classify the following as unsupported unless the dump format is later extended to normalize them:
- `OFFSET`
- `INDIRECT`
- external workbook links
- VBA/macros
- tables / structured references not already normalized into plain refs
- dynamic-array spill behavior beyond explicitly supported functions
- volatile workbook behavior not representable from the dump artifacts

Unsupported handling rules:
- mark unsupported cells explicitly
- never guess a value
- use cached values only for advisory/debug output
- if unsupported logic is material, route to prompt fallback

## Dependency Graph And Execution

### Graph construction

The evaluator must construct a workbook-level directed graph:
- each formula cell is a node
- edges point from formula cell to precedent formula cells
- refs to non-formula cells are treated as leaf inputs

The evaluator must detect:
- missing sheet refs
- missing cells or malformed refs
- self-reference
- SCCs

### Acyclic execution

For acyclic portions:
- evaluate in topological order
- propagate typed values
- stop and classify parse/graph failures deterministically

### Iterative execution

The evaluator must use `iterative_component` from the dump when present. If absent, it may derive SCCs from the graph.

For each iterative component:
- initialize each cell from cached value if available, else `0`
- repeatedly recompute all supported cells in the component
- stop on convergence or iteration cap

Default v1 iterative settings:
- max iterations: `200`
- absolute tolerance: `1e-8`
- relative tolerance: `1e-7`

Convergence rule:
- a component is converged only if every supported cell’s delta falls below the absolute or relative tolerance

Non-convergence handling:
- classify cells as `supported-nonconverged`
- emit a critical issue if the non-converged component affects required modeled outputs

## XIRR And Numerical Solvers

`XIRR` must be supported in v1.

Requirements:
- deterministic numerical solve with fixed defaults
- no hidden fallback to cached value as a real answer

Default v1 `XIRR` solver settings:
- initial guess: `0.1`
- max iterations: `100`
- tolerance: `1e-8`

Failure modes:
- no valid root
- invalid cash-flow sign pattern
- numerical divergence

Handling:
- classify as non-converged or unsupported
- never fabricate a rate
- include the reason in the issue details

## Classification Model

Every formula cell must end in one of these states:
- `supported-converged`
- `supported-nonconverged`
- `unsupported`
- `parse-failure`
- `graph-failure`

The evaluator must produce workbook-level summaries:
- total supported cells
- total unsupported cells
- total non-converged cells
- function coverage counts
- iterative component convergence summary

## Checks Beyond Formula Execution

### Sheet completeness
- all required sheets exist
- copied sheets remain separate when the spec requires that
- required outputs are not materially renamed

### Source-linking / hardcoding
- if the spec requires source-linked behavior, hardcoded modeled values that can drift future outputs are critical
- the evaluator should compare formula-bearing vs literal cells in required modeled regions and flag suspect literals

### Totals and percentages
- totals rows should be formula-driven, not hardcoded
- percentage columns expected to sum to 1 or ~100% must be checked with tolerance
- obvious identity/circular garbage cells should be flagged

### Placeholder / zero-row hygiene
- rows with blank or zero labels and all-zero data should be flagged
- repeated zero-only placeholder rows in modeled output areas are critical if they materially degrade output completeness

## Verdict Model

The evaluator must return the existing JSON shape:

```json
{
  "logic": {
    "passed": true,
    "issues": []
  },
  "visual": {
    "grade": "A",
    "issues": []
  }
}
```

For the machine evaluator:
- `logic` is authoritative
- `visual` may be omitted internally and populated later by prompt evaluation, or passed through from the prompt evaluator depending on harness integration

## Severity Rules

### Critical

Mark an issue `critical` if it does any of the following:
- wrong required sheet set
- broken or unresolved references affecting outputs
- supported formula computes a result inconsistent with the graph/spec/source requirements
- required iterative component does not converge
- governed mechanic is implemented differently from explicit spec/source rules where this is expressible from artifacts
- hardcoded value replaces required source-linked logic and can change future outputs

### Warning

Mark an issue `warning` if it:
- does not change current or future outputs
- is structurally sloppy but numerically inert
- reflects maintainability or hygiene problems only

### Unsupported / advisory

Use `unsupported` or advisory classification internally for:
- cells outside the v1 supported surface
- ambiguous artifact-only cases that require prompt fallback

The harness may translate unsupported findings into:
- logic fail when unsupported logic is material
- prompt fallback when unsupported logic is allowed but not decisive

## Harness Combination Rules

The spec should define the combination policy clearly:

1. Run deterministic evaluator first
2. If logic is fully supported and passes, proceed
3. If unsupported logic is present:
   - route to prompt fallback
   - include unsupported inventory in fallback context
4. Visual evaluation remains prompt-based in v1
5. Final logic pass must not ignore deterministic criticals or non-convergence

The machine evaluator must be treated as primary for supported logic. The prompt evaluator must not silently override deterministic critical findings.

## Performance Targets

The evaluator is intended to run every harness cycle on the whole workbook.

v1 performance target:
- under `2s` on current Sidekick-sized workbooks for graph build + supported-cell execution
- under `5s` including iterative solving and report generation on current modeled workbooks

Performance design rules:
- parse once per workbook
- reuse structured dump dependencies instead of reparsing formulas from text
- solve SCCs only once per evaluation pass
- do not require region authoring or human configuration

## Test Plan

### Unit tests
- parser/expression tests for each supported function family
- range expansion and cross-sheet ref tests
- graph construction tests
- SCC detection tests

### Iterative solver tests
- simple 2-cell fixed-point loop
- share-price/share-count loop
- option-pool top-up loop
- deliberately non-converging loop

### Function coverage tests
- conditional aggregates
- lookup functions
- date math
- `IRR`
- `XIRR`

### Acceptance tests from real artifacts
- current `Series A` formulas
- current `Series B` formulas
- known SAFE pricing mismatch case from run `112300`
- copied-source sheet integrity case

### Acceptance criteria
- machine evaluator classifies the `112300` SAFE pricing bug as logic-critical
- machine evaluator detects unsupported functions/features explicitly rather than passing them
- machine evaluator reports non-convergence when iterative logic fails to settle
- evaluator runs on whole sheets/workbooks without region configuration

## Implementation Defaults

- File path: `docs/evaluator_spec.md`
- Primary substrate: `formulas_json`
- Whole-sheet evaluation by default
- Prompt evaluator remains responsible for visual QA in v1
- Prefer explicit unsupported/fallback behavior over best-effort emulation
