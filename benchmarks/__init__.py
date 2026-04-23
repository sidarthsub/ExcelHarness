"""ExcelHarness benchmark suite.

Components:
- oracle: automated user-simulator that answers Planner clarifications
- grader: rubric-driven scoring of a candidate .xlsx
- runner: CLI that ties oracle + grader + (optional) builder together
- recalc: soffice-based formula recalculation shim
"""
