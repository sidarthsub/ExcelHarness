"""Tests for the strict Planner output schema."""
import json
from pathlib import Path

import pytest
from jsonschema import validate, ValidationError


SCHEMA_PATH = Path(__file__).parent.parent / "schemas" / "model_spec.schema.json"


@pytest.fixture
def schema():
    return json.loads(SCHEMA_PATH.read_text())


def test_valid_spec_passes(schema):
    spec = {
        "intent": "Build a 3-statement model for ConfigHub to answer what valuation the next round supports.",
        "sheets": [
            {
                "name": "Inputs",
                "purpose": "Centralized user-editable assumptions.",
                "must_contain": ["discount rate", "growth rates by year", "headcount plan"],
                "data_sources": ["/path/to/brief.pdf"],
            },
            {
                "name": "Revenue",
                "purpose": "Forward revenue build by product line.",
                "must_contain": ["product line breakout", "growth rate application"],
                "data_sources": ["Inputs"],
            },
        ],
        "constraints": ["fiscal year ends December 31", "USD only"],
        "out_of_scope": ["pre-2023 historicals", "monthly granularity"],
    }
    validate(instance=spec, schema=schema)


def test_cell_addresses_rejected(schema):
    spec = {
        "intent": "Build a model.",
        "sheets": [
            {
                "name": "Inputs",
                "purpose": "Assumptions.",
                "must_contain": ["discount rate in B7"],
                "data_sources": [],
            }
        ],
        "constraints": [],
        "out_of_scope": [],
    }
    # Schema doesn't reject this by content — but a linter step should.
    # The schema just prevents structural overprescription (no formulas/layout fields).
    # Keeping this test here as a placeholder; actual cell-address filtering is in
    # the self-review prompt, not the schema.
    validate(instance=spec, schema=schema)  # passes schema


def test_extra_fields_rejected(schema):
    """Overprescription via extra fields should fail validation."""
    spec = {
        "intent": "Build a model.",
        "sheets": [
            {
                "name": "Inputs",
                "purpose": "Assumptions.",
                "must_contain": [],
                "data_sources": [],
                "formulas": [{"address": "B7", "formula": "=0.08"}],  # <-- forbidden
            }
        ],
        "constraints": [],
        "out_of_scope": [],
    }
    with pytest.raises(ValidationError):
        validate(instance=spec, schema=schema)


def test_missing_required_fields_rejected(schema):
    spec = {"intent": "Build a model."}
    with pytest.raises(ValidationError):
        validate(instance=spec, schema=schema)
