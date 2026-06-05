"""Tests for QAEngine module."""

import pytest
from datetime import date

from audit_framework.models import (
    DataType,
    FailAction,
    QARule,
    RuleType,
    Severity,
)
from audit_framework.qa_engine import QAEngine, QAException


def test_qa_exception_creation():
    """Test creating QA exception."""
    exc = QAException(
        exception_id="EXC_000001",
        rule_id="RULE_001",
        field_name="Amount",
        extracted_value=5000,
        reason="Amount exceeds maximum of 1000",
        severity=Severity.HIGH,
        fail_action=FailAction.FLAG,
    )

    assert exc.exception_id == "EXC_000001"
    assert exc.status == "PENDING"
    assert exc.severity == Severity.HIGH


def test_qa_exception_to_dict():
    """Test converting exception to dictionary."""
    exc = QAException(
        exception_id="EXC_000001",
        rule_id="RULE_001",
        field_name="Amount",
        extracted_value=5000,
        reason="Amount exceeds maximum",
        severity=Severity.MEDIUM,
        fail_action=FailAction.FLAG,
    )

    exc_dict = exc.to_dict()
    assert exc_dict["ExceptionID"] == "EXC_000001"
    assert exc_dict["Status"] == "PENDING"


def test_qa_engine_range_rule():
    """Test QA engine applying range rule."""
    qa_engine = QAEngine()

    rule = QARule(
        rule_id="RULE_RANGE_001",
        rule_name="Amount Range Check",
        field_name="Amount",
        rule_type=RuleType.RANGE,
        rule_definition="0:1000",
        fail_action=FailAction.FLAG,
        severity=Severity.MEDIUM,
    )

    # Value within range
    exc = qa_engine.apply_rule(rule, 500, {"Amount": 500})
    assert exc is None

    # Value exceeds range
    exc = qa_engine.apply_rule(rule, 2000, {"Amount": 2000})
    assert exc is not None
    assert "exceeds maximum" in exc.reason


def test_qa_engine_lookup_rule():
    """Test QA engine applying lookup rule."""
    qa_engine = QAEngine()

    rule = QARule(
        rule_id="RULE_LOOKUP_001",
        rule_name="Currency Lookup",
        field_name="Currency",
        rule_type=RuleType.LOOKUP,
        rule_definition="USD,EUR,GBP",
        fail_action=FailAction.FLAG,
        severity=Severity.HIGH,
    )

    # Valid value
    exc = qa_engine.apply_rule(rule, "USD", {"Currency": "USD"})
    assert exc is None

    # Invalid value
    exc = qa_engine.apply_rule(rule, "JPY", {"Currency": "JPY"})
    assert exc is not None
    assert "not in allowed list" in exc.reason


def test_qa_engine_exception_decision():
    """Test recording QA exception decisions."""
    qa_engine = QAEngine()

    # Create an exception
    exc = QAException(
        exception_id="EXC_000001",
        rule_id="RULE_001",
        field_name="Amount",
        extracted_value=2000,
        reason="Amount exceeds maximum",
        severity=Severity.HIGH,
        fail_action=FailAction.FLAG,
    )
    qa_engine.exceptions.append(exc)

    # Accept exception
    qa_engine.accept_exception("EXC_000001", "john.doe")
    assert exc.status == "ACCEPTED"
    assert exc.reviewed_by == "john.doe"

    # Override exception
    qa_engine.override_exception(
        "EXC_000001", override_value=1000, justification="Corrected amount", reviewed_by="jane.doe"
    )
    assert exc.status == "OVERRIDDEN"
    assert exc.override_value == 1000


def test_qa_engine_summary():
    """Test QA engine summary."""
    qa_engine = QAEngine()

    # Create exceptions
    for i in range(3):
        exc = QAException(
            exception_id=f"EXC_{i:06d}",
            rule_id="RULE_001",
            field_name="Amount",
            extracted_value=i * 1000,
            reason="Test exception",
            severity=Severity.MEDIUM,
            fail_action=FailAction.FLAG,
        )
        qa_engine.exceptions.append(exc)

    summary = qa_engine.to_dict()
    assert summary["summary"]["total"] == 3
    assert summary["summary"]["pending"] == 3
