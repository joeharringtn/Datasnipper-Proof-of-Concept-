"""
Audit Framework - Enterprise-grade audit automation for DataSnipper integration.

A Python framework for orchestrating data extraction, validation, QA workflows,
and approval processes with full audit trail support.
"""

__version__ = "1.0.0"
__author__ = "Audit Framework Team"

from audit_framework.models import (
    ApprovalWorkflow,
    ConfigObject,
    DocumentSpec,
    QARule,
    SchemaField,
    TagDefinition,
)

__all__ = [
    "ConfigObject",
    "DocumentSpec",
    "TagDefinition",
    "QARule",
    "SchemaField",
    "ApprovalWorkflow",
]
