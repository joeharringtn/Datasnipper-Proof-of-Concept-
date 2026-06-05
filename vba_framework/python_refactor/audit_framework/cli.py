"""
CLI - Command-line interface and orchestration.

Provides user-facing commands for the entire audit workflow:
- load-engagement
- build-tags
- validate-input
- process-extraction
- review-exceptions
- approve-output
- export-results
- generate-report
"""

from pathlib import Path

import click

from audit_framework.audit_log import AuditLog, EventType
from audit_framework.config_manager import ConfigManager
from audit_framework.data_mapper import DataMapper
from audit_framework.exceptions import AuditFrameworkError
from audit_framework.qa_engine import QAEngine
from audit_framework.tag_builder import TagBuilder
from audit_framework.validator import Validator


@click.group()
@click.version_option()
def cli() -> None:
    """Audit Framework - Enterprise audit automation for DataSnipper extraction."""
    pass


@cli.command()
@click.argument("workbook_path", type=click.Path(exists=True))
@click.option("--user", default="system", help="Username for audit trail")
def load_engagement(workbook_path: str, user: str) -> None:
    """Load and validate engagement configuration."""
    try:
        click.echo(f"Loading engagement from: {workbook_path}")

        with ConfigManager(workbook_path) as config_mgr:
            # Load configuration
            config = config_mgr.load_config()
            click.echo(f"✓ Configuration loaded for engagement: {config.engagement_id}")

            # Validate configuration
            is_valid, errors = config_mgr.validate_config()

            if is_valid:
                click.echo("✓ Configuration validation passed")
                click.echo(f"  - Engagement Type: {config.engagement_type.value}")
                click.echo(f"  - Period: {config.period_start_date} to {config.period_end_date}")
                click.echo(f"  - Lead Auditor: {config.lead_auditor}")
                click.echo(f"  - Tags: {len(config.tags)}")
                click.echo(f"  - QA Rules: {len(config.qa_rules)}")
            else:
                click.echo("✗ Configuration validation failed:")
                for error in errors:
                    click.echo(f"  - {error}")
                raise click.Exit(1)

            # Log to audit trail
            audit_log = AuditLog(workbook_path)
            audit_log.log_config_loaded(config.engagement_id, user=user)
            audit_log.save()

    except AuditFrameworkError as e:
        click.echo(f"✗ Error: {e}", err=True)
        raise click.Exit(1)


@cli.command()
@click.argument("workbook_path", type=click.Path(exists=True))
@click.option("--user", default="system", help="Username for audit trail")
@click.option("--output-sheet", default="TAG_ENGINE", help="Sheet to write tags to")
def build_tags(workbook_path: str, user: str, output_sheet: str) -> None:
    """Generate DataSnipper extraction tags."""
    try:
        click.echo(f"Building tags from: {workbook_path}")

        with ConfigManager(workbook_path) as config_mgr:
            config = config_mgr.load_config()
            tags = config.tags

            click.echo(f"Processing {len(tags)} tag definitions...")

            # Build all tags
            generated_tags = TagBuilder.build_all_tags(tags)

            click.echo(f"✓ Generated {len(generated_tags)} tags")

            # Display sample tags
            if generated_tags:
                click.echo("\nSample generated tags:")
                for tag_id, tag_string in list(generated_tags.items())[:3]:
                    click.echo(f"  {tag_id}: {tag_string}")
                if len(generated_tags) > 3:
                    click.echo(f"  ... and {len(generated_tags) - 3} more")

            # Log to audit trail
            audit_log = AuditLog(workbook_path)
            audit_log.log_tags_built(len(generated_tags), user=user)
            audit_log.save()

    except AuditFrameworkError as e:
        click.echo(f"✗ Error: {e}", err=True)
        raise click.Exit(1)


@cli.command()
@click.argument("workbook_path", type=click.Path(exists=True))
@click.option("--input-sheet", default="EXTRACTION_INPUT", help="Sheet with extracted data")
@click.option("--user", default="system", help="Username for audit trail")
def validate_input(workbook_path: str, input_sheet: str, user: str) -> None:
    """Validate extracted data completeness and formats."""
    try:
        click.echo(f"Validating data from: {workbook_path}")

        with ConfigManager(workbook_path) as config_mgr:
            config = config_mgr.load_config()

            # Read extracted data
            from audit_framework.excel_utils import ExcelWorkbookManager

            excel_mgr = ExcelWorkbookManager(workbook_path)
            rows = excel_mgr.read_rows_as_dicts(input_sheet)

            click.echo(f"Validating {len(rows)} rows against schema...")

            # Validate batch
            validation_results = Validator.validate_batch(rows, config.output_schema)

            if validation_results:
                click.echo(f"✗ Validation failed for {len(validation_results)} rows:")
                for row_idx, errors in list(validation_results.items())[:5]:
                    click.echo(f"  Row {row_idx}:")
                    for error in errors[:2]:
                        click.echo(f"    - {error}")
                    if len(errors) > 2:
                        click.echo(f"    ... and {len(errors) - 2} more errors")
            else:
                click.echo("✓ All rows validated successfully")

            # Log to audit trail
            audit_log = AuditLog(workbook_path)
            if validation_results:
                audit_log.log_validation_failed(
                    len(validation_results),
                    user=user,
                    details={"failed_rows": len(validation_results)},
                )
            else:
                audit_log.log_validation_passed(len(rows), user=user)
            audit_log.save()

    except AuditFrameworkError as e:
        click.echo(f"✗ Error: {e}", err=True)
        raise click.Exit(1)


@cli.command()
@click.argument("workbook_path", type=click.Path(exists=True))
@click.option("--input-sheet", default="EXTRACTION_INPUT", help="Sheet with extracted data")
@click.option("--user", default="system", help="Username for audit trail")
def process_extraction(workbook_path: str, input_sheet: str, user: str) -> None:
    """Transform data and apply QA rules."""
    try:
        click.echo(f"Processing extraction from: {workbook_path}")

        with ConfigManager(workbook_path) as config_mgr:
            config = config_mgr.load_config()

            # Read extracted data
            from audit_framework.excel_utils import ExcelWorkbookManager

            excel_mgr = ExcelWorkbookManager(workbook_path)
            rows = excel_mgr.read_rows_as_dicts(input_sheet)

            click.echo(f"Processing {len(rows)} rows...")

            # Transform data
            mapper = DataMapper()
            mapped_rows = mapper.map_batch(rows, config.output_schema, config.tags)

            click.echo(f"✓ Mapped {len(mapped_rows)} rows")

            # Apply QA rules
            qa_engine = QAEngine()
            qa_results = qa_engine.apply_rules_to_batch(config.qa_rules, mapped_rows)

            click.echo(f"✓ Applied {len(config.qa_rules)} QA rules")

            if qa_results:
                total_exceptions = sum(len(excs) for excs in qa_results.values())
                click.echo(f"⚠ Found {total_exceptions} exceptions in {len(qa_results)} rows")

                # Show severity breakdown
                summary = qa_engine.to_dict()["summary"]
                click.echo(f"  - Pending: {summary['pending']}")
                click.echo(f"  - Blocking: {len(qa_engine.get_blocking_exceptions())}")
            else:
                click.echo("✓ No QA exceptions found")

            # Log to audit trail
            audit_log = AuditLog(workbook_path)
            audit_log.log_event(
                EventType.TRANSFORMATION_APPLIED,
                f"Transformed {len(mapped_rows)} rows",
                user=user,
                module="DataMapper",
            )
            if qa_results:
                audit_log.log_validation_failed(
                    total_exceptions,
                    user=user,
                    details={"exception_count": total_exceptions},
                )
            audit_log.save()

    except AuditFrameworkError as e:
        click.echo(f"✗ Error: {e}", err=True)
        raise click.Exit(1)


@cli.command()
@click.argument("workbook_path", type=click.Path(exists=True))
@click.option("--user", default="system", help="Username for audit trail")
def show_status(workbook_path: str, user: str) -> None:
    """Show engagement status and summary."""
    try:
        click.echo(f"Engagement Status: {workbook_path}\n")

        with ConfigManager(workbook_path) as config_mgr:
            config = config_mgr.load_config()

            click.echo("Configuration:")
            click.echo(f"  Engagement ID: {config.engagement_id}")
            click.echo(f"  Type: {config.engagement_type.value}")
            click.echo(f"  Period: {config.period_start_date} to {config.period_end_date}")
            click.echo(f"  Client: {config.client_name}")
            click.echo(f"  Lead Auditor: {config.lead_auditor}")

            click.echo("\nSchema:")
            click.echo(f"  Tags: {len(config.tags)}")
            click.echo(f"  QA Rules: {len(config.qa_rules)}")
            click.echo(f"  Output Fields: {len(config.output_schema)}")

            click.echo("\nApproval Workflow:")
            click.echo(f"  Type: {config.approval_workflow.workflow_type.value}")
            if config.approval_workflow.level1_approver:
                click.echo(f"  Level 1: {config.approval_workflow.level1_approver}")
            if config.approval_workflow.level2_approver:
                click.echo(f"  Level 2: {config.approval_workflow.level2_approver}")

            # Show audit log summary
            audit_log = AuditLog(workbook_path)
            summary = audit_log.get_summary()
            click.echo("\nAudit Log:")
            click.echo(f"  Total Events: {summary['total_events']}")
            if summary["first_event"]:
                click.echo(f"  First Event: {summary['first_event']}")
            if summary["last_event"]:
                click.echo(f"  Last Event: {summary['last_event']}")

    except AuditFrameworkError as e:
        click.echo(f"✗ Error: {e}", err=True)
        raise click.Exit(1)


def main() -> None:
    """Main entry point."""
    cli()


if __name__ == "__main__":
    main()
