"""
TagBuilder — constructs DataSnipper extraction tag strings from TagDefinitions.

Actual DataSnipper tag format (v4.0+, from official documentation)
-------------------------------------------------------------------

DS_SEARCH:
    DS_SEARCH[filename|pageNumber|query]

    - filename    PDF filename or path relative to the workbook location
                  (e.g. "invoice.pdf" or "PDFS\\invoice.pdf")
    - pageNumber  1-based page number to search on
    - query       exact text string DataSnipper will search for in the PDF

    Example: DS_SEARCH[invoice.pdf|1|Invoice Number]

DS_COORDS:
    DS_COORDS[filename|pageNumber|x1|y1|x2|y2]

    - filename    PDF filename or path relative to the workbook location
    - pageNumber  1-based page number
    - x1, y1     top-left corner of the extraction box (pixels from top-left of page)
    - x2, y2     bottom-right corner of the extraction box

    Example: DS_COORDS[invoice.pdf|1|302|81|460|101]

Workbook naming requirement
---------------------------
The Excel workbook filename MUST end in _(ds) for DataSnipper to auto-process
on open.  Example: engagement_(ds).xlsx

File path notes
---------------
- Relative paths are relative to the workbook's saved location on disk.
- Simplest setup for testing: put the PDF in the same folder as the workbook
  and use just the filename (no directory prefix).
- Absolute paths also work: C:\\Users\\name\\Documents\\invoice.pdf
"""

from audit_framework.exceptions import TagBuildError
from audit_framework.models import ExtractionMethod, TagDefinition


class TagBuilder:
    """Constructs DataSnipper tag strings from TagDefinition objects."""

    @staticmethod
    def build_search_tag(tag_def: TagDefinition) -> str:
        """
        Build a DS_SEARCH tag string.

        Format: DS_SEARCH[filename|pageNumber|query]

        Args:
            tag_def: Must have source_document and search_keywords set.

        Returns:
            Formatted DS_SEARCH tag string ready to paste into an Excel cell.

        Raises:
            TagBuildError: If source_document or search_keywords is missing.
        """
        if not tag_def.source_document:
            raise TagBuildError(
                f"Tag {tag_def.tag_id}: source_document (PDF filename) is required for DS_SEARCH"
            )
        if not tag_def.search_keywords:
            raise TagBuildError(
                f"Tag {tag_def.tag_id}: search_keywords (query text) is required for DS_SEARCH"
            )

        return f"DS_SEARCH[{tag_def.source_document}|{tag_def.search_page}|{tag_def.search_keywords}]"

    @staticmethod
    def build_coords_tag(tag_def: TagDefinition) -> str:
        """
        Build a DS_COORDS tag string.

        Format: DS_COORDS[filename|pageNumber|x1|y1|x2|y2]

        x2 and y2 are computed from (coord_x + coord_width) and
        (coord_y + coord_height) respectively, since DataSnipper expects
        two corner points rather than origin + dimensions.

        Args:
            tag_def: Must have source_document set and coord_width/height > 0.

        Returns:
            Formatted DS_COORDS tag string ready to paste into an Excel cell.

        Raises:
            TagBuildError: If source_document is missing or box dimensions are zero.
        """
        if not tag_def.source_document:
            raise TagBuildError(
                f"Tag {tag_def.tag_id}: source_document (PDF filename) is required for DS_COORDS"
            )
        if tag_def.coord_width == 0 or tag_def.coord_height == 0:
            raise TagBuildError(
                f"Tag {tag_def.tag_id}: coord_width and coord_height must be > 0 for DS_COORDS"
            )

        x2 = tag_def.coord_x + tag_def.coord_width
        y2 = tag_def.coord_y + tag_def.coord_height

        return (
            f"DS_COORDS[{tag_def.source_document}|{tag_def.coord_page}"
            f"|{tag_def.coord_x}|{tag_def.coord_y}|{x2}|{y2}]"
        )

    @staticmethod
    def build_hybrid_tag(tag_def: TagDefinition) -> str:
        """DS_HYBRID is not part of the DataSnipper v4.0 format. Use DS_SEARCH or DS_COORDS."""
        raise TagBuildError(
            f"Tag {tag_def.tag_id}: DS_HYBRID is not supported by DataSnipper. "
            "Use DS_SEARCH for text-based extraction or DS_COORDS for coordinate-based extraction."
        )

    @staticmethod
    def build_tag(tag_def: TagDefinition) -> str:
        """
        Dispatch to the correct builder based on tag_def.extraction_method.

        Raises:
            TagBuildError: If building fails or extraction_method is unrecognised.
        """
        try:
            if tag_def.extraction_method == ExtractionMethod.DS_SEARCH:
                return TagBuilder.build_search_tag(tag_def)
            elif tag_def.extraction_method == ExtractionMethod.DS_COORDS:
                return TagBuilder.build_coords_tag(tag_def)
            elif tag_def.extraction_method == ExtractionMethod.HYBRID:
                return TagBuilder.build_hybrid_tag(tag_def)
            else:
                raise TagBuildError(f"Unknown extraction method: {tag_def.extraction_method}")
        except TagBuildError:
            raise
        except Exception as e:
            raise TagBuildError(f"Failed to build tag {tag_def.tag_id}: {e}") from e

    @staticmethod
    def build_all_tags(tag_definitions: list[TagDefinition]) -> dict[str, str]:
        """
        Build all tag strings, accumulating errors rather than short-circuiting.

        Returns:
            Dict mapping tag_id → tag string for all successful tags.

        Raises:
            TagBuildError: Aggregated message listing every tag that failed.
        """
        tags: dict[str, str] = {}
        errors: list[str] = []

        for tag_def in tag_definitions:
            try:
                tags[tag_def.tag_id] = TagBuilder.build_tag(tag_def)
            except TagBuildError as e:
                errors.append(str(e))

        if errors:
            raise TagBuildError("Failed to build tags:\n" + "\n".join(errors))

        return tags

    @staticmethod
    def validate_tag_syntax(tag_string: str) -> tuple[bool, str]:
        """
        Validate that a tag string matches the DataSnipper format.

        Valid formats:
            DS_SEARCH[filename|page|query]
            DS_COORDS[filename|page|x1|y1|x2|y2]

        Returns:
            (True, "")                  if valid.
            (False, error_description)  if invalid.
        """
        if not tag_string or not isinstance(tag_string, str):
            return False, "Tag must be a non-empty string"

        valid_prefixes = ("DS_SEARCH[", "DS_COORDS[")
        if not tag_string.startswith(valid_prefixes):
            return False, "Tag must start with DS_SEARCH[ or DS_COORDS["

        if not tag_string.endswith("]"):
            return False, "Tag must end with ]"

        inner = tag_string[tag_string.index("[") + 1:-1]
        parts = inner.split("|")

        if tag_string.startswith("DS_SEARCH[") and len(parts) < 3:
            return False, "DS_SEARCH tag requires 3 pipe-separated parts: filename|page|query"

        if tag_string.startswith("DS_COORDS[") and len(parts) < 6:
            return False, "DS_COORDS tag requires 6 pipe-separated parts: filename|page|x1|y1|x2|y2"

        return True, ""
