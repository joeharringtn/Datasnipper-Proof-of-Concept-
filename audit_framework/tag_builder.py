"""
TagBuilder — constructs DataSnipper extraction tag strings from TagDefinitions.

Tag format (per TAG_SPEC.md)
-----------------------------

DS_SEARCH:
    DS_SEARCH:{field_id}:{output_field}:(parameters)

    Parameters (pipe-separated):
        start={anchor}          text immediately before the target value
        end={anchor}            text immediately after the target value
        type={type}             data type hint: text|number|currency|date|percentage
        context={scope}         line|paragraph|page|document (optional)
        case_sensitive={yes|no} default no
        fallback={alt_anchor}   alternative start anchor if primary not found
        regex={pattern}         regex alternative to start/end anchors

    Example:
        DS_SEARCH:InvoiceAmount:invoice_total:(start=Total Amount:|end=Tax|type=currency)

DS_COORDS:
    DS_COORDS:{field_id}:{output_field}:(parameters)

    Parameters (all required except file, tolerance):
        file={filename}         source PDF filename
        page={n}                1-based page number
        x={pixels}              horizontal distance from left edge
        y={pixels}              vertical distance from top edge
        width={pixels}          selection box width
        height={pixels}         selection box height
        type={type}             data type hint
        tolerance={pixels}      fuzzy ±pixel tolerance for position variation

    Example:
        DS_COORDS:DepositDate:deposit_date:(page=1|x=300|y=450|width=100|height=20|type=date)

DS_HYBRID:
    DS_HYBRID:{field_id}:{output_field}:(coord_params)|fallback_to:(search_params)

    Tries coordinate extraction first; falls back to keyword search if no result.

    Example:
        DS_HYBRID:Total:receipt_total:(page=1|x=350|y=450|width=80|height=20|type=currency)|fallback_to:(start=Receipt Total:|end=Tax Amount|type=currency)
"""

from audit_framework.exceptions import TagBuildError
from audit_framework.models import ExtractionMethod, TagDefinition


class TagBuilder:
    """Constructs DataSnipper tag strings from TagDefinition objects."""

    @staticmethod
    def build_search_tag(tag_def: TagDefinition) -> str:
        """
        Build a DS_SEARCH tag string.

        Format: DS_SEARCH:{tag_id}:{field_name}:(parameters)

        search_keywords maps to the start= anchor when no explicit start_anchor
        is provided — DataSnipper uses it as the text to search for.

        Args:
            tag_def: Must have at least one of start_anchor, end_anchor,
                     or search_keywords set.

        Returns:
            Formatted DS_SEARCH tag string.

        Raises:
            TagBuildError: If no search criterion is provided.
        """
        if not tag_def.start_anchor and not tag_def.end_anchor and not tag_def.search_keywords:
            raise TagBuildError(
                f"Tag {tag_def.tag_id}: DS_SEARCH requires at least one of "
                "start_anchor, end_anchor, or search_keywords"
            )

        params: list[str] = []

        if tag_def.start_anchor:
            params.append(f"start={tag_def.start_anchor}")
        elif tag_def.search_keywords:
            # Treat keywords as the start anchor when no explicit anchor given
            params.append(f"start={tag_def.search_keywords}")

        if tag_def.end_anchor:
            params.append(f"end={tag_def.end_anchor}")

        params.append(f"type={tag_def.field_type.value.lower()}")

        if tag_def.fallback_keywords:
            params.append(f"fallback={tag_def.fallback_keywords}")

        return f"DS_SEARCH:{tag_def.tag_id}:{tag_def.field_name}:({'|'.join(params)})"

    @staticmethod
    def build_coords_tag(tag_def: TagDefinition) -> str:
        """
        Build a DS_COORDS tag string.

        Format: DS_COORDS:{tag_id}:{field_name}:(parameters)

        coord_page must be >= 1; DataSnipper pages are 1-based.

        Args:
            tag_def: Must have source_document set and coord_page >= 1.

        Returns:
            Formatted DS_COORDS tag string.

        Raises:
            TagBuildError: If source_document is missing or coord_page is 0.
        """
        if not tag_def.source_document:
            raise TagBuildError(f"Tag {tag_def.tag_id}: source_document required for DS_COORDS")

        if tag_def.coord_page < 1:
            raise TagBuildError(
                f"Tag {tag_def.tag_id}: coord_page must be >= 1 (got {tag_def.coord_page})"
            )

        params: list[str] = [
            f"file={tag_def.source_document}",
            f"page={tag_def.coord_page}",
            f"x={tag_def.coord_x}",
            f"y={tag_def.coord_y}",
            f"width={tag_def.coord_width}",
            f"height={tag_def.coord_height}",
            f"type={tag_def.field_type.value.lower()}",
        ]

        if tag_def.tolerance > 0:
            params.append(f"tolerance={tag_def.tolerance}")

        return f"DS_COORDS:{tag_def.tag_id}:{tag_def.field_name}:({'|'.join(params)})"

    @staticmethod
    def build_hybrid_tag(tag_def: TagDefinition) -> str:
        """
        Build a DS_HYBRID tag string.

        Format: DS_HYBRID:{tag_id}:{field_name}:(coord_params)|fallback_to:(search_params)

        DataSnipper tries coordinate extraction first; falls back to keyword
        search if the coordinate region yields no value.

        Args:
            tag_def: Must have source_document + valid coords AND at least one
                     search criterion (start_anchor or search_keywords).

        Returns:
            Formatted DS_HYBRID tag string.

        Raises:
            TagBuildError: If either strategy's required fields are missing.
        """
        if not tag_def.source_document:
            raise TagBuildError(
                f"Tag {tag_def.tag_id}: source_document required for HYBRID (coords strategy)"
            )
        if not tag_def.start_anchor and not tag_def.search_keywords:
            raise TagBuildError(
                f"Tag {tag_def.tag_id}: start_anchor or search_keywords required for HYBRID (search fallback)"
            )

        coord_params = [
            f"file={tag_def.source_document}",
            f"page={tag_def.coord_page}",
            f"x={tag_def.coord_x}",
            f"y={tag_def.coord_y}",
            f"width={tag_def.coord_width}",
            f"height={tag_def.coord_height}",
            f"type={tag_def.field_type.value.lower()}",
        ]
        if tag_def.tolerance > 0:
            coord_params.append(f"tolerance={tag_def.tolerance}")

        search_params: list[str] = []
        if tag_def.start_anchor:
            search_params.append(f"start={tag_def.start_anchor}")
        elif tag_def.search_keywords:
            search_params.append(f"start={tag_def.search_keywords}")
        if tag_def.end_anchor:
            search_params.append(f"end={tag_def.end_anchor}")
        search_params.append(f"type={tag_def.field_type.value.lower()}")

        coord_str = "|".join(coord_params)
        search_str = "|".join(search_params)
        return f"DS_HYBRID:{tag_def.tag_id}:{tag_def.field_name}:({coord_str})|fallback_to:({search_str})"

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
        Validate that a tag string has the expected structural format.

        Accepts DS_SEARCH, DS_COORDS, and DS_HYBRID tag types.

        Returns:
            (True, "")                  if valid.
            (False, error_description)  if invalid.
        """
        if not tag_string or not isinstance(tag_string, str):
            return False, "Tag must be a non-empty string"

        valid_prefixes = ("DS_SEARCH:", "DS_COORDS:", "DS_HYBRID:")
        if not tag_string.startswith(valid_prefixes):
            return False, f"Tag must start with one of: {', '.join(valid_prefixes)}"

        # All types use TYPE:field_id:output_field:(params) — need >= 3 colons
        if tag_string.count(":") < 3:
            return False, "Tag must have at least 4 colon-separated parts"

        if "(" not in tag_string or ")" not in tag_string:
            return False, "Tag parameters must be enclosed in parentheses"

        return True, ""
