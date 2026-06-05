"""
TagBuilder - Generate DataSnipper extraction tags.
"""

from audit_framework.exceptions import TagBuildError
from audit_framework.models import ExtractionMethod, TagDefinition


class TagBuilder:
    """Constructs DataSnipper tags for extraction."""

    @staticmethod
    def build_search_tag(tag_def: TagDefinition) -> str:
        """Build DS_SEARCH tag for keyword-based extraction."""
        if not tag_def.search_keywords:
            raise TagBuildError(f"Tag {tag_def.tag_id}: search_keywords required for DS_SEARCH")

        field_name = tag_def.field_name
        keywords = tag_def.search_keywords

        start_anchor = f"start={tag_def.start_anchor}" if tag_def.start_anchor else ""
        end_anchor = f"end={tag_def.end_anchor}" if tag_def.end_anchor else ""
        type_spec = f"type={tag_def.field_type.value.lower()}"

        anchor_parts = [p for p in [start_anchor, end_anchor, type_spec] if p]
        anchor_spec = "|".join(anchor_parts)

        tag = f"DS_SEARCH:{field_name}:{keywords}:({anchor_spec})"
        return tag

    @staticmethod
    def build_coords_tag(tag_def: TagDefinition) -> str:
        """Build DS_COORDS tag for coordinate-based extraction."""
        if not tag_def.source_document:
            raise TagBuildError(f"Tag {tag_def.tag_id}: source_document required for DS_COORDS")

        field_name = tag_def.field_name
        doc_id = tag_def.source_document

        coords = (
            f"file={doc_id}|page={tag_def.coord_page}|"
            f"x={tag_def.coord_x}|y={tag_def.coord_y}|"
            f"width={tag_def.coord_width}|height={tag_def.coord_height}|"
            f"type={tag_def.field_type.value.lower()}"
        )

        if tag_def.tolerance > 0:
            coords += f"|tolerance={tag_def.tolerance}"

        tag = f"DS_COORDS:{field_name}:{coords}"
        return tag

    @staticmethod
    def build_hybrid_tag(tag_def: TagDefinition) -> str:
        """Build hybrid tag combining search and coordinate methods."""
        if not tag_def.search_keywords or not tag_def.source_document:
            raise TagBuildError(
                f"Tag {tag_def.tag_id}: both search_keywords and source_document required for HYBRID"
            )

        search_tag = TagBuilder.build_search_tag(tag_def)
        coords_tag = TagBuilder.build_coords_tag(tag_def)

        tag = f"{search_tag}|{coords_tag}"
        return tag

    @staticmethod
    def build_tag(tag_def: TagDefinition) -> str:
        """Build tag based on extraction method."""
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
        """Build all tags from definitions."""
        tags = {}
        errors = []

        for tag_def in tag_definitions:
            try:
                tag = TagBuilder.build_tag(tag_def)
                tags[tag_def.tag_id] = tag
            except TagBuildError as e:
                errors.append(str(e))

        if errors:
            raise TagBuildError(f"Failed to build tags:\n" + "\n".join(errors))

        return tags

    @staticmethod
    def validate_tag_syntax(tag_string: str) -> tuple[bool, str]:
        """Validate tag syntax."""
        if not tag_string or not isinstance(tag_string, str):
            return False, "Tag must be a non-empty string"

        if not tag_string.startswith(("DS_SEARCH:", "DS_COORDS:")):
            return False, "Tag must start with DS_SEARCH: or DS_COORDS:"

        if tag_string.count(":") < 2:
            return False, "Tag must have at least 3 colon-separated parts"

        return True, ""
