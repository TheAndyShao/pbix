"""A module providing tools for manipulating visual elements within a Power BI Report."""

import json
import re
from typing import Any, Union

from jsonpath_ng.ext import parse

JsonType = dict[str, "JsonType"]


class GenericVisual:
    """A base class to represent a generic visual object."""

    def __init__(self, layout: dict[str, Any]) -> None:
        self.layout: dict[str, Any] = layout
        self.config: dict[str, Any] = json.loads(self.layout.get("config"))  # type: ignore
        self.type: Union[str, None] = self._return_visual_type()
        non_data_visuals = ["image", "textbox", "shape", "actionButton", None]
        self.is_data_visual: bool = self.type not in non_data_visuals
        self.updated: int = 0

    def _return_visual_title(self) -> Union[str, None]:
        """Return title of visual."""
        path = parse(
            "$.singleVisual.vcObjects.title[0].properties.text.expr.Literal.Value"
        )
        node = path.find(self.config)
        return node[0].value if node else None

    def _return_visual_type(self) -> Union[str, None]:
        """Return type of visual."""
        path = parse("$.singleVisual.visualType")
        node = path.find(self.config)
        return node[0].value if node else None


class DataVisual(GenericVisual):
    """A class representing visuals that depend on a data model."""

    def __init__(self, layout: dict[str, Any]) -> None:
        super().__init__(layout)
        self.title: Union[str, None] = self._return_visual_title()
        self.filters: Filters = Filters(json.loads(self.layout.get("filters")))  # type: ignore
        self.query: Union[Query, None] = (
            Query(json.loads(self.layout.get("query")))  # type: ignore
            if "query" in self.layout
            else None
        )
        self.data_transforms: Union[DataTransforms, None] = (
            DataTransforms(json.loads(self.layout.get("dataTransforms")))  # type: ignore
            if "dataTransforms" in self.layout
            else None
        )
        self.config: Config = Config(self.config)
        self.visual_options = {
            "config": self.config.config,
            "filters": self.filters.filters,
            "query": self.query.visual_query if self.query else None,
            "dataTransforms": self.data_transforms.data_transforms
            if self.data_transforms
            else None,
        }

    def find_field(self, table_field: str) -> bool:
        """Find if a field is used in the visual"""
        table_measure_path = parse(
            f"$..@[?(@.*=='{table_field}')].[queryRef, Name, queryName]"
        )
        return any(table_measure_path.find(v) for v in self.visual_options.values())

    def update_fields(self, old: str, new: str) -> None:
        """Searches for relevant keys for fields and updates their value pairs."""
        table_old, field_old = old.split(".")
        table_new, field_new = new.split(".")

        if self.find_field(old):
            self.config.update_fields(
                old, new, table_old, table_new, field_old, field_new
            )
            if self.data_transforms:
                self.data_transforms.update_fields(
                    old, new, table_old, table_new, field_old, field_new
                )
            if self.query:
                self.query.update_fields(
                    old, new, table_old, table_new, field_old, field_new
                )
            self.filters.update_fields(
                old, new, table_old, table_new, field_old, field_new
            )
            for option, value in self.visual_options.items():
                if value:
                    self.layout[option] = json.dumps(value)
            self.updated = 1
            print(f"Updated: {self.title}")


class Config:
    """A class representing the config settings of a visual."""

    def __init__(self, config: dict[str, Any]) -> None:
        self.config = config
        self.single_visual = self.config["singleVisual"]
        self.objects = self.single_visual.get("objects")
        self.datapoint = self.objects.get("dataPoint")
        self.prototypequery: SemanticQuery = SemanticQuery(
            self.single_visual["prototypeQuery"]
        )

    def update_fields(
        self,
        table_field_old: str,
        table_field_new: str,
        table_old: str,
        table_new: str,
        field_old: str,
        field_new: str,
    ) -> None:
        """Replace fields in all relevant config settings."""
        self.prototypequery.update_fields(
            table_field_old, table_field_new, table_old, table_new, field_old, field_new
        )
        self._prune_column_properties(table_field_old, table_field_new)
        self._update_object_datapoints(table_old, table_new, field_old, field_new)
        self._update_singlevisual_table_fields(table_field_old, table_field_new)

        # Table field measures act like ids so update these last
        self._update_projections_table_fields(table_field_old, table_field_new)

    def _update_projections_table_fields(
        self, table_field_old: str, table_field_new: str
    ) -> None:
        """Updating projections."""
        path = parse(f"$.projections.*[?(@.queryRef=='{table_field_old}')].queryRef")
        path.update(self.single_visual, table_field_new)

    def _update_singlevisual_table_fields(
        self, table_field_old: str, table_field_new: str
    ) -> None:
        """Update single visual."""
        path = parse(
            f"$.objects.*[?(@.selector.metadata=='{table_field_old}')].selector.metadata"
        )
        path.update(self.single_visual, table_field_new)

    def _update_object_datapoints(
        self, table_old: str, table_new: str, field_old: str, field_new: str
    ) -> None:
        root_path = "$..@[?(@.Property=='{field_old}')]"
        path = parse(root_path.format(field_old=field_old))
        nodes = path.find(self.datapoint)
        for node in nodes:
            if node.value.get("Expression").get("SourceRef").get("Entity") == table_old:
                field_path = parse(root_path.format(field_old=field_old) + ".Property")
                field_path.update(node, field_new)
                table_path = parse(
                    root_path.format(field_old=field_new)
                    + ".Expression.SourceRef.Entity"
                )
                table_path.update(node, table_new)

    def _prune_column_properties(
        self, table_field_old: str, table_field_new: str
    ) -> None:
        """Update column properties if necessary."""
        column_properties = self.single_visual.get("columnProperties", {})
        if table_field_old in column_properties:
            column_properties[table_field_new] = column_properties.pop(table_field_old)


class Query:
    """A class representing the query settings of a visual"""

    def __init__(self, visual_query) -> None:
        self.visual_query = visual_query
        self.commands = self.visual_query.get("Commands")

    def update_fields(
        self,
        table_field_old: str,
        table_field_new: str,
        table_old: str,
        table_new: str,
        field_old: str,
        field_new: str,
    ) -> None:
        """Replace field in all relevant query settings."""
        for command in self.commands:
            query = SemanticQuery(command["SemanticQueryDataShapeCommand"]["Query"])
            query.update_fields(
                table_field_old,
                table_field_new,
                table_old,
                table_new,
                field_old,
                field_new,
            )


class DataTransforms:
    """A class representing the datatransforms settings of a Visual."""

    def __init__(self, data_transforms: dict[str, Any]) -> None:
        self.data_transforms: dict[str, Any] = data_transforms
        self.metadata = self.data_transforms.get("queryMetadata")
        self.objects = self.data_transforms.get("objects")
        self.datapoint = self.objects.get("dataPoint")

    def update_fields(
        self,
        table_field_old: str,
        table_field_new: str,
        table_old: str,
        table_new: str,
        field_old: str,
        field_new: str,
    ) -> None:
        """Replace fields in all relevant datatransforms settings."""
        self._update_objects_metadata(table_field_old, table_field_new)
        self._update_object_datapoints(table_old, table_new, field_old, field_new)
        self._update_selects_tables(table_field_old, table_new)
        self._update_selects_fields(table_field_old, field_new)
        self._update_selects_display_names(table_field_old, field_old, field_new)
        self._update_filters_metadata_tables(field_old, table_new)
        self._update_filters_metadata_fields(field_old, field_new)

        # Table field measures act like ids so update these last
        self._update_selects_table_fields(table_field_old, table_field_new)
        self._update_query_metadata_table_fields(table_field_old, table_field_new)

    def _update_selects_tables(self, table_field_old: str, table_new: str) -> None:
        """Update table references in selects."""
        path = parse(
            f"$.selects[?(@.queryName=='{table_field_old}')].expr.*.Expression.SourceRef.Entity"
        )
        path.update(self.data_transforms, table_new)

    def _update_selects_fields(self, table_field_old: str, field: str) -> None:
        """Update field in selects."""
        path = parse(f"$.selects[?(@.queryName=='{table_field_old}')].expr.*.Property")
        path.update(self.data_transforms, field)

    def _update_selects_display_names(
        self, table_field_old: str, field_old: str, field_new: str
    ) -> None:
        """Updates display name if no custom display name is found."""
        path = parse(f"$.selects[?(@.queryName=='{table_field_old}')].displayName")
        node = path.find(self.data_transforms)
        if node[0].value == field_old:
            path.update(self.data_transforms, field_new)

    def _update_selects_table_fields(
        self, table_field_old: str, table_field_new: str
    ) -> None:
        """Update table.field in selects."""
        path = parse(f"$.selects[?(@.queryName=='{table_field_old}')].queryName")
        path.update(self.data_transforms, table_field_new)

    def _update_objects_metadata(
        self, table_field_old: str, table_field_new: str
    ) -> None:
        """Update table references in metadata."""
        path = parse(
            f"$.objects.*[?(@.selector.metadata=='{table_field_old}')].selector.metadata"
        )
        path.update(self.data_transforms, table_field_new)

    def _update_object_datapoints(
        self, table_old: str, table_new: str, field_old: str, field_new: str
    ) -> None:
        root_path = "$..@[?(@.Property=='{field_old}')]"
        path = parse(root_path.format(field_old=field_old))
        nodes = path.find(self.datapoint)
        for node in nodes:
            if node.value.get("Expression").get("SourceRef").get("Entity") == table_old:
                field_path = parse(root_path.format(field_old=field_old) + ".Property")
                field_path.update(node, field_new)
                table_path = parse(
                    root_path.format(field_old=field_new)
                    + ".Expression.SourceRef.Entity"
                )
                table_path.update(node, table_new)

    def _update_query_metadata_table_fields(
        self, table_field_old: str, table_field_new: str
    ) -> None:
        """Update table.field in query metadata."""
        path = parse(f"$.queryMetadata.Select[?(@.Name=='{table_field_old}')].Name")
        path.update(self.data_transforms, table_field_new)

    def _update_filters_metadata_tables(self, field_old: str, table_new: str) -> None:
        path = parse(
            f"$.Filters[?(@.expression.*.Property=='{field_old}')].expression.*.Expression.SourceRef.Entity"
        )
        path.update(self.metadata, table_new)

    def _update_filters_metadata_fields(self, field_old: str, field_new: str) -> None:
        path = parse(
            f"$.Filters[?(@.expression.*.Property=='{field_old}')].expression.*.Property"
        )
        path.update(self.metadata, field_new)


class Filters:
    """A class representing the filter object of a visual."""

    def __init__(self, filters: JsonType) -> None:
        self.filters = filters

    def update_fields(
        self,
        table_field_old: str,
        table_field_new: str,
        table_old: str,
        table_new: str,
        field_old: str,
        field_new: str,
    ) -> None:
        """Finds usage of an existing field and replaces it with a new specified field."""
        self._update_semantic_queries(
            table_field_old, table_field_new, table_old, table_new, field_old, field_new
        )
        self._update_tables(field_old, table_new)
        self._update_fields(field_old, field_new)

    def _update_tables(self, field_old: str, table_new: str) -> None:
        path = parse(
            f"$[?(@.expression.*.Property=='{field_old}')].expression.*.Expression.SourceRef.Entity"
        )
        path.update(self.filters, table_new)

    def _update_fields(self, field_old: str, field_new: str) -> None:
        path = parse(
            f"$[?(@.expression.*.Property=='{field_old}')].expression.*.Property"
        )
        path.update(self.filters, field_new)

    def _update_semantic_queries(
        self,
        table_field_old: str,
        table_field_new: str,
        table_old: str,
        table_new: str,
        field_old: str,
        field_new: str,
    ) -> None:
        path = parse(f"$[?(@.expression.*.Property=='{field_old}')].filter")
        for flt in path.find(self.filters):
            flt = SemanticQuery(flt.value)
            flt.update_fields(
                table_field_old,
                table_field_new,
                table_old,
                table_new,
                field_old,
                field_new,
            )


class SemanticQuery:
    """A class representing a query object used by visuals to query the associated data model."""

    def __init__(self, query) -> None:
        self.query = query
        self.frm = self.query.get("From")
        self.select = self.query.get("Select")
        self.where = self.query.get("Where")
        self.order_by = self.query.get("OrderBy")

    def update_fields(
        self,
        table_field_old: str,
        table_field_new: str,
        table_old: str,
        table_new: str,
        field_old: str,
        field_new: str,
    ) -> None:
        """Finds usage of an existing field and replaces it with a new specified field."""
        self._prune_from_tables(table_field_old, table_old)
        table_alias_new = self._return_from_table_alias(table_new)
        if not table_alias_new:
            table_alias_new = self._generate_table_aliases(table_new)
            self._add_from_table(table_new, table_alias_new)
        self._update_select_table_aliases(table_field_old, table_alias_new)
        self._update_select_fields(table_field_old, field_new)
        self._update_orderby_table_aliases(field_old, table_alias_new)
        self._update_orderby_fields(field_old, field_new)
        if self.where:
            self._update_where_settings(field_old, field_new, table_alias_new)

        # Table field measures act like ids so update these last
        self._update_select_table_fields(table_field_old, table_field_new)

    # 'From' setting methods
    def _return_from_table_alias(self, table: str) -> Union[str, None]:
        """Finds if a table is present as a source in the prototypequery object."""
        path = parse(f"$[?(@.Entity=='{table}')].Name")
        node = path.find(self.frm)
        if node:
            return node[0].value
        return

    def _return_from_tables(self) -> list[str]:
        """Returns all the currently used table name aliases in the prototypequery."""
        path = parse("$.[*].Name")
        nodes = path.find(self.frm)
        return [name.value for name in nodes]

    def _add_from_table(self, table: str, name: str) -> None:
        """Adds a new table to the prototypequery."""
        entry = {"Name": name, "Entity": table, "Type": 0}
        self.frm.append(entry)

    def _prune_from_tables(self, table_field_old: str, table_old: str) -> None:
        table_alias_old = self._return_from_table_alias(table_old)
        selects = self._return_select_tables(table_field_old)
        wheres = self._return_where_tables(table_alias_old)
        if table_alias_old not in selects and table_alias_old not in wheres:
            for table in self.frm:
                if table["Name"] == table_alias_old:
                    self.frm.remove(table)

    def _generate_table_aliases(self, table_new: str) -> str:
        """Returns a new table name alias for additions to the prototypequery."""
        # Power BI reassigns aliases when visual is updated so don't need to replicate exactly
        alias = table_new[:1].lower()
        regex = re.compile("[^0-9]")
        names = self._return_from_tables()
        names = [int(regex.sub("0", name)) for name in names if name[:1] == alias]
        if names:
            return alias + str(max(names) + 1)
        return alias

    # 'Select' setting methods
    def _return_select_tables(self, table_field_old: str) -> list[str]:
        path = parse(f"$[?(@.Name!='{table_field_old}')].*.Expression.SourceRef.Source")
        nodes = path.find(self.select)
        return [node.value for node in nodes]

    def _update_select_table_aliases(self, table_field: str, name: str) -> None:
        """Updates prototypequery table name alias."""
        path = parse(f"$[?(@.Name=='{table_field}')].*.Expression.SourceRef.Source")
        path.update(self.select, name)

    def _update_select_fields(self, table_field: str, field: str) -> None:
        """Updating prototypequery fields."""
        path = parse(f"$[?(@.Name=='{table_field}')].*.Property")
        path.update(self.select, field)

    def _update_select_table_fields(
        self, table_field_old: str, table_field_new: str
    ) -> None:
        """Updating prototypequery table fields."""
        path = parse(f"$[?(@.Name=='{table_field_old}')].Name")
        path.update(self.select, table_field_new)

    # 'OrderBy' setting methods
    def _update_orderby_table_aliases(self, field_old: str, alias_new: str) -> None:
        path = parse(
            f"$[?(@.Expression.*.Property=='{field_old}')].Expression.*.Expression.SourceRef.Source"
        )
        path.update(self.order_by, alias_new)

    def _update_orderby_fields(self, field_old: str, field_new: str) -> None:
        path = parse(
            f"$[?(@.Expression.*.Property=='{field_old}')].Expression.*.Property"
        )
        path.update(self.order_by, field_new)

    # 'Where' setting methods
    def _return_where_tables(self, alias: str) -> list[str]:
        path = parse(f"$[?(@..Source!='{alias}')]..Source")
        nodes = path.find(self.where)
        return [node.value for node in nodes]

    def _update_where_settings(self, field_old: str, field_new: str, name: str) -> None:
        for node in self.where:
            for _, setting in node.items():
                self._update_where_table_aliases(field_old, name, setting)
                self._update_where_fields(field_old, field_new, setting)

    def _update_where_table_aliases(
        self, field_old: str, name: str, condition: str
    ) -> None:
        path = parse("$..Property")
        if path.find(condition)[0].value == field_old:
            path = parse("$..Source")
            path.update(condition, name)

    def _update_where_fields(
        self, field_old: str, field_new: str, condition: str
    ) -> None:
        path = parse("$..Property")
        if path.find(condition)[0].value == field_old:
            path.update(condition, field_new)


class Slicer(DataVisual):
    """A class representing a slicer visual."""

    def unselect_all_items(self) -> None:
        """Unselects all slicer members where no default selection is defined"""
        slicer_path = parse(
            "$.singleVisual.objects.data[*].properties.isInvertedSelectionMode.`parent`"
        )
        selection_path = parse("$.singleVisual.objects.general[*].properties.filter")
        slicer = slicer_path.find(self.config.config)
        selection = selection_path.find(self.config.config)
        if slicer and not selection:
            slicer[0].value.pop("isInvertedSelectionMode")
            self.layout["config"] = json.dumps(self.config.config)
            self.updated = 1
            print(f"Updated: {self.title if self.title else 'Untitled Slicer'}")
