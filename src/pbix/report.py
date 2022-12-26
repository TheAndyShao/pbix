"""A module providing tools for manipulating thin Power BI reports"""

import json
import os
import re
import zipfile as zf

from jsonpath_ng.ext import parse


class Report:
    """A class to represent a thin Power BI report."""

    def __init__(self, filepath: str) -> None:
        self.filepath: str = filepath
        self.filename: str = os.path.basename(filepath)
        self.layout: str = self._read_layout(filepath)
        self.layout_full_json: str = self._read_full_json_layout(filepath)
        self.field_set: set[str] = self.get_all_fields()
        # self.pages = self.layout.get('sections')
        self.updated: int = 0

    def get_all_fields(self) -> None:
        """Get a list of used fields in the pbix file."""
        filter_path = parse(
            "$.sections[*].visualContainers[*].filters[*].expression.Measure.Property"
        )
        field_path = parse(
            "$.sections[*].visualContainers[*].config.singleVisual.projections[*].*.[*].queryRef"
        )

        filter_set = set(
            [match.value for match in filter_path.find(self.layout_full_json)]
        )
        measure_set = set(
            [match.value for match in field_path.find(self.layout_full_json)]
        )
        return filter_set.union(measure_set)

    def find_instances(self, fields: list[str]) -> dict[str, bool]:
        """Compare input fields with the fields used in the pbix file."""
        matches = {}
        for field in fields:
            if "." in field:
                field_set = self.field_set
            else:
                field_set = [measure.split(".")[-1] for measure in self.field_set]

            if field in field_set:
                matches[field] = True
        return matches

    def update_fields(self, old: str, new: str) -> None:
        """Iterates through pages and visuals in a pbix and replaces specified measure/column."""
        print(f"Updating: {self.filename}")
        for i, j, visual in self._generic_visuals_generator():
            if visual.is_data_visual:
                visual = DataVisual(visual)
                visual.update_fields(old, new)
                self._update_visual_layout(i, j, visual.layout)
                self.updated += visual.updated
        # TODO: Currently the below causes report level slicers to break.
        # for page in self.pages:
        #     page = ReportPage(page)
        #     page.update_fields(old, new)
        #     self.updated += page.updated
        if self.updated == 0:
            print("No fields to update")

    def update_slicers(self) -> None:
        """Iterates through pages and genric visuals and updates slicers."""
        for i, j, visual in self._generic_visuals_generator():
            if visual.type == "slicer":
                slicer = Slicer(visual)
                slicer.unselect_all_items()
                self._update_visual_layout(i, j, slicer.layout)
                self.updated += slicer.updated
        if self.updated == 0:
            print("No slicers to update")

    def write_file(self) -> None:
        """Writes the pbix json to file."""
        base, ext = os.path.splitext(self.filename)
        temp_filepath = os.path.join(f"{base} Temp{ext}")

        with zf.ZipFile(self.filepath, "r") as original_zip:
            with zf.ZipFile(temp_filepath, "w", compression=zf.ZIP_DEFLATED) as new_zip:
                for file in original_zip.namelist():
                    if file == "Report/Layout":
                        new_zip.writestr(
                            file, json.dumps(self.layout).encode("utf-16-le")
                        )
                    elif file != "SecurityBindings":
                        new_zip.writestr(file, original_zip.read(file))

        os.remove(self.filepath)
        os.rename(temp_filepath, self.filepath)

    def _read_layout(self, filepath: str) -> str:
        """Return a JSON object of the layout file within the PBIX file."""
        with zf.ZipFile(filepath, "r") as zip_file:
            string = zip_file.read("Report/Layout").decode("utf-16")
            return json.loads(string)

    def _read_full_json_layout(self, filepath: str) -> str:
        """Return a fully JSONified object of the layout file within the PBIX file."""
        with zf.ZipFile(filepath, "r") as zip_file:
            string = zip_file.read("Report/Layout").decode("utf-16")
            replacements = {
                chr(0): "",
                chr(28): "",
                chr(29): "",
                chr(25): "",
                '"[': "[",
                ']"': "]",
                '"{': "{",
                '}"': "}",
                "\\\\": "\\",
                '\\"': '"',
            }
            for original, replacement in replacements.items():
                string = string.replace(original, replacement)
            return json.loads(string)

    def write_json_layout(self) -> None:
        """Write the cleaned JSON object to file."""
        with open("layout.json", "w", encoding="utf-16") as outfile:
            json.dump(self.layout_full_json, outfile)

    def _generic_visuals_generator(self) -> None:
        """Generator for iterating through all visuals in a file."""
        for i, page in enumerate(self.layout["sections"]):
            visuals = page["visualContainers"]
            for j, visual in enumerate(visuals):
                yield i, j, GenericVisual(visual)

    def _update_visual_layout(self, page: int, visual: int, layout: str) -> None:
        """Updates visual layout with new definition."""
        self.layout["sections"][page]["visualContainers"][visual] = layout


class ReportPage:
    def __init__(self, page) -> None:
        self.page = page
        self.filters = VisualFilters(json.loads(self.page.get("filters")))
        self.updated = 0

    def update_fields(self, table_field_old, table_field_new):
        # TODO: Currently this causes report level slicers to break,
        # even when the slicers conditions are cleared.
        # Return to in the future to enable report level slicer updating.
        table_old, field_old = table_field_old.split(".")
        table_new, field_new = table_field_new.split(".")
        self.filters.update_fields(
            table_field_old, table_field_new, table_old, table_new, field_old, field_new
        )

        self.page["filters"] = self.filters.filters


class GenericVisual:
    """A base class to represent a generic visual object."""

    def __init__(self, layout: str) -> None:
        self.layout: str = layout
        self.config: str = json.loads(self.layout.get("config"))
        self.title: str or None = None
        self.type: str or None = self._return_visual_type()
        non_data_visuals = ["image", "textbox", "shape", "actionButton", None]
        self.is_data_visual: bool = self.type not in non_data_visuals
        self.updated: int = 0

    def _return_visual_title(self) -> str or None:
        """Return title of visual."""
        title_path = parse(
            "$.singleVisual.vcObjects.title[0].properties.text.expr.Literal.Value"
        )
        title = title_path.find(self.config)
        return title[0].value if title else None

    def _return_visual_type(self) -> str or None:
        """Return type of visual."""
        typ_path = parse("$.singleVisual.visualType")
        typ = typ_path.find(self.config)
        return typ[0].value if typ else None


class DataVisual(GenericVisual):
    """A class representing visuals that depend on a data model."""

    field_path = "$..@[?(@.*=='{field}')].[Property, displayName, Restatement]"
    table_field_path = "$..@[?(@.*=='{table_field}')].[queryRef, Name, queryName]"

    def __init__(self, Visual: GenericVisual) -> None:
        super().__init__(Visual.layout)
        self.title: str = self._return_visual_title()
        self.filters: str = VisualFilters(json.loads(self.layout.get("filters")))
        self.query: str or None = (
            VisualQuery(json.loads(self.layout.get("query")))
            if "query" in self.layout
            else None
        )
        self.data_transforms: str or None = (
            VisualDataTransforms(json.loads(self.layout.get("dataTransforms")))
            if "dataTransforms" in self.layout
            else None
        )
        self.config = VisualConfig(self.config)
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
            self.table_field_path.format(table_field=table_field)
        )
        return any(table_measure_path.find(v) for v in self.visual_options.values())

    def update_fields(self, old: str, new: str) -> None:
        """Searches for relevant keys for fields and updates their value pairs."""
        old_table, old_measure = old.split(".")
        new_table, new_measure = new.split(".")

        field_path = parse(self.field_path.format(field=old_measure))
        table_field_path = parse(self.table_field_path.format(table_field=old))

        if self.find_field(old):
            self.config.update_fields(
                old, new, old_table, new_table, old_measure, new_measure
            )
            if self.data_transforms:
                self.data_transforms.update_fields(
                    old, new, new_table, old_measure, new_measure
                )
            if self.query:
                self.query.update_fields(
                    old, new, old_table, new_table, old_measure, new_measure
                )
            self.filters.update_fields(
                old, new, old_table, new_table, old_measure, new_measure
            )
            for option, value in self.visual_options.items():
                # field_path.update(value, new_measure)
                # table_field_path.update(value, new)
                if value:
                    self.layout[option] = json.dumps(value)
            self.updated = 1
            print(f"Updated: {self.title}")


class VisualConfig:
    """A class representing the config settings of a visual."""

    def __init__(self, config: str) -> None:
        self.config = config
        self.single_visual = self.config["singleVisual"]
        self.prototypequery = GenericVisualQuery(self.single_visual["prototypeQuery"])

    def update_fields(
        self,
        table_field_old,
        table_field_new,
        table_old,
        table_new,
        field_old,
        field_new,
    ):
        """Replace fields in all relevant config settings."""
        self.prototypequery.update_fields(
            table_field_old, table_field_new, table_old, table_new, field_old, field_new
        )
        self._update_column_properties(table_field_old, table_field_new)
        self._update_singlevisual(table_field_old, table_field_new)

        # Table field measures act like ids so update these last
        self._update_projections(table_field_old, table_field_new)

    def _update_projections(self, table_field_old, table_field_new):
        """Updating projections."""
        path = parse(f"$.projections.*[?(@.queryRef=='{table_field_old}')].queryRef")
        path.update(self.single_visual, table_field_new)

    def _update_column_properties(self, table_field_old, table_field_new):
        """Update column properties if necessary."""
        column_properties = self.single_visual.get("columnProperties", [])
        if table_field_old in column_properties:
            column_properties[table_field_new] = column_properties.pop(table_field_old)

    def _update_singlevisual(self, table_field_old, table_field_new):
        """Update single visual."""
        path = parse(
            f"$.objects.*[?(@.selector.metadata=='{table_field_old}')].selector.metadata"
        )
        path.update(self.single_visual, table_field_new)


class GenericVisualQuery:
    def __init__(self, query) -> None:
        self.query = query
        self.frm = self.query.get("From")
        self.select = self.query.get("Select")
        self.where = self.query.get("Where")
        self.order_by = self.query.get("OrderBy")

    def update_fields(
        self,
        table_field_old,
        table_field_new,
        table_old,
        table_new,
        field_old,
        field_new,
    ):
        self._cleanup_tables(table_field_old, table_old)
        table_alias_new = self._find_from_table_alias(table_new)
        if not table_alias_new:
            table_alias_new = self._generate_table_alias(table_new)
            self._add_prototypequery_table(table_new, table_alias_new)
        self._update_select_table_alias(table_field_old, table_alias_new)
        self._update_select_fields(table_field_old, field_new)
        self._update_orderby_table_alias(field_old, table_alias_new)
        self._update_orderby_field(field_old, field_new)
        if self.where:
            self._update_where_settings(field_old, field_new, table_alias_new)

        # Table field measures act like ids so update these last
        self._update_select_table_fields(table_field_old, table_field_new)

    def _find_from_table_alias(self, table) -> str:
        """Finds if a table is present as a source in the prototypequery object."""
        table_path = parse(f"$[?(@.Entity=='{table}')].Name")
        table = table_path.find(self.frm)
        if table:
            return table[0].value
        return

    def _return_from_tables(self):
        """Returns all the currently used table name aliases in the prototypequery."""
        table_path = parse(f"$.[*].Name")
        tables = table_path.find(self.frm)
        return [name.value for name in tables]

    def _generate_table_alias(self, table_new):
        """Returns a new table name alias for additions to the prototypequery."""
        # Power BI reassigns aliases when visual is updated so don't need to replicate exactly
        alias = table_new[:1].lower()
        regex = re.compile("[^0-9]")
        names = self._return_from_tables()
        names = [int(regex.sub("0", name)) for name in names if name[:1] == alias]
        if names:
            return alias + str(max(names) + 1)
        return alias

    def _add_prototypequery_table(self, table, name):
        """Adds a new table to the prototypequery."""
        table = {"Name": name, "Entity": table, "Type": 0}
        self.frm.append(table)

    def _update_select_fields(self, table_field, field):
        """Updating prototypequery fields."""
        path = parse(f"$[?(@.Name=='{table_field}')].*.Property")
        path.update(self.select, field)

    def _update_select_table_alias(self, table_field, name):
        """Updates prototypequery table name alias."""
        path = parse(f"$[?(@.Name=='{table_field}')].*.Expression.SourceRef.Source")
        path.update(self.select, name)

    def _update_select_table_fields(self, table_field_old, table_field_new):
        """Updating prototypequery table fields."""
        path = parse(f"$[?(@.Name=='{table_field_old}')].Name")
        path.update(self.select, table_field_new)

    def _return_select_tables(self, table_field_old):
        path = parse(f"$[?(@.Name!='{table_field_old}')].*.Expression.SourceRef.Source")
        nodes = path.find(self.select)
        return [node.value for node in nodes]

    def _cleanup_tables(self, table_field_old, table_old):
        table_alias_old = self._find_from_table_alias(table_old)
        selects = self._return_select_tables(table_field_old)
        wheres = self._return_where_tables(table_alias_old)
        if table_alias_old not in selects and table_alias_old not in wheres:
            for table in self.frm:
                if table["Name"] == table_alias_old:
                    self.frm.remove(table)

    def _update_orderby_table_alias(self, field_old, alias_new):
        path = parse(
            f"$[?(@.Expression.*.Property=='{field_old}')].Expression.*.Expression.SourceRef.Source"
        )
        path.update(self.order_by, alias_new)

    def _update_orderby_field(self, field_old, field_new):
        path = parse(
            f"$[?(@.Expression.*.Property=='{field_old}')].Expression.*.Property"
        )
        path.update(self.order_by, field_new)

    def _update_where_settings(self, field_old, field_new, name):
        for node in self.where:
            for _, setting in node.items():
                self._update_where_table_alias(field_old, name, setting)
                self._update_where_field(field_old, field_new, setting)

    def _update_where_table_alias(self, field_old, name, condition):
        path = parse(f"$..Property")
        if path.find(condition)[0].value == field_old:
            path = parse(f"$..Source")
            path.update(condition, name)

    def _update_where_field(self, field_old, field_new, condition):
        path = parse(f"$..Property")
        if path.find(condition)[0].value == field_old:
            path.update(condition, field_new)

    def _return_where_tables(self, alias):
        path = parse(f"$[?(@..Source!='{alias}')]..Source")
        nodes = path.find(self.where)
        return [node.value for node in nodes]


class VisualQuery:
    """A class representing the query settings of a visual"""

    def __init__(self, visual_query) -> None:
        self.visual_query = visual_query
        self.commands = self.visual_query.get("Commands")

    def update_fields(
        self,
        table_field_old,
        table_field_new,
        table_old,
        table_new,
        field_old,
        field_new,
    ):
        """Replace field in all relevant query settings."""
        for command in self.commands:
            query = GenericVisualQuery(
                command["SemanticQueryDataShapeCommand"]["Query"]
            )
            query.update_fields(
                table_field_old,
                table_field_new,
                table_old,
                table_new,
                field_old,
                field_new,
            )


class VisualDataTransforms:
    """A class representing the datatransforms settings of a Visual."""

    def __init__(self, data_transforms) -> None:
        self.data_transforms = data_transforms
        self.metadata = self.data_transforms.get("queryMetadata")

    def update_fields(
        self, table_field_old, table_field_new, table_new, field_old, field_new
    ):
        """Replace fields in all relevant datatransforms settings."""
        self._update_datatransforms_metadata(table_field_old, table_field_new)
        self._update_datatransforms_selects(table_field_old, table_new)
        self._update_datatransforms_selects_field(table_field_old, field_new)
        self._update_query_metadata_filters_table(field_old, table_new)
        self._update_query_metadata_filters_property(field_old, field_new)

        # Table field measures act like ids so update these last
        self._update_datatransforms_selects_table_field(
            table_field_old, table_field_new
        )
        self._update_query_meta_data(table_field_old, table_field_new)

    def _update_datatransforms_metadata(self, table_field_old, table_field_new):
        """Update table references in metadata."""
        path = parse(
            f"$.objects.*[?(@.selector.metadata=='{table_field_old}')].selector.metadata"
        )
        path.update(self.data_transforms, table_field_new)

    def _update_datatransforms_selects(self, table_field_old, table_new):
        """Update table references in selects."""
        path = parse(
            f"$.selects[?(@.queryName=='{table_field_old}')].expr.*.Expression.SourceRef.Entity"
        )
        path.update(self.data_transforms, table_new)

    def _update_datatransforms_selects_field(self, table_field_old, field):
        """Update field in selects."""
        path = parse(f"$.selects[?(@.queryName=='{table_field_old}')].expr.*.Property")
        path.update(self.data_transforms, field)

    def _update_datatransforms_selects_table_field(
        self, table_field_old, table_field_new
    ):
        """Update table.field in selects."""
        path = parse(f"$.selects[?(@.queryName=='{table_field_old}')].queryName")
        path.update(self.data_transforms, table_field_new)

    def _update_query_meta_data(self, table_field_old, table_field_new):
        """Update table.field in query metadata."""
        path = parse(f"$.queryMetadata.Select[?(@.Name=='{table_field_old}')].Name")
        path.update(self.data_transforms, table_field_new)

    def _update_query_metadata_filters_table(self, field_old, table_new):
        path = parse(
            f"$.Filters[?(@.expression.*.Property=='{field_old}')].expression.*.Expression.SourceRef.Entity"
        )
        path.update(self.metadata, table_new)

    def _update_query_metadata_filters_property(self, field_old, field_new):
        path = parse(
            f"$.Filters[?(@.expression.*.Property=='{field_old}')].expression.*.Property"
        )
        path.update(self.metadata, field_new)


class VisualFilters:
    def __init__(self, filters) -> None:
        self.filters = filters

    def update_fields(
        self,
        table_field_old,
        table_field_new,
        table_old,
        table_new,
        field_old,
        field_new,
    ):
        self._update_filters(
            table_field_old, table_field_new, table_old, table_new, field_old, field_new
        )
        self._update_table(field_old, table_new)
        self._update_field(field_old, field_new)

    def _update_table(self, field_old, table_new):
        path = parse(
            f"$[?(@.expression.*.Property=='{field_old}')].expression.*.Expression.SourceRef.Entity"
        )
        path.update(self.filters, table_new)

    def _update_field(self, field_old, field_new):
        path = parse(
            f"$[?(@.expression.*.Property=='{field_old}')].expression.*.Property"
        )
        path.update(self.filters, field_new)

    def _update_filters(
        self,
        table_field_old,
        table_field_new,
        table_old,
        table_new,
        field_old,
        field_new,
    ):
        path = parse(f"$[?(@.expression.*.Property=='{field_old}')].filter")
        for filter in path.find(self.filters):
            filtr = GenericVisualQuery(filter.value)
            filtr.update_fields(
                table_field_old,
                table_field_new,
                table_old,
                table_new,
                field_old,
                field_new,
            )

    def _clear_filters(self):
        for filter in self.filters:
            filter.pop("filter", None)


class Slicer(DataVisual):
    """A class representing a slicer."""

    def unselect_all_items(self) -> None:
        """Unselects all slicer members where no default selection is defined"""
        slicer_path = parse(
            "$.singleVisual.objects.data[*].properties.isInvertedSelectionMode.`parent`"
        )
        selection_path = parse("$.singleVisual.objects.general[*].properties.filter")
        slicer = slicer_path.find(self.config)
        if slicer and not selection_path.find(self.config):
            slicer[0].value.pop("isInvertedSelectionMode")
            self.layout["config"] = json.dumps(self.config)
            self.updated = 1
            print(f"Updated: {self.title}")
