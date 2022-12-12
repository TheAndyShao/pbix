"""A module providing tools for manipulating thin Power BI reports"""

import zipfile as zf
import json
import os
from jsonpath_ng.ext import parse # ext implements filter functionality to parse


class Report:
    """A class to represent a thin Power BI report."""

    def __init__(self, filepath: str) -> None:
        self.filepath: str = filepath
        self.filename: str = os.path.basename(filepath)
        self.layout: str = self._read_layout(filepath)
        self.layout_full_json: str = self._read_full_json_layout(filepath)
        self.field_set: set[str] = self.get_all_fields()
        self.updated: int = 0

    def get_all_fields(self) -> None:
        """Get a list of used fields in the pbix file."""
        filters_path = parse(
                '$.sections[*].visualContainers[*].filters[*].expression.Measure.Property'
            )
        measures_path = parse(
                '$.sections[*].visualContainers[*].config.singleVisual.projections[*].*.[*].queryRef'
            )

        filter_set = set([match.value for match in filters_path.find(self.layout_full_json)])
        measure_set = set([match.value for match in measures_path.find(self.layout_full_json)])
        return filter_set.union(measure_set)

    def find_instances(self, fields: list[str]) -> dict[str, bool]:
        """Compare input fields with the fields used in the pbix file."""
        matches = {}
        for field in fields:
            if '.' in field:
                field_set = self.field_set
            else:
                field_set = [measure.split('.')[-1] for measure in self.field_set]

            if field in field_set:
                matches[field] = True
        return matches

    def update_measures(self, old: str, new: str) -> None:
        """Iterates through pages and visuals in a pbix and replaces specified measure/column."""
        print(f'Updating: {self.filename}')
        for i, j, visual in self._generic_visuals_generator():
            if visual.is_data_visual:
                visual = DataVisual(visual)
                visual.update_measures(old, new)
                self._update_visual_layout(i, j, visual.layout)
                self.updated += visual.updated
        if self.updated == 0:
            print('No measures to update')

    def update_slicers(self) -> None:
        """Iterates through pages and genric visuals and updates slicers."""
        for i, j, visual in self._generic_visuals_generator():
            if visual.type == 'slicer':
                slicer = Slicer(visual.layout)
                slicer.unselect_all_items()
                self._update_visual_layout(i, j, slicer.layout)
                self.updated += slicer.updated
        if self.updated == 0:
            print('No slicers to update')

    def write_file(self) -> None:
        """Writes the pbix json to file."""
        base, ext = os.path.splitext(self.filename)
        temp_filepath = os.path.join(f"{base} Temp{ext}")

        with zf.ZipFile(self.filepath, "r") as original_zip:
            with zf.ZipFile(temp_filepath, "w", compression=zf.ZIP_DEFLATED) as new_zip:
                for file in original_zip.namelist():
                    if file == "Report/Layout":
                        new_zip.writestr(file, json.dumps(self.layout).encode("utf-16-le"))
                    elif file != "SecurityBindings":
                        new_zip.writestr(file, original_zip.read(file))

        os.remove(self.filepath)
        os.rename(temp_filepath, self.filepath)

    def _read_layout(self, filepath: str) -> str:
        """Return a JSON object of the layout file within the PBIX file."""
        with zf.ZipFile(filepath, 'r') as zip_file:
            string = zip_file.read('Report/Layout').decode("utf-16")
            return json.loads(string)

    def _read_full_json_layout(self, filepath: str) -> str:
        """Return a fully JSONified object of the layout file within the PBIX file."""
        with zf.ZipFile(filepath, 'r') as zip_file:
            string = zip_file.read('Report/Layout').decode("utf-16")
            replacements = {
                chr(0): "",
                chr(28): "",
                chr(29): "",
                chr(25): "",
                "\"[": "[",
                "]\"": "]",
                "\"{": "{",
                "}\"": "}",
                "\\\\": "\\",
                "\\\"": "\""
            }
            for original, replacement in replacements.items():
                string = string.replace(original, replacement)
            return json.loads(string)

    def write_json_layout(self) -> None:
        """Write the cleaned JSON object to file."""
        with open('layout.json', 'w', encoding="utf-16") as outfile:
            json.dump(self.layout_full_json, outfile)

    def _generic_visuals_generator(self) -> None:
        """Generator for iterating through all visuals in a file."""
        for i, page in enumerate(self.layout['sections']):
            visuals = page['visualContainers']
            for j, visual in enumerate(visuals):
                yield i, j, GenericVisual(visual)

    def _update_visual_layout(self, page: int, visual: int, layout: str) -> None:
        """Updates visual layout with new definition."""
        self.layout['sections'][page]['visualContainers'][visual] = layout

class GenericVisual:
    """A base class to represent a generic visual object."""

    def __init__(self, layout: str) -> None:
        self.layout: str = layout
        self.config: str = self._parse_visual_option('config')
        self.title: str or None = None
        self.type: str or None = self._return_visual_type()
        non_data_visuals = ['image', 'textbox', 'shape', 'actionButton', None]
        self.is_data_visual: bool = self.type not in non_data_visuals
        self.updated: int = 0

    def _return_visual_title(self) -> str or None:
        """Return title of visual."""
        title_path = parse("$.singleVisual.vcObjects.title[0].properties.text.expr.Literal.Value")
        title = title_path.find(self.config)
        return title[0].value if title else None

    def _return_visual_type(self) -> str or None:
        """Return type of visual."""
        typ_path = parse("$.singleVisual.visualType")
        typ = typ_path.find(self.config)
        return typ[0].value if typ else None

    def _parse_visual_option(self, visual_option: str) -> str or None:
        """Returns a JSON object or None from visual option string"""
        if visual_option in self.layout.keys():
            return json.loads(self.layout[visual_option])
        return None


class DataVisual(GenericVisual):
    """A class representing visuals that depend on a data model."""
    def __init__(self, Visual: GenericVisual) -> None:
        super().__init__(Visual.layout)
        self.title: str = self._return_visual_title()
        self.filters: str = self._parse_visual_option('filters')
        self.query: str = self._parse_visual_option('query')
        self.data_transforms: str = self._parse_visual_option('dataTransforms')

    def update_measures(self, old: str, new: str) -> None:
        """Searches for relevant keys for measures and updates their value pairs."""
        if self.query: # Ignore shapes, textboxes etc.

            old_table, old_measure = old.split('.')
            new_table, new_measure = new.split('.')

            measure_path = parse(
                    f"$..@[?(@.*=='{old_measure}')].[Property, displayName, Restatement]"
                )
            table_path = parse(
                    f"$..@[?(@.*=='{old_table}')].Entity"
                )
            table_measure_path = parse(
                    f"$..@[?(@.*=='{old}')].[queryRef, Name, queryName]"
                )

            if measure_path.find(self.config) or measure_path.find(self.filters):
                visual_options = {
                    "config": self.config,
                    "filters": self.filters,
                    "query": self.query,
                    "dataTransforms": self.data_transforms
                }
                for option, value in visual_options.items():
                    measure_path.update(value, new_measure)
                    table_path.update(value, new_table)
                    table_measure_path.update(value, new)
                    self.layout[option] = json.dumps(value)
                self.updated = 1

                print(f"Updated: {self.title}")


class NonDataVisual(GenericVisual):
    """A calss representing visuals that don't depend on a data model."""


class Slicer(GenericVisual):
    """A class representing a slicer."""

    def unselect_all_items(self) -> None:
        """Unselects all slicer members where no default selection is defined"""
        slicer_path = parse(
            "$.singleVisual.objects.data[*].properties.isInvertedSelectionMode.`parent`"
        )
        selection_path = parse("$.singleVisual.objects.general[*].properties.filter")
        slicer = slicer_path.find(self.config)
        if slicer and not selection_path.find(self.config):
            slicer[0].value.pop('isInvertedSelectionMode')
            self.layout['config'] = json.dumps(self.config)
            self.updated = 1
            print(f"Updated: {self.title}")
