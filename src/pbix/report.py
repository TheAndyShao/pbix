"""A module providing tools for manipulating thin Power BI reports"""

import json
import os
import zipfile as zf
from typing import Any, Iterable

from jsonpath_ng.ext import parse

from pbix import visual as Visual


class Report:
    """A class to represent a thin Power BI report."""

    def __init__(self, filepath: str) -> None:
        self.filepath: str = filepath
        self.filename: str = os.path.basename(filepath)
        self.layout: dict[str, Any] = self._read_layout(filepath)
        self.layout_full_json: dict[str, Any] = self._read_full_json_layout(filepath)
        self.field_set: set[str] = self.get_all_fields()
        # self.pages = self.layout.get('sections')
        self.updated: int = 0

    def get_all_fields(self) -> set[str]:
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
                visual = Visual.DataVisual(visual)
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
                slicer = Visual.Slicer(visual)
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

    def _read_layout(self, filepath: str) -> dict[str, Any]:
        """Return a JSON object of the layout file within the PBIX file."""
        with zf.ZipFile(filepath, "r") as zip_file:
            string = zip_file.read("Report/Layout").decode("utf-16")
            return json.loads(string)

    def _read_full_json_layout(self, filepath: str) -> dict[str, Any]:
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

    def _generic_visuals_generator(self) -> Iterable:
        """Generator for iterating through all visuals in a file."""
        for i, page in enumerate(self.layout["sections"]):
            visuals = page["visualContainers"]
            for j, visual in enumerate(visuals):
                yield i, j, Visual.GenericVisual(visual)

    def _update_visual_layout(self, page: int, visual: int, layout: str) -> None:
        """Updates visual layout with new definition."""
        self.layout["sections"][page]["visualContainers"][visual] = layout


class ReportPage:
    """A class representing a single page within a report."""

    def __init__(self, page) -> None:
        self.page = page
        self.filters = Visual.Filters(json.loads(self.page.get("filters")))
        self.updated = 0

    def update_fields(self, table_field_old: str, table_field_new: str) -> None:
        """Finds usage of an existing field and replaces it with a new specified field."""
        # TODO: Currently this causes report level slicers to break,
        # even when the slicers conditions are cleared.
        # Return to in the future to enable report level slicer updating.
        table_old, field_old = table_field_old.split(".")
        table_new, field_new = table_field_new.split(".")
        self.filters.update_fields(
            table_field_old, table_field_new, table_old, table_new, field_old, field_new
        )

        self.page["filters"] = self.filters.filters
