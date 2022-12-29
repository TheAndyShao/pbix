"""A module providing tools for manipulating thin Power BI reports"""

import json
import os
import zipfile as zf
from typing import Any, Iterable, Union

from jsonpath_ng.ext import parse

from pbix import visual as Visual

JsonType = dict[str, Union[list["JsonType"], "JsonType"]]


class Report:
    """A class to represent a thin Power BI report."""

    def __init__(self, filepath: str) -> None:
        self.filepath: str = filepath
        self.filename: str = os.path.basename(filepath)
        self.layout: dict[str, Any] = self._read_layout(filepath)
        self.config: JsonType = json.loads(self.layout.get("config"))
        self.bookmarks: list[JsonType] = self.config.get("bookmarks", {})
        # self.pages = self.layout.get('sections')
        self.updated: int = 0

    def update_fields(self, old: str, new: str) -> None:
        """Iterates through pages and visuals in a pbix and replaces specified measure/column."""
        print(f"Updating: {self.filename}")
        for i, j, visual in self._generic_visuals_generator():
            if visual.is_data_visual:
                visual = Visual.DataVisual(visual.layout)
                visual.update_fields(old, new)
                self.layout["sections"][i]["visualContainers"][j] = visual.layout
                self.updated += visual.updated
        for i, bookmark in enumerate(self.bookmarks):
            bookmark = Bookmark(bookmark)
            bookmark.update_fields(old, new)
            self.bookmarks[i] = bookmark.bookmark
        self.config["bookmarks"] = self.bookmarks
        self.layout["config"] = json.dumps(self.config)
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
                self.layout["sections"][i]["visualContainers"][j] = slicer.layout
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

    def write_json_layout(self) -> None:
        """Write the cleaned JSON object to file."""
        with open("layout.json", "w", encoding="utf-16") as outfile:
            layout = self._return_full_json_layout()
            json.dump(layout, outfile)

    def _read_layout(self, filepath: str) -> dict[str, Any]:
        """Return a JSON object of the layout file within the PBIX file."""
        with zf.ZipFile(filepath, "r") as zip_file:
            string = zip_file.read("Report/Layout").decode("utf-16")
            return json.loads(string)

    def _return_full_json_layout(self) -> dict[str, Any]:
        """Return a fully JSONified object of the layout file within the PBIX file."""
        string = json.dumps(self.layout)
        string = self._unescape_json_string(string)
        return json.loads(string)

    @staticmethod
    def _unescape_json_string(string: str):
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
        return string

    def _generic_visuals_generator(self) -> Iterable:
        """Generator for iterating through all visuals in a file."""
        for i, page in enumerate(self.layout["sections"]):
            visuals = page["visualContainers"]
            for j, visual in enumerate(visuals):
                yield i, j, Visual.GenericVisual(visual)


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


class Bookmark:
    """A class representing the bookmark settings of a report."""

    def __init__(self, bookmark: JsonType) -> None:
        self.bookmark: JsonType = bookmark
        self.name = bookmark.get("displayName")

    def update_fields(self, table_field_old: str, table_field_new: str) -> None:
        """Finds old references to old field and replacements with new field."""
        table_old, field_old = table_field_old.split(".")
        table_new, field_new = table_field_new.split(".")
        self._update_fields(table_old, table_new, field_old, field_new)

    def _update_fields(
        self, table_old: str, table_new: str, field_old: str, field_new: str
    ) -> None:
        root_path = "$..@[?(@.Property=='{field_old}')]"
        path = parse(root_path.format(field_old=field_old))
        nodes = path.find(self.bookmark)
        for node in nodes:
            if node.value.get("Expression").get("SourceRef").get("Entity") == table_old:
                field_path = parse(root_path.format(field_old=field_old) + ".Property")
                field_path.update(node, field_new)
                table_path = parse(
                    root_path.format(field_old=field_new)
                    + ".Expression.SourceRef.Entity"
                )
                table_path.update(node, table_new)
