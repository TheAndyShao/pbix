import zipfile as zf
import json
import os
from jsonpath_ng.ext import parse # ext implements filter functionality to parse

class PBIFile:

    query_filter = parse('$.sections[*].visualContainers[*].query')
    filter_filter = parse('$.sections[*].visualContainers[*].filters')
    config_filter = parse('$.sections[*].visualContainers[*].config')

    def __init__(self, filepath):
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.layout = self.read_layout(filepath)
        self.field_set = None
        self.layout_modified = self.read_modified_layout(filepath)
        self.updated = 0


    def read_layout(self, filepath):
        """Return a cleaned JSON object of the layout file within the PBIX file"""
        with zf.ZipFile(filepath, 'r') as zip_file:
            data = zip_file.read('Report/Layout').decode("utf-16")
            layout = json.loads(data)
            return layout


    def read_modified_layout(self, filepath):
        """Return a cleaned JSON object of the layout file within the PBIX file"""
        with zf.ZipFile(filepath, 'r') as zip_file:
            data = zip_file.read('Report/Layout').decode("utf-16")
            string = data.replace(chr(0), "").replace(chr(28), "").replace(chr(29), "").replace(chr(25), "").replace("\"[", "[").replace("]\"", "]").replace("\"{", "{").replace("}\"", "}").replace("\\\\", "\\").replace("\\\"", "\"")
            layout_modified = json.loads(string)
            return layout_modified


    def _write_modified_layout(self):
        """Write the cleaned JSON object to file"""
        with open('layout.json', 'w', encoding="utf-16") as outfile:
            json.dump(self.layout_modified, outfile)


    def update_measures(self, old, new):
        """Iterates through pages and visuals in a pbix and replaces specified measure/column"""
        print(f'Updating: {self.filename}')
        pages = self.layout['sections']
        for i, page in enumerate(pages):
            visuals = page['visualContainers']
            for j, visual in enumerate(visuals):
                vis = Visual(visual)
                vis.update_measures(old, new)
                self.layout['sections'][i]['visualContainers'][j] = vis.layout_string
                self.updated += vis.updated
        if self.updated == 0:
            print('No measures to update')

    def generic_visuals_generator(self):
        """Generator for iterating through all visuals in a file"""
        for i, page in enumerate(self.layout['sections']):
            visuals = page['visualContainers']
            for j, visual in enumerate(visuals):
                yield i, j, Visual(visual)

    def _update_visual_layout(self, page, visual, layout):
        """Updates visual layout with new definition."""
        self.layout['sections'][page]['visualContainers'][visual] = layout

    def update_slicers(self):
        """Iterates through pages and genric visuals and updates slicers."""
        for i, j, visual in self.generic_visuals_generator():
            if visual.type == 'slicer':
                slicer = Slicer(visual.layout_string)
                slicer.unselect_all_items()
                self._update_visual_layout(i, j, slicer.layout_string)
                self.updated += slicer.updated
        if self.updated == 0:
            print('No slicers to update')

    def get_all_fields(self):
        """Get a list of used fields in the pbix file"""
        jsonpath_filters = parse('$.sections[*].visualContainers[*].filters[*].expression.Measure.Property')
        jsonpath_measures = parse('$.sections[*].visualContainers[*].config.singleVisual.projections[*].*.[*].queryRef')

        filter_set = set([match.value for match in jsonpath_filters.find(self.layout_modified)])
        measure_set = set([match.value for match in jsonpath_measures.find(self.layout_modified)])

        self.field_set = filter_set.union(measure_set)


    def find_instances(self, fields):
        """Compare input fields with the fields used in the pbix file"""
        self.get_all_fields()

        matches = {}
        for field in fields:
            if '.' in field:
                field_set = self.field_set
            else:
                field_set = [measure.split('.')[-1] for measure in self.field_set]

            if field in field_set:
                matches[field] = True
        return matches


    def write_file(self):
        """Writes the pbix json to file"""
        _, filename = os.path.split(self.filepath)
        base, ext = os.path.splitext(filename)
        temp_filepath = os.path.join(f"{base} Temp{ext}")

        with zf.ZipFile(self.filepath, "r") as original_zip:
            with zf.ZipFile(temp_filepath, "w", compression=zf.ZIP_DEFLATED) as new_zip:
                for f in original_zip.namelist():
                    if f == "Report/Layout":
                        new_zip.writestr(f, json.dumps(self.layout).encode("utf-16-le"))
                    elif f != "SecurityBindings":
                        new_zip.writestr(f, original_zip.read(f))

        os.remove(self.filepath)
        os.rename(temp_filepath, self.filepath)


class Visual:
    def __init__(self, layout):
        self.layout_string = layout
        self.x = self.layout_string['x']
        self.y = self.layout_string['y']
        self.z = self.layout_string['z']
        self.width = self.layout_string['width']
        self.height = self.layout_string['height']
        self.config = json.loads(self.layout_string['config'])
        self.title = self.find_title()
        self.type = self.find_type()
        try:
            self.filters = json.loads(self.layout_string['filters'])
            self.query = json.loads(self.layout_string['query'])
            self.dataTransforms = json.loads(self.layout_string['dataTransforms'])
        except KeyError:
            self.filters = None
            self.query = None
            self.dataTransforms = None
        self.updated = 0

    def find_title(self):
        """
        Find title of visual
        """
        title_filter = parse(f"$..@.title[*].properties.text.expr.Literal.Value")
        title = title_filter.find(self.config)
        if title:
            title = title[0].value
        else:
            title = None
        return title

    def find_type(self):
        """
        Find type of visual
        """
        vis_type_filter = parse(f"$.singleVisual.visualType")
        vis_type = vis_type_filter.find(self.config)
        if vis_type:
            vis_type = vis_type[0].value
        else:
            vis_type = None
        return vis_type

    def update_measures(self, old, new):
        """
        Searches for relevant keys for measures and updates their value pairs
        """

        if self.query: # Ignore shapes, textboxes etc.

            old_table, old_measure = old.split('.')
            new_table, new_measure = new.split('.')

            # ? filter condition can only be used when looking inside a element, i.e. $..@[?(@) not $..[?(@)
            measure_filter = parse(f"$..@[?(@.*=='{old_measure}')].[Property, displayName, Restatement]")
            table_filter = parse(f"$..@[?(@.*=='{old_table}')].Entity")
            table_measure_filter = parse(f"$..@[?(@.*=='{old}')].[queryRef, Name, queryName]")

            if measure_filter.find(self.config) or measure_filter.find(self.filters):
                sections = [self.config, self.filters, self.query, self.dataTransforms]
                for section in sections:
                    measure_filter.update(section, new_measure)
                    table_filter.update(section, new_table)
                    table_measure_filter.update(section, new)

                self.layout_string['config'] = json.dumps(self.config)
                self.layout_string['filters'] = json.dumps(self.filters)
                self.layout_string['query'] = json.dumps(self.query)
                self.layout_string['dataTransforms'] = json.dumps(self.dataTransforms)

                self.updated = 1

                print(f"Updated: {self.title}")


class Slicer(Visual):
    """A class representing a slicer."""

    def unselect_all_items(self):
        """Unselects all slicer members where no default selection is defined"""
        slicer_path = parse("$.singleVisual.objects.data[*].properties.isInvertedSelectionMode.`parent`")
        selection_path = parse("$.singleVisual.objects.general[*].properties.filter")
        slicer = slicer_path.find(self.config)
        if slicer and not selection_path.find(self.config):
            slicer[0].value.pop('isInvertedSelectionMode')
            self.layout_string['config'] = json.dumps(self.config)
            self.updated = 1
            print(f"Updated: {self.title}")
