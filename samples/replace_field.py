import multiprocessing
import os
from pathlib import Path

from pbix.report import Report


def replace_field(path, old, new):
    """Replace field in a single pbix file."""
    report = Report(path)
    report.update_fields(old, new)
    if report.updated > 0:
        report.write_file()


def replace_field_in_all_reports(path, old, new, model):
    """Replace field in all pbix files in specified directory and all child directories."""
    pool = multiprocessing.Pool()
    for subdir, _, files in os.walk(path):
        pool.starmap(
            replace_field,
            [
                (os.path.join(subdir, file), old, new)
                for file in files
                if file.split(".")[-1] == "pbix" and file != model
            ],
        )


def main(path, old, new, model):
    """Replace field fuction."""
    path = Path(path)
    if path.is_file():
        replace_field(path, old, new)
    else:
        replace_field_in_all_reports(path, old, new, model)


if __name__ == "__main__":
    # Update these variables
    path = "Path to file or parent directory"
    old = "Table.Measure"
    new = "Table.Measure"
    model = "Model.pbix"

    main(path, old, new, model)
