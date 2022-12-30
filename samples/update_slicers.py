import multiprocessing
import os
from pathlib import Path

from pbix.report import Report


# Current update_slicers just unselects all items for multiselect slicers with 'All' as an option
def update_slicers(path):
    """Update slicers in a single pbix file."""
    report = Report(path)
    report.update_slicers()
    if report.updated > 0:
        report.write_file()


def update_slicers_in_all_reports(path, model):
    """Update slicers in all pbix files in specified directory and all child directories."""
    pool = multiprocessing.Pool(1)
    for subdir, _, files in os.walk(path):
        pool.map(
            update_slicers,
            [
                os.path.join(subdir, file)
                for file in files
                if file.split(".")[-1] == "pbix" and file != model
            ],
        )


def main(path, model):
    """Replace field fuction."""
    path = Path(path)
    if path.is_file():
        update_slicers(path)
    else:
        update_slicers_in_all_reports(path, model)


if __name__ == "__main__":
    # Update these variables
    path = r"/Users/andyshao/My Files/Diageo/Europe/cpd_powerbi/pbi/CPD"
    model = "Model.pbix"

    main(path, model)
