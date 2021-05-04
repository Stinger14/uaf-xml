import sys
import pathlib


def get_resources_path(relative_path):
    """Return relative path to root.

    relative_path = data/mapped_elements.xlsx
       relative_path = pathlib.Path("data") / "mapped_elements.xlsx"
       relative_path = os.join.path("data", "mapped_elements.xlsx")

    Args:
        relative_path ([type]): [description]
    """

    rel_path = pathlib.Path(relative_path)
    dev_base_path = pathlib.Path(__file__).resolve().parent.parent
    base_path = getattr(sys, "_MEIPASS", dev_base_path)
    return base_path / rel_path
