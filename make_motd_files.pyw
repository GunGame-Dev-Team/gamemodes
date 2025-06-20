from operator import itemgetter
from random import sample
import openpyxl
from path import Path


DOCUMENT_FILE = Path("F:\\Projects\\gamemodes\\GunGame Data.xlsx")
wb = openpyxl.load_workbook(DOCUMENT_FILE)
GAME_MODES = sorted(
    map(
        itemgetter(0),
        wb["Maps"].iter_rows(min_row=2, values_only=True)
    )
)

with (BASE_PATH / "template.html").open() as open_file:
    template_contents = open_file.read()

color_base = [0, 63, 127, 191, 255]
color_reverse = (255, 191, 127, 63, 0)
color_options = color_base * 3
used_colors = [[0, 0, 0], [255, 255, 255]]


def get_color():
    while True:
        value = sample(color_options, 3)
        if value not in used_colors:
            used_colors.append(value)
            return value


for game_mode in GAME_MODES:
    game_mode_file_name = "_".join(
        game_mode.replace("Free for All", "ffa").split()
    ).lower() + ".html"
    color = get_color()
    color2 = [color_reverse[color_base.index(i)] for i in color]
    contents = template_contents.replace(
        "{color}",
        "{:02X}{:02X}{:02X}".format(*color),
    ).replace(
        "{color2}",
        "{:02X}{:02X}{:02X}".format(*color2),
    ).replace(
        "{gamemode}",
        game_mode,
    )
    with (BASE_PATH / "templates" / game_mode_file_name).open("w") as open_file:
        open_file.write(contents)
