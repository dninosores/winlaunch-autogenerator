import argparse
import os
import pathlib
import hashlib

# Make it so that it just crawls local and global start menus, don't need to overcomplicate it
# However these options could be helpful:
# 1. Whether to replicate folder structure or put everything on one level (default all one level)
# 2. Whether to remove duplicates (and whether priority is given to appdata or programdata) (default remove duplicates with appdata priority)
# 3. Whether to exclude uninstallers (an uninstaller is a shortcut just called 'uninstall'
# in the same directory as another program, or of the form "uninstall.*PROGRAMNAME") (default true)
# 4. Whether to replace current configuration or add to it (use hashes to figure out if a link is already present)

CONFIG_LOC = pathlib.Path(os.getenv("APPDATA", "")).joinpath("WinLaunch")
LOCAL_STARTMENU = pathlib.Path(os.getenv("APPDATA", "")).joinpath(
    "Microsoft", "Windows", "Start Menu", "Programs"
)
GLOBAL_STARTMENU = pathlib.Path(os.getenv("PROGRAMDATA", "")).joinpath(
    "Microsoft", "Windows", "Start Menu", "Programs"
)

SEARCH_LOCATIONS = [LOCAL_STARTMENU, GLOBAL_STARTMENU]


parser = argparse.ArgumentParser(
    prog="WinLaunch Autogenerator",
    description=(
        "Searches for applications on your system and generates a WinLaunch config to"
        " display them."
    ),
)

parser.add_argument(
    "-o",
    "--output",
    default=CONFIG_LOC,
    help=(
        "Directory to save configuration files to. Will save to %APPDATA%/WinLaunch by"
        " default."
    ),
)

links = []

for location in SEARCH_LOCATIONS:
    for path in location.glob("**/*.lnk"):
        links.append(path)

print(links)
# Glob all .lnk files and make hashes out of their filepaths
