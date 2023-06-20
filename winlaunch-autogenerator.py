import argparse
import os
import pathlib
import hashlib
from xml.dom import minidom
import re

# Make it so that it just crawls local and global start menus, don't need to overcomplicate it
# However these options could be helpful:
# 1. Whether to replicate folder structure or put everything on one level (default all one level)
# 2. Whether to remove duplicates (and whether priority is given to appdata or programdata) (default remove duplicates with appdata priority)
# 3. Whether to exclude uninstallers (an uninstaller is a shortcut just called 'uninstall'
# in the same directory as another program, or of the form "uninstall.*PROGRAMNAME") (default true)
# 4. Whether to replace current configuration or add to it (use hashes to figure out if a link is already present)
# Whether to exclude the startup folder
# 5. Regex to exclude certain file patterns

CONFIG_LOC = pathlib.Path(os.getenv("APPDATA", "")).joinpath("WinLaunch")
LOCAL_STARTMENU = pathlib.Path(os.getenv("APPDATA", "")).joinpath(
    "Microsoft", "Windows", "Start Menu", "Programs"
)
GLOBAL_STARTMENU = pathlib.Path(os.getenv("PROGRAMDATA", "")).joinpath(
    "Microsoft", "Windows", "Start Menu", "Programs"
)

SEARCH_LOCATIONS = [LOCAL_STARTMENU, GLOBAL_STARTMENU]

ITEMS_PER_PAGE = 40


# Winlaunch expects link hashes to be formatted as a dash-separated string of 8-4-4-4-12 digits
def get_hash(object: str):
    hasher = hashlib.sha1()
    hasher.update(bytes(object, "UTF-8"))
    rawhash = hasher.hexdigest()
    formatted_hash = f"{rawhash[0:8]}-{rawhash[8:12]}-{rawhash[12:16]}-{rawhash[16:20]}-{rawhash[20:32]}"
    return formatted_hash


def add_node(
    parent: minidom.Element,
    document: minidom.Document,
    name: str,
    value: str | None = None,
) -> minidom.Element:
    newnode = document.createElement(name)
    if value:
        newnode.appendChild(document.createTextNode(value))
    parent.appendChild(newnode)
    return newnode


parser = argparse.ArgumentParser(
    prog="WinLaunch Autogenerator",
    description=(
        "Searches for applications on your system and generates a WinLaunch config to"
        " display them."
    ),
)


links: list[pathlib.Path] = []

for location in SEARCH_LOCATIONS:
    for path in location.glob("**/*.lnk"):
        links.append(path)

links.sort(key=lambda path: path.name)

# Glob all .lnk files and make hashes out of their filepaths

added_links = set()

config = minidom.Document()

config_root = config.createElement("ArrayOfICItem")
config_root.setAttribute(r"xmlns:xsd", r"http://www.w3.org/2001/XMLSchema")
config_root.setAttribute(r"xmlns:xsi", r"http://www.w3.org/2001/XMLSchema-instance")
config.appendChild(config_root)

page = 0
grid_index = 0

backup_path = CONFIG_LOC.joinpath("ICBackup")
linkcache_path = CONFIG_LOC.joinpath("LinkCache")
items_path = CONFIG_LOC.joinpath("Items.xml")

if backup_path.exists():
    os.system('rmdir /S /Q "{}"'.format(backup_path.as_posix()))
if linkcache_path.exists():
    os.system('rmdir /S /Q "{}"'.format(linkcache_path.as_posix()))
    os.mkdir(linkcache_path)
if items_path.exists():
    os.remove(items_path)

for link in links:
    if re.match(r".* uninstall\.lnk", link.name.lower()) or re.match(
        r"^uninstall .*", link.name.lower()
    ):
        print(f"Skipping uninstaller: {link}")
        continue
    if re.match(r".*/Programs/Startup/.*$", link.as_posix()):
        print(f"Skipping startup program: {link}")

    if link.stem in added_links:
        print(f"Skipping duplicate program: {link}")
        continue

    linkhash = get_hash(link.as_posix())
    linkhash_filename = linkhash + ".lnk"

    link_item = config.createElement("ICItem")
    link_item.appendChild(config.createElement("Items"))
    add_node(link_item, config, "Items")
    add_node(link_item, config, "IsFolder", "false")
    add_node(link_item, config, "GridIndex", str(grid_index))
    add_node(link_item, config, "Page", str(page))
    add_node(link_item, config, "Name", link.stem)
    add_node(link_item, config, "Keywords")
    add_node(link_item, config, "Application", linkhash_filename)
    add_node(link_item, config, "Arguments")
    add_node(link_item, config, "RunAsAdmin", "false")
    add_node(link_item, config, "ShowMiniatures", "true")
    config_root.appendChild(link_item)

    # shutil.copyfile(link, linkcache_path.joinpath(linkhash_filename))

    copycmd = (
        f'copy "{link.as_posix()}"'
        + f' "{linkcache_path.joinpath(linkhash_filename).as_posix()}"'
    )
    copycmd = copycmd.replace("/", "\\")
    print(copycmd)
    os.system(copycmd)

    added_links.add(link.stem)

    grid_index += 1
    if grid_index >= ITEMS_PER_PAGE:
        grid_index = 0
        page += 1

with open(items_path, "w") as itemsfile:
    itemsfile.write(config.toprettyxml())
