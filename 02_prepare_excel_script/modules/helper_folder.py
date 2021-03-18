# Regex library
import re


def get_name(name):
    obj = {}

    folder_name = re.search("^([HD](.+))(_\d{3})$", name)

    if folder_name:
        obj["full name"] = name
        obj["name"] = folder_name.group(1)
        obj["page"] = folder_name.group(3)
        # print(folder_name.groups())
    else:
        obj["name"] = name

    return obj
