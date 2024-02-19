import os
import re

def get_pl_path(file_paths):
    pattern = re.compile(r'.*PL.*', re.IGNORECASE)
    paths_with_pl = []
    for path in file_paths:
        filename = os.path.basename(path)
        if re.match(pattern, filename):
            paths_with_pl.append(path)
    return paths_with_pl


# test_list = ["pl", "part_spl"]

# for p in test_list:
#     if "pl" in p.casefold():
#         print("pass")
