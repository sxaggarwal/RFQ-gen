# Extracting the paths of PL documents from all the selected  Documents

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
