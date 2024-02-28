#Cleans a boeing part list pdf 

import re

def capture_data(filepath, outpath):
    start_pattern = re.compile(r'\bREQD\s+IDENTIFYING\s+NUMBER\s+DESCRIPTION\s+CODE\s+SYM\b')
    start_pattern1 = re.compile(r"\.\s*\.\s*\.\s*ASSEMBLY BREAKDOWN LIST\s*\.\s*\.\s*\.")

    end_pattern = re.compile(r"\*\s*\*\s*\*\s*.*\*\s*\*\s*\*")
    end_pattern1 = r'\bCONTRACT\s+NUMBER\b'

    start_pattern2 = r"ASSY.*\(CONTINUED\s*ON\s*NEXT\s*PAGE\)"
    end_pattern2 = r"ASSY.*\(CONTINUED\s*FROM\s*PRECEDING\s*PAGE\)"
    capturing_data = False
    skip_mode = False
    with open(filepath, "r") as file:
        with open(outpath, "w") as outfile:
            lines = file.readlines()
            for line in  lines:
                var1 = line.split()
                var = " ".join(var1)
                if not capturing_data:
                    if start_pattern.match(var) or start_pattern1.match(var):
                        capturing_data = True
                elif capturing_data:
                    if end_pattern.match(var):
                        capturing_data = False
                    elif capturing_data:
                        if re.search(start_pattern2, var) or re.search(end_pattern1, var):
                            skip_mode = True
                        elif var1[0] == "QTY" or var1[0] == "REQD":
                            pass
                        elif re.search(end_pattern2, var):
                            skip_mode = False
                        elif not skip_mode:
                            outfile.write(line)
