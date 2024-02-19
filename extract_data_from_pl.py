import re

def extract_finish_codes_from_file(file_path, dash_number):
    finish_codes = set()
    identified_data_set = set()
    pattern = re.compile(r'.*?(?P<Material>\d+)-(?P<Temper>[A-Z0-9]+)\s*(?P<Shape>[A-Z\d]*\s*[A-Z\d]*\s*[A-Z\d]*)?\s+PER\s+(?P<Spec>[A-Z-]+\d+/\d+).*')

    current_part_number = None

    with open(file_path, 'r') as file:
        for line in file:
            var = line.split()
            if var and (var[0].startswith('-') or var[0].isdigit()) and var[1].startswith('-'):
                current_part_number = var[1]

            if current_part_number and current_part_number == dash_number and current_part_number not in identified_data_set:
                var1 = ' '.join(var)
                match = pattern.match(var1)
                for x in var:
                    if x.startswith('F-') or x.startswith('SRF-'):
                        finish_codes.add(x)

                if match:
                    identified_data = {
                        "Material": match.group('Material'),
                        "Temper": match.group('Temper'),
                        "Shape": match.group('Shape'),
                        "Spec": match.group('Spec')
                    }
                    identified_data_set.add(current_part_number)
                    return list(finish_codes), [identified_data]

    return list(finish_codes), []

def extract_dash_number(part_number):
    match = re.search(r'-\d+$', part_number)
    if match:
        dash_number = match.group(0)
    return dash_number

# Example usage:
# file_path = "example.txt"  # Provide the file path here
# part_number_to_search = "1234"  # Provide the part number here

# result, part_info_list = extract_finish_codes_from_file(file_path, part_number_to_search)

# if result:
#     print(f"Finish codes for part {part_number_to_search}: {result}")
#     if part_info_list:
#         print("Material Information:")
#         for part_info in part_info_list:
#             print(part_info)
#     else:
#         print("No material information available for this part.")
# else:
#     print(f"No Finish codes for {part_number_to_search}")
