import re

def extract_start_end_time(program_str):
    match = re.search(r'\d{2}\.\d{2}', program_str)
    if match:
        times = re.findall(r'\d{2}\.\d{2}', program_str)
        if len(times) == 2:
            return times[0], times[1]
    return "07.00", "24.00" # Default values if no match is found    