import csv
from os.path import join, dirname
import re

# Dictionary of keywords-pattern
patterns_dict = {
    "Prov": r"\w{2}",
    "NRI": r".+",
    "Sezioni RI": r"[\w -]+",
    "Data iscrizione RI": r"\d{2}/\d{2}/\d{4}",
    "N.REA": r"\d{6}",
    "F.G.": r"\w{2}",
    "Denominazione": r".+",
    "C.fiscale": r"\d{11}",
    "Partita IVA": r"\d{11}",
    "Telefono": r'\d{2,3}[/]{1}\d{6,7}',
    "Indirizzo": r".+",
    "CAP": r"\d{5}",
    "Comune": r".+",
    "Indirizzo posta certificata": r"[\w\.-]+@[\w\.-]+\.\w{2,4}",
    "indipendenti": r"\d+",
    "dipendenti": r"\d+",
    "Data inizio attività": r"\d{2}/\d{2}/\d{4}",
    "Attività": r".+",
    "C. Attività": r"[\w \/]+",
    "Capitale Sociale": r"deliberato [\d,.]+",
    "Valuta capitale sociale": r"EURO"
}
# Dictionary of sub-pattern
sub_patterns_dict = {
    "Comune": {
        "CAP": r"\d{5}",
        "Comune": r"[^\d]+?",
        "Provincia": r"\w{2}"
    }
}
# List of keywords excluded from single line
excluded_keywords_list = [
    "Data iscrizione RI", "Denominazione", "Telefono", "Indirizzo", "Comune", "Frazione",
    "Alt. ind.", "Indirizzo posta certificata", "Data inizio attività", "Attività",
    "Valuta capitale sociale"
]
# List of patterns to be excluded
excluded_patterns_list = [r'^\d+\)', r'^ulisse', r'InfoCamere$']


def read_file(file_path):
    """Reads the content of the text file and returns it as a list of lines."""
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.readlines()


def apply_sub_patterns(line, sub_patterns):
    """Applies the sub-patterns to the given line."""
    for key, sub_pattern in sub_patterns.items():
        regex_parts = [f"({pattern})" for pattern in sub_pattern.values()]
        full_pattern = f"{key}: {' '.join(regex_parts)}"
        replacement_parts = [f"{subkey}: \\{i + 1}" for i, subkey in enumerate(sub_pattern.keys())]
        replacement = "\n".join(replacement_parts)
        line = re.sub(full_pattern, replacement, line)
    return line


def should_exclude_line(line, excluded_patterns):
    """Checks if a line should be excluded based on the excluded patterns."""
    for pattern in excluded_patterns:
        if re.match(pattern, line) or re.search(pattern, line):
            return True
    return False


def split_into_blocks(lines, patterns, sub_patterns, excluded_keys, excluded_patterns):
    """Splits the file content into blocks, one for each company."""
    blocks = []
    current_block = []
    inside_block = False

    for line in lines:
        if re.match(r'^\d+\)', line):
            if current_block:
                blocks.append("\n".join(current_block))
                current_block = []
            inside_block = True

        if inside_block:
            line = re.sub(r' \n', r'\n', line)
            for key, pattern in patterns.items():
                if key not in excluded_keys:
                    line = re.sub(rf' {key}: ', rf'\n{key}: ', line)
                    line = apply_sub_patterns(line, sub_patterns)

            formatted_lines = []
            for subline in line.split('\n'):
                if should_exclude_line(subline, excluded_patterns):
                    continue
                formatted_subline = re.sub(r'([^:]+): (.+)', r'\1: "\2"', subline)
                formatted_lines.append(formatted_subline)

            current_block.append("\n".join(formatted_lines))
            if "Valuta capitale sociale: EURO" in line:
                blocks.append("".join(current_block))
                current_block = []
                inside_block = False

    if current_block:
        blocks.append("".join(current_block))

    return blocks


def sanitize_blocks(blocks, patterns):
    """Cleans the blocks to ensure all keys are present and values are consistent."""
    sanitized_blocks = []
    keys = patterns.keys()

    for block in blocks:
        block_data = {key: "" for key in keys}  # Initialize dictionary with empty strings
        for line in block.split('\n'):
            match = re.match(r'([^:]+): "(.+)"', line)
            if match:
                key, value = match.groups()
                if key in block_data:
                    match_pattern = re.match(rf"{patterns[key]}", value)
                    if match_pattern:
                        block_data[key] = value.strip(" -")
                    else:
                        raise ValueError(f"'{value}' not matching with pattern '{patterns[key]}' for keyword '{key}'")

        sanitized_blocks.append(block_data)

    return sanitized_blocks


def write_to_csv(data, csv_path):
    """Writes the data to a CSV file."""
    if not data:
        print("No data to write to CSV file.")
        return

    keys = data[0].keys()
    with open(csv_path, 'w', newline='', encoding='utf-8') as output_file:
        dict_writer = csv.DictWriter(output_file, fieldnames=keys)
        dict_writer.writeheader()
        dict_writer.writerows(data)


def main():
    # Path to the input txt file
    input_txt_file = join(dirname(__file__), "./txt", "elenco Marconi.txt")

    # Read the file content
    lines = read_file(input_txt_file)

    # Split the content into blocks
    blocks = split_into_blocks(lines, patterns_dict, sub_patterns_dict, excluded_keywords_list, excluded_patterns_list)

    # Sanitize the blocks to ensure all keys are present
    sanitized_blocks = sanitize_blocks(blocks, patterns_dict)

    # Path to the output csv file
    output_csv_file = join(dirname(__file__), "./csv", "elenco Marconi.csv")

    # Write the extracted data to csv file
    write_to_csv(sanitized_blocks, output_csv_file)

    print("Data extraction and CSV file creation completed.")


if __name__ == '__main__':
    main()
