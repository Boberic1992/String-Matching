import pandas as pd
import os
from fuzzywuzzy import fuzz
import Levenshtein
import re

# path to excel file where script is
excel_file_path = os.path.join(os.getcwd(), 'program_list_v2.xlsx')

# read excel file
df = pd.read_excel(excel_file_path, usecols=['Software name'])

def clean_string(s):
    # remove special characters, lower case, and split into words
    return re.sub(r"[^a-zA-Z0-9\s]", "", s).lower().split()

def find_best_match_among_100(input_name, matches):
    best_match = None
    highest_score = 0

    for match in matches:
        # calculate a score based on the length difference and fuzziness
        length_score = 100 - abs(len(input_name) - len(match))
        fuzz_score = fuzz.ratio(input_name.lower(), match.lower())
        total_score = (length_score + fuzz_score) / 2

        if total_score > highest_score:
            highest_score = total_score
            best_match = match

    return best_match

def find_matches(input_name, dataset):
    matches = {
        'Perfect match': [],
        'Above 90% match' : [],
        'Between 80% and 90% match': [],
        'Between 70% and 80% match': []
    }

    input_words = set(clean_string(input_name))

    for _, row in dataset.iterrows():
        db_name = row['Software name'].lower()
        db_words = set(clean_string(db_name))

        # enhanced partial match score calculation
        partial_match_score = 0
        for word in db_words:
            if any(word.startswith(inp) for inp in input_words):
                partial_match_score = 100
                break

        fuzz_score = fuzz.ratio(input_name.lower(), db_name)
        lev_distance = Levenshtein.distance(input_name.lower(), db_name)
        max_length = max(len(input_name), len(db_name))
        adjusted_score = 100 - (lev_distance / max_length) * 100
        # use the highest of the three scores
        score = max(fuzz_score, adjusted_score, partial_match_score)

        if score == 100:
            matches['Perfect match'].append(row['Software name'])
        elif score > 90:
            matches['Above 90% match'].append(row['Software name'])
        elif 80 <= score <= 90:
            matches['Between 80% and 90% match'].append(row['Software name'])
        elif 70 <= score < 80:
            matches['Between 70% and 80% match'].append(row['Software name'])

    # find the best match among 100% matches
    if len(matches['Perfect match']) > 1:
        best_100_match = find_best_match_among_100(input_name, matches['Perfect match'])
        matches['Perfect match'] = [best_100_match] if best_100_match else matches['Perfect match']
    
    return matches



def perform_search():
    # input program names
    input_programs = input('Enter program names (comma-separated): ').split(',')
    print()

    for program in input_programs:
        program = program.strip()
        matches = find_matches(program, df)

        # display results for each program
        print(f"Results for '{program}':")
        for category, names in matches.items():
            print(f"  {category}: {', '.join(names) if names else 'No matches'}")
        print()

# main loop
while True:
    perform_search()