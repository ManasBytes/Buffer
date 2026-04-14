import pandas as pd
import os



answersheet_path = "answer.xlsx"
config_file_path = "noc26-cs79_S4.xlsx"
all_sheets = pd.read_excel(config_file_path, sheet_name=None)



# Get Configuration Details sheet
config_df = all_sheets["Configuration Details"]
english_df = all_sheets["English"]

# Filter for MCQ, MSQ, and SA question types
filtered_df = config_df[config_df['Question Type'].isin(['MCQ', 'MSQ', 'SA'])]

# Option id columns are right after 'No Of Options'
all_columns = list(config_df.columns)
no_of_options_index = all_columns.index('No Of Options')
option_source_columns = all_columns[no_of_options_index + 1:no_of_options_index + 5]


def normalize_id(value):
    if pd.isna(value):
        return ""
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value)


def option_token_to_number(token):
    normalized_token = str(token).strip().lower()
    if normalized_token == "":
        return ""

    if len(normalized_token) == 1 and "a" <= normalized_token <= "z":
        return str(ord(normalized_token) - ord("a") + 1)

    try:
        return str(int(float(normalized_token)))
    except (ValueError, TypeError):
        return ""


def concatenate_options(option_values, separator=","):
    ordered_non_empty = [value for value in option_values if value != ""]
    return separator.join(ordered_non_empty)


def parse_correct_option(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text == "":
        return ""

    normalized_values = []
    for token in text.split(","):
        option_number = option_token_to_number(token)
        if option_number == "":
            continue
        normalized_values.append(option_number)

    return ",".join(normalized_values)


def get_correct_option_id(concatenated_option_string, correct_option_number, separator=","):
    if correct_option_number == "" or pd.isna(correct_option_number):
        return ""

    option_ids = [item for item in str(concatenated_option_string).split(separator) if item != ""]
    correct_option_tokens = [token.strip() for token in str(correct_option_number).split(separator) if token.strip() != ""]

    mapped_correct_option_ids = []
    for token in correct_option_tokens:
        try:
            index = int(float(token)) - 1
        except (ValueError, TypeError):
            return ""
        if index < 0 or index >= len(option_ids):
            return ""
        mapped_correct_option_ids.append(option_ids[index])

    return separator.join(mapped_correct_option_ids)


def normalize_candidate_entry(value):
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def candidate_entry_to_redirect(candidate_entry, question_type):
    candidate_text = normalize_candidate_entry(candidate_entry)
    if candidate_text == "":
        return ""

    if question_type in ("MCQ", "MSQ"):
        mapped_values = []
        for token in candidate_text.split(","):
            option_number = option_token_to_number(token)
            if option_number == "":
                continue
            mapped_values.append(option_number)
        return ",".join(mapped_values)

    if question_type == "SA":
        return candidate_text

    return candidate_text


def redirect_to_selected_id(concatenated_option_string, redirect_to_number, separator=","):
    redirect_text = normalize_candidate_entry(redirect_to_number)
    if redirect_text == "":
        return ""

    option_ids = [item for item in str(concatenated_option_string).split(separator) if item != ""]
    selected_ids = []
    for token in redirect_text.split(separator):
        token = token.strip()
        if token == "":
            continue
        try:
            option_index = int(token) - 1
        except ValueError:
            return ""
        if option_index < 0 or option_index >= len(option_ids):
            return ""
        selected_ids.append(option_ids[option_index])

    return separator.join(selected_ids)


def normalize_marks(value):
    if pd.isna(value):
        return ""
    try:
        numeric_value = float(value)
        if numeric_value.is_integer():
            return int(numeric_value)
        return numeric_value
    except (ValueError, TypeError):
        return str(value).strip()


def split_non_empty_csv(value, separator=","):
    return [token.strip() for token in str(value).split(separator) if token.strip() != ""]


def evaluate_correct_mark_and_result(question_type, selected_id, correct_option_id, marks):
    if question_type not in ("MCQ", "MSQ"):
        return "", ""

    selected_tokens = split_non_empty_csv(selected_id)
    correct_tokens = split_non_empty_csv(correct_option_id)

    # Keep placeholders for unanswered rows.
    if len(selected_tokens) == 0:
        return "", ""

    if len(correct_tokens) == 0:
        return 0, "W"

    if question_type == "MCQ":
        if len(selected_tokens) == 1 and len(correct_tokens) == 1 and selected_tokens[0] == correct_tokens[0]:
            return marks, "C"
        return 0, "W"

    selected_set = set(selected_tokens)
    correct_set = set(correct_tokens)

    if selected_set == correct_set and len(selected_tokens) == len(correct_tokens):
        return marks, "C"

    if selected_set.issubset(correct_set):
        return 0, "PC"

    return 0, "W"


def normalize_column_name(column_name):
    return "".join(char for char in str(column_name).lower() if char.isalnum())


def build_candidate_entry_map(answer_sheet_file_path):
    candidate_entry_map = {}

    if str(answer_sheet_file_path).strip() == "":
        return candidate_entry_map

    if not os.path.exists(answer_sheet_file_path):
        print(f"Warning: answersheet_path not found: {answer_sheet_file_path}")
        return candidate_entry_map

    answer_sheets = pd.read_excel(answer_sheet_file_path, sheet_name=None)
    answer_column_candidates = [
        "enteryouranswer",
        "enteeryouranswer",
        "candidateentry",
        "answer",
    ]

    for sheet_name, answer_df in answer_sheets.items():
        standardized_columns = {
            normalize_column_name(column): column
            for column in answer_df.columns
        }

        question_id_column = standardized_columns.get("questionid")
        answer_column = None
        for answer_column_key in answer_column_candidates:
            if answer_column_key in standardized_columns:
                answer_column = standardized_columns[answer_column_key]
                break

        if question_id_column is None or answer_column is None:
            continue

        answer_question_ids = answer_df[question_id_column].apply(normalize_id)
        answer_values = answer_df[answer_column].apply(
            lambda value: "" if pd.isna(value) else str(value).strip()
        )

        for question_id, answer_value in zip(answer_question_ids, answer_values):
            if question_id != "":
                candidate_entry_map[question_id] = answer_value

        print(f"Using answer sheet tab: {sheet_name}")
        return candidate_entry_map

    print("Warning: no valid answer sheet tab found with Question ID and Enter Your Answer columns")
    return candidate_entry_map

# Normalize option ids in the exact source order from Configuration Details
option_1 = filtered_df[option_source_columns[0]].apply(normalize_id)
option_2 = filtered_df[option_source_columns[1]].apply(normalize_id)
option_3 = filtered_df[option_source_columns[2]].apply(normalize_id)
option_4 = filtered_df[option_source_columns[3]].apply(normalize_id)

# Create concatenated options in strict OPTION 1 -> OPTION 4 order
separator_value = ","
concatenated_options = [
    concatenate_options([o1, o2, o3, o4], separator_value)
    for o1, o2, o3, o4 in zip(option_1, option_2, option_3, option_4)
]

# Build a Question ID -> Correct Option map from English sheet
english_question_ids = english_df['Question ID'].apply(normalize_id)
english_correct_options = english_df['Correct Option'].apply(parse_correct_option)
correct_option_map = dict(zip(english_question_ids, english_correct_options))

# Match using Question id and capture Correct Option (supports single and multi values)
correct_option_values = filtered_df['Question id'].apply(normalize_id).map(correct_option_map).fillna("")

# Derive Correct Option Id from the concatenated option ids in strict order
correct_option_ids = [
    get_correct_option_id(option_list, correct_option, separator_value)
    for option_list, correct_option in zip(concatenated_options, correct_option_values)
]

# Candidate entries are loaded only when answersheet_path has a value.
candidate_entry_map = build_candidate_entry_map(answersheet_path)
if str(answersheet_path).strip() == "":
    candidate_entry_values = [""] * len(filtered_df)
else:
    candidate_entry_values = filtered_df['Question id'].apply(normalize_id).map(candidate_entry_map).fillna("").values

redirect_to_number_values = [
    candidate_entry_to_redirect(candidate_entry, question_type)
    for candidate_entry, question_type in zip(candidate_entry_values, filtered_df['Question Type'].values)
]
selected_id_values = [
    redirect_to_selected_id(option_list, redirect_to_number, separator_value)
    for option_list, redirect_to_number in zip(concatenated_options, redirect_to_number_values)
]

# Marks from Configuration Details for the filtered question rows
marks_values = filtered_df['Marks'].apply(normalize_marks).values

# Evaluate awarded mark and result for MCQ/MSQ using Selected Id vs Correct Option Id.
correct_mark_and_result = [
    evaluate_correct_mark_and_result(question_type, selected_id, correct_option_id, marks)
    for question_type, selected_id, correct_option_id, marks in zip(
        filtered_df['Question Type'].values,
        selected_id_values,
        correct_option_ids,
        marks_values,
    )
]
correct_mark_values = [item[0] for item in correct_mark_and_result]
result_values = [item[1] for item in correct_mark_and_result]

# Create a new dataframe with S.No, Question id, Question Type, OPTION 1-4, separator, and concatenated options
output_df = pd.DataFrame({
    'S.No': range(1, len(filtered_df) + 1),
    'Question id': filtered_df['Question id'].apply(normalize_id).values,
    'Question Type': filtered_df['Question Type'].values,
    'OPTION 1': option_1.values,
    'OPTION 2': option_2.values,
    'OPTION 3': option_3.values,
    'OPTION 4': option_4.values,
    'Separator': separator_value,
    'Concatenated Options': concatenated_options,
    'Correct Option': correct_option_values.values,
    'Correct Option Id': correct_option_ids,
    'Candidate Entry': candidate_entry_values,
    'Redirect To Number': redirect_to_number_values,
    'Selected Id': selected_id_values,
    'Marks': marks_values,
    'Correct Mark': correct_mark_values,
    'Result': result_values,
})

# Create a new Excel file with the filtered data
output_file = "filtered_questions.xlsx"
output_df.to_excel(output_file, index=False, sheet_name='Questions')

print(f"File created: {output_file}")
print(f"Total questions filtered: {len(output_df)}")
print(f"\nSummary:")
print(output_df.head(10))
print(f"\nQuestion Type Distribution:")
print(output_df['Question Type'].value_counts())