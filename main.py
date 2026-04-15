import pandas as pd
import os
import math



answersheet_path = "sam.xlsx"
config_file_path = "noc26-cs79_S4.xlsx"
all_sheets = pd.read_excel(config_file_path, sheet_name=None)



# Get Configuration Details sheet
config_df = all_sheets["Configuration Details"]
english_df = all_sheets["English"]

# Filter for MCQ, MSQ, and SA question types
filtered_df = config_df[config_df['Question Type'].isin(['MCQ', 'MSQ', 'SA'])]


def find_column(columns, candidates):
    def normalize_lookup(value):
        return "".join(ch.lower() for ch in str(value).strip() if ch.isalnum())

    normalized_map = {normalize_lookup(column): column for column in columns}
    normalized_candidates = [normalize_lookup(candidate) for candidate in candidates]

    for normalized_candidate in normalized_candidates:
        if normalized_candidate in normalized_map:
            return normalized_map[normalized_candidate]

    for normalized_candidate in normalized_candidates:
        for normalized_column, original_column in normalized_map.items():
            if normalized_column.startswith(normalized_candidate) or normalized_candidate.startswith(normalized_column):
                return original_column

    return None


config_question_id_col = find_column(config_df.columns, ["Question id", "Question ID", "QuestionId"])
config_question_type_col = find_column(config_df.columns, ["Question Type", "QuestionType"])
config_marks_col = find_column(config_df.columns, ["Marks", "Mark"])

if config_question_id_col is None:
    raise KeyError("Missing required column in Configuration Details: Question id")
if config_question_type_col is None:
    raise KeyError("Missing required column in Configuration Details: Question Type")
if config_marks_col is None:
    raise KeyError("Missing required column in Configuration Details: Marks")

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


def normalize_question_id_for_match(value):
    return normalize_id(value).strip()


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
        return "", False

    option_ids = [item for item in str(concatenated_option_string).split(separator) if item != ""]
    selected_ids = []
    for token in redirect_text.split(separator):
        token = token.strip()
        if token == "":
            continue
        try:
            option_index = int(token) - 1
        except ValueError:
            return "", True
        if option_index < 0 or option_index >= len(option_ids):
            return "", True
        selected_ids.append(option_ids[option_index])

    return separator.join(selected_ids), False


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


def normalize_sa_answer(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def parse_yes_no(value):
    text = "" if pd.isna(value) else str(value).strip().lower()
    if text in ("yes", "y", "true", "1"):
        return True
    if text in ("no", "n", "false", "0"):
        return False
    return None


def split_sa_answer_tokens(answer_text, separator="<sa_ans_sep>"):
    return [token.strip() for token in str(answer_text).split(separator) if token.strip() != ""]


def parse_numeric_token(value):
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return None


def numeric_equal(left, right, rel_tol=1e-9, abs_tol=1e-9):
    return math.isclose(left, right, rel_tol=rel_tol, abs_tol=abs_tol)


def normalize_sa_text(value, case_sensitive):
    text = "" if pd.isna(value) else str(value).strip()
    return text if case_sensitive else text.lower()


def evaluate_sa_answer(
    candidate_entry,
    correct_answer,
    marks,
    response_type,
    evaluation_required,
    answer_type,
    case_sensitive,
):
    evaluation_required_value = parse_yes_no(evaluation_required)
    if evaluation_required_value is False:
        return 0, "M"

    candidate_text = "" if pd.isna(candidate_entry) else str(candidate_entry).strip()
    if candidate_text == "":
        return 0, "U"

    correct_answer_text = "" if pd.isna(correct_answer) else str(correct_answer).strip()
    if correct_answer_text == "":
        return 0, "W"

    response_type_text = "" if pd.isna(response_type) else str(response_type).strip().lower()
    answer_type_text = "" if pd.isna(answer_type) else str(answer_type).strip().lower()
    case_sensitive_value = parse_yes_no(case_sensitive) is True

    if response_type_text == "numeric":
        candidate_number = parse_numeric_token(candidate_text)
        if candidate_number is None:
            return 0, "W"

        if answer_type_text == "range":
            range_tokens = split_sa_answer_tokens(correct_answer_text)
            if len(range_tokens) != 2:
                return 0, "W"
            range_start = parse_numeric_token(range_tokens[0])
            range_end = parse_numeric_token(range_tokens[1])
            if range_start is None or range_end is None:
                return 0, "W"
            lower_bound = min(range_start, range_end)
            upper_bound = max(range_start, range_end)
            if lower_bound <= candidate_number <= upper_bound:
                return marks, "C"
            return 0, "W"

        if answer_type_text == "equal":
            expected_number = parse_numeric_token(correct_answer_text)
            if expected_number is None:
                return 0, "W"
            if numeric_equal(candidate_number, expected_number):
                return marks, "C"
            return 0, "W"

        if answer_type_text == "set":
            expected_tokens = split_sa_answer_tokens(correct_answer_text)
            expected_numbers = [parse_numeric_token(token) for token in expected_tokens]
            expected_numbers = [value for value in expected_numbers if value is not None]
            if len(expected_numbers) == 0:
                return 0, "W"
            if any(numeric_equal(candidate_number, expected_number) for expected_number in expected_numbers):
                return marks, "C"
            return 0, "W"

        return 0, "W"

    if response_type_text == "alphanumeric":
        candidate_normalized = normalize_sa_text(candidate_text, case_sensitive_value)

        if answer_type_text == "equal":
            expected_normalized = normalize_sa_text(correct_answer_text, case_sensitive_value)
            if candidate_normalized == expected_normalized:
                return marks, "C"
            return 0, "W"

        if answer_type_text == "set":
            expected_tokens = split_sa_answer_tokens(correct_answer_text)
            expected_normalized_tokens = [normalize_sa_text(token, case_sensitive_value) for token in expected_tokens]
            if candidate_normalized in expected_normalized_tokens:
                return marks, "C"
            return 0, "W"

        return 0, "W"

    return 0, "W"


def split_non_empty_csv(value, separator=","):
    return [token.strip() for token in str(value).split(separator) if token.strip() != ""]


def evaluate_correct_mark_and_result(
    question_type,
    selected_id,
    correct_option_id,
    marks,
    invalid_redirect=False,
    sa_response_type="",
    sa_evaluation_required="",
    sa_answer_type="",
    sa_case_sensitive="",
):
    if question_type == "SA":
        return evaluate_sa_answer(
            candidate_entry=selected_id,
            correct_answer=correct_option_id,
            marks=marks,
            response_type=sa_response_type,
            evaluation_required=sa_evaluation_required,
            answer_type=sa_answer_type,
            case_sensitive=sa_case_sensitive,
        )

    if question_type not in ("MCQ", "MSQ"):
        return "", ""

    if invalid_redirect:
        return 0, "W"

    selected_tokens = split_non_empty_csv(selected_id)
    correct_tokens = split_non_empty_csv(correct_option_id)

    # Keep placeholders for unanswered rows.
    if len(selected_tokens) == 0:
        return 0, "U"

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


def resolve_column_name(columns, candidates):
    normalized_column_map = {
        normalize_column_name(column): column
        for column in columns
    }

    normalized_candidates = [normalize_column_name(candidate) for candidate in candidates]

    for normalized_candidate in normalized_candidates:
        if normalized_candidate in normalized_column_map:
            return normalized_column_map[normalized_candidate]

    for normalized_candidate in normalized_candidates:
        for normalized_column, original_column in normalized_column_map.items():
            if normalized_column.startswith(normalized_candidate) or normalized_candidate.startswith(normalized_column):
                return original_column

    return None


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

        answer_question_ids = answer_df[question_id_column].apply(normalize_question_id_for_match)
        answer_values = answer_df[answer_column].apply(
            lambda value: "" if pd.isna(value) else str(value)
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

english_question_id_col = find_column(english_df.columns, ["Question ID", "Question id", "QuestionId"])
english_correct_option_col = find_column(english_df.columns, ["Correct Option", "CorrectOption"])
english_answer_col = find_column(
    english_df.columns,
    [
        "Answer(For SA)/Skeletal Code(For Programming Test)/Static text (For Typing Test)",
        "Answer(For SA)/\nSkeletal Code(For Programming Test)/\nStatic text (For Typing Test)",
        "Answer(For SA)",
    ],
)

if english_question_id_col is None:
    raise KeyError("Missing required column in English sheet: Question ID")
if english_correct_option_col is None:
    raise KeyError("Missing required column in English sheet: Correct Option")

# Build a Question ID -> Correct Option map from English sheet
english_question_ids = english_df[english_question_id_col].apply(normalize_question_id_for_match)
english_correct_options = english_df[english_correct_option_col].apply(parse_correct_option)
correct_option_map = dict(zip(english_question_ids, english_correct_options))

english_sa_answers = (
    english_df[english_answer_col].apply(normalize_sa_answer)
    if english_answer_col is not None
    else pd.Series([""] * len(english_df))
)
english_sa_answer_map = dict(zip(english_question_ids, english_sa_answers))

sa_response_type_column = resolve_column_name(
    filtered_df.columns,
    [
        "Response Type (For SA type of Questions)",
        "Response Type",
    ],
)
sa_evaluation_required_column = resolve_column_name(
    filtered_df.columns,
    [
        "Is Evaluation Required (For SA type of Questions)",
        "Is Evaluation Required",
    ],
)
sa_answer_type_column = resolve_column_name(
    filtered_df.columns,
    [
        "Answer type (For SA type of Questions)",
        "Answer Type",
    ],
)
sa_case_sensitive_column = resolve_column_name(
    filtered_df.columns,
    [
        "Answers case sensitive? (For SA type of Questions)",
        "Answers case sensitive",
    ],
)

empty_sa_series = pd.Series([""] * len(filtered_df), index=filtered_df.index)
sa_response_type_values = (
    filtered_df[sa_response_type_column].fillna("")
    if sa_response_type_column is not None
    else empty_sa_series
)
sa_evaluation_required_values = (
    filtered_df[sa_evaluation_required_column].fillna("")
    if sa_evaluation_required_column is not None
    else empty_sa_series
)
sa_answer_type_values = (
    filtered_df[sa_answer_type_column].fillna("")
    if sa_answer_type_column is not None
    else empty_sa_series
)
sa_case_sensitive_values = (
    filtered_df[sa_case_sensitive_column].fillna("")
    if sa_case_sensitive_column is not None
    else empty_sa_series
)

# Match using Question id and capture Correct Option (supports single and multi values)
config_question_ids = filtered_df[config_question_id_col].apply(normalize_question_id_for_match)
correct_option_values = config_question_ids.map(correct_option_map).fillna("")

sa_question_mask = filtered_df[config_question_type_col] == 'SA'
sa_question_ids = config_question_ids
sa_correct_answers = sa_question_ids.map(english_sa_answer_map).fillna("")

correct_option_values = correct_option_values.where(~sa_question_mask, sa_correct_answers)

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
    candidate_entry_values = config_question_ids.map(candidate_entry_map).fillna("").values

redirect_to_number_values = [
    candidate_entry_to_redirect(candidate_entry, question_type)
    for candidate_entry, question_type in zip(candidate_entry_values, filtered_df['Question Type'].values)
]
selected_id_and_invalid_flags = [
    redirect_to_selected_id(option_list, redirect_to_number, separator_value)
    for option_list, redirect_to_number in zip(concatenated_options, redirect_to_number_values)
]
selected_id_values = [item[0] for item in selected_id_and_invalid_flags]
invalid_redirect_values = [item[1] for item in selected_id_and_invalid_flags]

option_not_applicable = "Not Applicable"
option_1_values = option_1.mask(sa_question_mask, option_not_applicable).values
option_2_values = option_2.mask(sa_question_mask, option_not_applicable).values
option_3_values = option_3.mask(sa_question_mask, option_not_applicable).values
option_4_values = option_4.mask(sa_question_mask, option_not_applicable).values
concatenated_options_values = [
    option_not_applicable if is_sa else option_value
    for is_sa, option_value in zip(sa_question_mask, concatenated_options)
]
correct_option_id_values = [
    sa_correct_answer if is_sa else correct_option_id
    for is_sa, sa_correct_answer, correct_option_id in zip(
        sa_question_mask,
        sa_correct_answers,
        correct_option_ids,
    )
]
redirect_to_number_values = [
    option_not_applicable if is_sa else redirect_to_number
    for is_sa, redirect_to_number in zip(sa_question_mask, redirect_to_number_values)
]
selected_id_values = [
    candidate_entry if is_sa else selected_id
    for is_sa, candidate_entry, selected_id in zip(
        sa_question_mask,
        candidate_entry_values,
        selected_id_values,
    )
]
invalid_redirect_values = [
    False if is_sa else invalid_redirect
    for is_sa, invalid_redirect in zip(sa_question_mask, invalid_redirect_values)
]

# Marks from Configuration Details for the filtered question rows
marks_values = filtered_df[config_marks_col].apply(normalize_marks).values

# Evaluate awarded mark and result for MCQ/MSQ/SA using selected answer vs configured correct answer.
correct_mark_and_result = [
    evaluate_correct_mark_and_result(
        question_type,
        selected_id,
        correct_option_id,
        marks,
        invalid_redirect,
        sa_response_type,
        sa_evaluation_required,
        sa_answer_type,
        sa_case_sensitive,
    )
    for (
        question_type,
        selected_id,
        correct_option_id,
        marks,
        invalid_redirect,
        sa_response_type,
        sa_evaluation_required,
        sa_answer_type,
        sa_case_sensitive,
    ) in zip(
        filtered_df[config_question_type_col].values,
        selected_id_values,
        correct_option_id_values,
        marks_values,
        invalid_redirect_values,
        sa_response_type_values.values,
        sa_evaluation_required_values.values,
        sa_answer_type_values.values,
        sa_case_sensitive_values.values,
    )
]
correct_mark_values = [item[0] for item in correct_mark_and_result]
result_values = [item[1] for item in correct_mark_and_result]

# Create a new dataframe with S.No, Question id, Question Type, OPTION 1-4, separator, and concatenated options
output_df = pd.DataFrame({
    'S.No': range(1, len(filtered_df) + 1),
    'Question id': config_question_ids.values,
    'Question Type': filtered_df[config_question_type_col].values,
    'OPTION 1': option_1_values,
    'OPTION 2': option_2_values,
    'OPTION 3': option_3_values,
    'OPTION 4': option_4_values,
    'Separator': separator_value,
    'Concatenated Options': concatenated_options_values,
    'Correct Option': correct_option_values.values,
    'Correct Option Id': correct_option_id_values,
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