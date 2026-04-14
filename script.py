import pandas as pd

config_file = "noc26-cs79_S4.xlsx"
output_file = "sample_question_answer_sheet.xlsx"


def normalize_name(value):
	return "".join(ch.lower() for ch in str(value).strip() if ch.isalnum())


def find_column(columns, candidates):
	normalized_map = {normalize_name(col): col for col in columns}
	for candidate in candidates:
		key = normalize_name(candidate)
		if key in normalized_map:
			return normalized_map[key]
	return None


all_sheets = pd.read_excel(config_file, sheet_name=None)

if "Configuration Details" not in all_sheets:
	raise KeyError("Sheet 'Configuration Details' not found in the workbook.")

config_df = all_sheets["Configuration Details"]

question_id_col = find_column(config_df.columns, ["Question id", "Question ID", "QuestionId"])
question_type_col = find_column(config_df.columns, ["Question Type", "QuestionType"])
marks_col = find_column(config_df.columns, ["Marks", "Mark"])

missing = []
if question_id_col is None:
	missing.append("Question id")
if question_type_col is None:
	missing.append("Question Type")
if marks_col is None:
	missing.append("Marks")
if missing:
	raise KeyError(f"Missing required column(s) in 'Configuration Details': {', '.join(missing)}")

# Keep only non-comprehension questions (case-insensitive match).
question_type_text = config_df[question_type_col].fillna("").astype(str)
filtered_df = config_df[~question_type_text.str.contains("comprehension", case=False, regex=False)].copy()

output_df = pd.DataFrame({
	"Question id": filtered_df[question_id_col],
	"S.No": range(1, len(filtered_df) + 1),
	"Question Type": filtered_df[question_type_col],
	"Marks": filtered_df[marks_col],
	"Enter your answer": "",
})

output_df.to_excel(output_file, index=False, sheet_name="Questions")

print(f"Created: {output_file}")
print(f"Rows written: {len(output_df)}")