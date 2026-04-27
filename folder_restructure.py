import argparse
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd


MAX_WORKERS = 8


def normalize_key(value):
	if value is None:
		return ""
	return "".join(ch.lower() for ch in str(value).strip() if ch.isalnum())


def normalize_subject(value):
	if value is None:
		return ""
	return " ".join(str(value).strip().split()).lower()


def find_name_match(names, candidates):
	normalized_map = {normalize_key(name): name for name in names}
	normalized_candidates = [normalize_key(candidate) for candidate in candidates]

	for item in normalized_candidates:
		if item in normalized_map:
			return normalized_map[item]

	for item in normalized_candidates:
		for normalized_name, actual_name in normalized_map.items():
			if normalized_name.startswith(item) or item.startswith(normalized_name):
				return actual_name

	return None


def excel_files_in(folder_path):
	if not folder_path.exists() or not folder_path.is_dir():
		return []
	files = [
		file_path
		for file_path in folder_path.iterdir()
		if file_path.is_file()
		and file_path.suffix.lower() in {".xlsx", ".xls"}
		and not file_path.name.startswith("~$")
	]
	files.sort(key=lambda item: item.name.lower())
	return files


def compute_worker_count(item_count):
	if item_count <= 0:
		return 1
	return min(MAX_WORKERS, item_count)


def read_config_subject(config_file_path):
	all_sheets = pd.read_excel(config_file_path, sheet_name=None)
	config_sheet_name = find_name_match(
		list(all_sheets.keys()),
		["Configuration Details", "ConfigurationDetails"],
	)
	if config_sheet_name is None:
		return ""

	config_df = all_sheets[config_sheet_name]
	subject_column = find_name_match(list(config_df.columns), ["Subject"]) 
	if subject_column is None:
		return ""

	subject_series = config_df[subject_column].dropna().astype(str).map(str.strip)
	subject_series = subject_series[subject_series != ""]
	if len(subject_series) == 0:
		return ""

	return subject_series.iloc[0]


def read_candidate_subject(candidate_file_path):
	all_sheets = pd.read_excel(candidate_file_path, sheet_name=None, header=None)
	basic_details_sheet_name = find_name_match(
		list(all_sheets.keys()),
		["Basic Details", "BasicDetails"],
	)
	if basic_details_sheet_name is None:
		return ""

	basic_df = all_sheets[basic_details_sheet_name]
	row_count, col_count = basic_df.shape

	for row_index in range(row_count):
		for col_index in range(col_count):
			cell_value = basic_df.iat[row_index, col_index]
			if normalize_key(cell_value) != "subject":
				continue

			# Prefer same-row value on the right.
			for next_col_index in range(col_index + 1, col_count):
				candidate_value = basic_df.iat[row_index, next_col_index]
				candidate_text = "" if pd.isna(candidate_value) else str(candidate_value).strip()
				if candidate_text != "":
					return candidate_text

			# Fallback to next row value.
			next_row_index = row_index + 1
			if next_row_index < row_count:
				for any_col_index in range(col_count):
					candidate_value = basic_df.iat[next_row_index, any_col_index]
					candidate_text = "" if pd.isna(candidate_value) else str(candidate_value).strip()
					if candidate_text != "":
						return candidate_text

			return ""

	return ""


def build_combined_structure(metadata_template_dir, combined_dir):
	config_files = excel_files_in(metadata_template_dir)
	if len(config_files) == 0:
		raise FileNotFoundError(f"No config files found in metadata template folder: {metadata_template_dir}")

	subject_to_destinations = {}
	duplicate_subjects = []
	errored_configs = []

	def process_single_config(config_file):
		subject_folder = combined_dir / config_file.stem
		config_folder = subject_folder / "config_file"
		candidates_folder = subject_folder / "candidates"

		config_folder.mkdir(parents=True, exist_ok=True)
		candidates_folder.mkdir(parents=True, exist_ok=True)

		shutil.copy2(config_file, config_folder / config_file.name)

		subject_name = read_config_subject(config_file)
		normalized_subject = normalize_subject(subject_name)
		return {
			"config_name": config_file.name,
			"subject_name": subject_name,
			"normalized_subject": normalized_subject,
			"candidates_folder": candidates_folder,
		}

	worker_count = compute_worker_count(len(config_files))
	with ThreadPoolExecutor(max_workers=worker_count) as executor:
		future_map = {
			executor.submit(process_single_config, config_file): config_file
			for config_file in config_files
		}

		for future in as_completed(future_map):
			config_file = future_map[future]
			try:
				result = future.result()
			except Exception as exc:  # pylint: disable=broad-except
				errored_configs.append((config_file.name, str(exc)))
				continue

			normalized_subject = result["normalized_subject"]
			if normalized_subject == "":
				continue

			if normalized_subject in subject_to_destinations:
				duplicate_subjects.append((result["subject_name"], result["config_name"]))

			subject_to_destinations.setdefault(normalized_subject, []).append(result["candidates_folder"])

	return subject_to_destinations, config_files, duplicate_subjects, errored_configs


def route_candidates(candidate_entries_dir, subject_to_destinations):
	candidate_files = excel_files_in(candidate_entries_dir)
	copied_count = 0
	unmatched_candidates = []
	errored_candidates = []

	def process_single_candidate(candidate_file):
		candidate_subject = read_candidate_subject(candidate_file)
		normalized_subject = normalize_subject(candidate_subject)

		if normalized_subject == "" or normalized_subject not in subject_to_destinations:
			return {
				"status": "unmatched",
				"candidate_name": candidate_file.name,
				"candidate_subject": candidate_subject,
				"copied_count": 0,
			}

		destination_folders = subject_to_destinations[normalized_subject]
		copy_count = 0
		for destination_folder in destination_folders:
			shutil.copy2(candidate_file, destination_folder / candidate_file.name)
			copy_count += 1

		return {
			"status": "copied",
			"candidate_name": candidate_file.name,
			"candidate_subject": candidate_subject,
			"copied_count": copy_count,
		}

	worker_count = compute_worker_count(len(candidate_files))
	with ThreadPoolExecutor(max_workers=worker_count) as executor:
		future_map = {
			executor.submit(process_single_candidate, candidate_file): candidate_file
			for candidate_file in candidate_files
		}

		for future in as_completed(future_map):
			candidate_file = future_map[future]
			try:
				result = future.result()
			except Exception as exc:  # pylint: disable=broad-except
				errored_candidates.append((candidate_file.name, str(exc)))
				continue

			if result["status"] == "unmatched":
				unmatched_candidates.append((result["candidate_name"], result["candidate_subject"]))
			else:
				copied_count += result["copied_count"]

	return candidate_files, copied_count, unmatched_candidates, errored_candidates


def parse_args():
	parser = argparse.ArgumentParser(
		description=(
			"Restructure candidate and metadata folders into a combined folder grouped by config files and subject match."
		)
	)
	parser.add_argument(
		"--candidate-entries",
		default="candidate entries",
		help="Path to the candidate entries directory (default: 'candidate entries').",
	)
	parser.add_argument(
		"--metadata-template",
		default="metadata template",
		help="Path to the metadata template directory (default: 'metadata template').",
	)
	parser.add_argument(
		"--combined",
		default="combined",
		help="Output folder path to create the combined structure (default: 'combined').",
	)
	return parser.parse_args()


def main():
	args = parse_args()

	candidate_entries_dir = Path(args.candidate_entries)
	metadata_template_dir = Path(args.metadata_template)
	combined_dir = Path(args.combined)

	if not candidate_entries_dir.exists() or not candidate_entries_dir.is_dir():
		raise FileNotFoundError(f"Candidate entries folder not found: {candidate_entries_dir}")
	if not metadata_template_dir.exists() or not metadata_template_dir.is_dir():
		raise FileNotFoundError(f"Metadata template folder not found: {metadata_template_dir}")

	combined_dir.mkdir(parents=True, exist_ok=True)

	subject_to_destinations, config_files, duplicate_subjects, errored_configs = build_combined_structure(
		metadata_template_dir,
		combined_dir,
	)
	candidate_files, copied_count, unmatched_candidates, errored_candidates = route_candidates(
		candidate_entries_dir,
		subject_to_destinations,
	)

	print("Folder restructure complete")
	print(f"Config files discovered: {len(config_files)}")
	print(f"Candidate files discovered: {len(candidate_files)}")
	print(f"Candidate copies created: {copied_count}")
	print(f"Unmatched candidates: {len(unmatched_candidates)}")
	print(f"Config read/copy errors: {len(errored_configs)}")
	print(f"Candidate read/copy errors: {len(errored_candidates)}")

	if len(errored_configs) > 0:
		print("\nConfigs with processing errors:")
		for config_name, error_message in errored_configs:
			print(f"- {config_name}: {error_message}")

	if len(duplicate_subjects) > 0:
		print("\nWarning: duplicate subject values found in config files.")
		for subject_name, config_name in duplicate_subjects:
			print(f"- Subject '{subject_name}' repeated in config file: {config_name}")

	if len(unmatched_candidates) > 0:
		print("\nCandidates with missing/unmatched subject:")
		for candidate_name, candidate_subject in unmatched_candidates:
			display_subject = candidate_subject if str(candidate_subject).strip() != "" else "<missing>"
			print(f"- {candidate_name}: subject={display_subject}")

	if len(errored_candidates) > 0:
		print("\nCandidates with processing errors:")
		for candidate_name, error_message in errored_candidates:
			print(f"- {candidate_name}: {error_message}")


if __name__ == "__main__":
	main()
