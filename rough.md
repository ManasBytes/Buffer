# This

To evaluate a short answer type question there are certain rules that are to be followed.

1. In the short answer type also there are many types of questions, Like there that you can get from the configuration details file.
I am explaining here what are all those types and what are we supposed to do in the script.

In the configuration sheet only, you will find a column named, "Response Type  (For SA type of Questions)", this will say either the answer for this question id is going to be "Alphanumeric" or "Numeric".

Then there is another column named "Is Evaluation Required (For SA type of Questions)", in this column there are two values, one is "Yes" or "No" (the yes and no can be not capitalized so .. like be careful). If it is yes, then we evaluate the answer for the question , and if it is no, then in the correct mark we put a zero and then in the result we put the value "M" for manual evaluation. So make it conditional that this is going to be the first check, if it is no then there is no point in finding if it is alphanumeric or numeric and all those sort of things.

Then there is the column "Answer type (For SA type of Questions)" . This column can have three values. "Range", or "Equal" or "Set".
Range -> This is going to be there only for numeric response type question. In this a range would be provided and baased on that you are going to evaluate whether the value entered by the student is correct or not, meaning the value entered by the student lies in the given range or not.

Equal -> This method is going to be there for the numeric and the alphanumeric response types and this going to equate this. So the entered answer is going to exactly what the given answer is.
(case sensitive or insensitive is described below)

Set -> This method is going to be there for the numeric and the alphanumeric response types and this is going to have more than one answer for this. The provided answer is going to be like... suppose MLF is a answer to some question then some people can write MLF as Machine Learning Foundations that is also correct or Machine Learning Foundation that is also correct. or say for a question the answer can be 12 or 13 or 14, so for these type of sa questions this set type is used.

There is another column of importance which is the "Answers case sensitive? (For SA type of Questions)". This is going to be self explanatory, for case sensitive yes, you have to compare the alphanumeric exactly as it is with the answer... and with case sensitive no, you would convert the answers and the candidate entered thing into lowercase and then compare and evaluate them.

The answer in the proper syntax is mentioned in the row "Answer(For SA)/Skeletal Code(For Programming Test)/Static text (For Typing Test)". In this row you are going to encounter some syntaxical code which  i am mentioning below.

1. For Separation : Anywhere where separation is required there this is entered "&lt;sa_ans_sep&gt;" i dont know how md is going to take this i want to say this <sa_ans_sep> this is the separator used.
Where is it used ?
    a. For describing a range -> 6.23<sa_ans_sep>7.23   means -> [6.23,7.23] both ends included.
    b. For describing a set  -> mango<sa_ans_sep>banana<sa_ans_sep>apple   means -> answer can be     mango or banana or apple. Set can be used to describe numberd output also like  12 of 15 or 18 would be written here as 12<sa_ans_sep>15<sa_ans_sep>18. Done.

## Main Script

I want you to make a script for something really important. There is a folder here called Main and in the Main Folder there is going to be Folders with dates 28 March 2026 , 29 March 2026.... and many more ... meaning in this folder there is going to be like folders like in the dates in which the exams will be conducted all the dates would be there. So it is going to be a list.

Now after that inside that folder there is going to be shift 1 and shift 2 folder.

Inside that folder there is going to be excel sheets. for each exam. for each of the excel sheet here we have to create one excel sheet as the output.

-- To Structure the outputs.
To structure the outputs we would be creating another folder as output in the same directory as main foldder and then inside that folder we would create the same date folder as earlier we had accessed and also in side that same folder we wouold create the same shift thhat we accessed and inside that we would create the file named output_file_name ... what was their earlier

for example the file was here
main/28_March_2026/Shift_1/ae_2026.xlsx

the output file will be at
output/28_March_2026/Shift_1/ae_2026.xlsx

-- What are we going to process ?
the file we are going to take we are going to run this script, for each of those files.
import pandas as pd

config_file = "noc26-cs79_S4.xlsx"
output_file = "sam.xlsx"


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

### Keep only non-comprehension questions (case-insensitive match).
question_type_text = config_df[question_type_col].fillna("").astype(str)
filtered_df = config_df[~question_type_text.str.contains("comprehension", case=False, regex=False)].copy()

output_df = pd.DataFrame({
	"Question id": filtered_df[question_id_col].apply(lambda value: "" if pd.isna(value) else str(value).strip()),
	"S.No": range(1, len(filtered_df) + 1),
	"Question Type": filtered_df[question_type_col],
	"Marks": filtered_df[marks_col],
	"Enter your answer": "",
})

output_df.to_excel(output_file, index=False, sheet_name="Questions")

print(f"Created: {output_file}")
print(f"Rows written: {len(output_df)}")


this script will be ran with some addiotional changes... the sheet in which these details are there... question id and number and all those things... thaat sheet would be named Question Paper Detail
and another sheet would be created with the name "basic details"
in that sheet these details would be there... 
Name :
DOB :
Roll No :
Subject : the file name like in our earlier example it was ae_2026.xlsx.. so the subject field should have ae_2026 ... 
Exam Date :
Session :



make it 




# This again

the script there in this main.py is just for one file. I want this to be iterating over a number of files. 
The folder structure is going to be like this ... there would be a variable within which i would give the name of the main file... the main file from which input is going to be taken.
Now inside that main file there is going to be a list of files for each subject.
main/ 
	subject1
	subject2
	subject3
	subject4


	i dont even knnow what name is going to be there... and inside that suubject this is going to be there.

main/ 
	subject1/
		examplename.xlsx
		candidates(dirctory)
	subject2/
		examplename.xlsx
		candidates(directory)
	subject3/
		examplename.xlsx
		candidates(directory)

like this. 
So this is going to be my folder sstructure for the script.

now the script that is their in main.py .. in that the config file is going to be examplename.xlsx (this is going to be a short form of the subject code only)
and the answer sheet file is going to be, each file in the candidates directory is going to be the answersheet path. 

For each answersheet path, one new sheet in one ooutput.xlsx is going to be there.
so each file in the candidates directory is going to have the answersheeet of one student, in the basic details sheet of that thing thee detials of the student is going to be there. and then in each sheet of output.xlsx a unique student would be there


so the output structure is going to be like this
main/ 
	subject1/
		examplename.xlsx   
		candidates(dirctory)/
			1.xlsx
			2.xlsx
	subject2/
		examplename.xlsx
		candidates(directory)
	subject3/
		examplename.xlsx
		candidates(directory)
(for this above input file structure the output structure i expecvt is )

output/
	subject1/
		examplename.xlsx
	subject2/
		examplename.xlsx
	subject3/
		examplename.xlsx


see what i did i maintained the name structure as it is in the input file structure. 

so i want the names to be just copied but to the output folder which is oging too be in the same directory as in the main folder .


just this.


there is a change in the answersheet .... in the answersheet thee are now two sheets, first one is the basic detailss sheet in which the student is going to fill his basic details and then another is going to be the question paper details sheet in which  the student is going to aanswer the questiion. so make the script to like fetch the answers from the question  paper details sheet and. another thing is . 


in the output sheet after the result column ...like in the last column, create another column as basic details... and from the basic detaails sheet of that answer sheet copy the Name , DOB, Roll No and print it like row wise...

like Name then in the next row  the value of his name... which is the structure in which the basic details are there iin the given sheet.