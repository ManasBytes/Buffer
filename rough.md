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
