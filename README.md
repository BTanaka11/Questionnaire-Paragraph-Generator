Project

This is a demo of a work project I did in early 2022, with all company information, data, and proprietary code removed and replaced with generic, non-identifiable information. This required renaming some variables/filenames, deleting some functions, and making up new data. Around 50% of the
code was removed, and the data files were reduced to just the essentials. It now represents a fictional
general store's customer questionnaires' conversion to paragraph form. This is different from the business problem the original addressed,
but the underlying concepts and algorithm work, which this project intends to highlight, are the same.

In both the original and this demo, I made all code and design decisions, though for some of the sentence wording I sought colleague feedback,
and someone else did the AWS deployment and assisted with testing.
I factored in many non-technical aspects when creating this, such as the existing manual writing process it was replacing, the role writers would play after
launch (lightly reviewing output and free to edit), how future questionnaire updates from a different upstream department would
requiring certain updates, the possibility of future questionnaire updates occurring without notice due to communication gaps, translation tool
(acting on upstream input) changing, etc.

Why useful
This project is a solution to the following problem:

There are a set of questionnaires, each with preset questions. Each question has a set of selection answers and only 1 can be picked. Any number of answers
within the set may represent a free text answer, which would be typed right after the selection answer.
Each questionnaire has 1 english version and may also have any number of foreign language versions.
The filled out questionnaires are auto-translated to english, and the translation process may "divergently" change a small amount of the text due to
surrounding text content or translator inconsistencies. EX: "...hola como estas..." could unpredictably translate to
"Hello how are you" or "Hi how are you?" or other unknown variants, but none of these variants would differ by more than a few letters.

Each questionnaire could have child questionnaires.

The translated questionnaires need to be transformed to grammatically correct english paragraphs. If there's any uncertainty on the input, such as new
unexpected questionnaires or question/selection-answer text difference that's more than a minor translation effect, it's better to not generate any
paragraph. If present, child questionnaires should be part of the generated text as a 2nd paragraph.

Summary of main steps:
1. class Callscript loads the questionnaires from LookupTable.xlsx and precompiles regex fuzzy text matchers for each question.
2. text_Generator function takes the Input.xlsx data and checks each record for a Callscript object match.
3. To determine a match, the Callscript object's Get_Answers function parses the input questionnaire text top-to-bottom, verifying the question and answers
match the lookup table expectations.
4. If no match, a useful description of what the discrepancy was is returned and will be the final output.
5. If all match, it counts as a match and extracted answers are returned.
6. Answers go through hard-coded q&a logic to generate the paragraph 1 sentence at a time.

How to run:
1. pip install regex
2. pip install DateTime
3. optionally do "pip install symspellpy"
4. pip install openpyxl
5. pip install pandas
6. pip install xlsxwriter
6. run the py file. Output will be saved as an Excel file and printed in terminal.
7. If you don't have Excel installed, pandas saving and reading still work fine but
   you can't as easily edit the input or lookuptable if you want to experiment.
6. If you have Excel installed, feel free to mess around with the input and/or lookup table between runs.
   Remember, the goal of this project was to both allow a small (8 chars or 8% of q/a length whichever is shorter) amount of
   spelling variation and have high security of not generating any paragraph text if there's any question or answer that's
   not in agreement between all 3 of: input, lookuptable, hard-coded questions/answers.