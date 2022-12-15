## Summary
This is a demo of a work project I did in early 2022, with all company information, data, and proprietary code replaced with generic, non-identifiable information. This required renaming some variables/filenames, deleting some functions, and making up new data. Around 50% of the code was removed, and the data files were reduced to just the essentials. It now represents a fictional general store's customer questionnaires, which require conversion to paragraphs. This is different from the business problem the original addressed, but the underlying concepts and algorithm work, which this demo intends to highlight, are the same.

## Problem Background, Constraints, and Input Description
There are a set of questionnaires, each with preset questions. Each question has a set of selection answers and only 1 can be picked. Any number of answers within the set may represent a free text answer, which would be typed right after the selection answer. Each questionnaire has 1 English version and may also have any number of foreign language versions. The filled out questionnaires are auto-translated to English, and the translation process may divergently change a small amount of the text due to surrounding text content or translator inconsistencies. EX: "...hola como estas..." could unpredictably translate to "Hello how are you", "Hi how are you?", or other unknown variants, but none of these variants would differ by more than a few letters. Each questionnaire could have child questionnaires.

## Problem Statement
The translated questionnaires need to be transformed to grammatically correct English paragraphs. If there's any uncertainty on the input, such as new/unexpected questionnaires or question/selection-answer differences more than a minor translation effect, it's better to not generate any
paragraph. If present, child questionnaires should be part of the generated text as a 2nd paragraph.

## Languages and Tools
<p align="left"> <a href="https://www.python.org" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/master/icons/python/python-original.svg" alt="python" width="40" height="40"/> </a><a href="https://pandas.pydata.org/" target="_blank" rel="noreferrer"> <img src="https://raw.githubusercontent.com/devicons/devicon/2ae2a900d2f041da66e950e4d48052658d850630/icons/pandas/pandas-original.svg" alt="pandas" width="40" height="40"/> </a>  </p>

## What the Code Does
1. class ***Callscript*** loads the questionnaires from ***LookupTable.xlsx*** and precompiles regex fuzzy text matchers for each question
2. ***text_Generator*** function takes **Input.xlsx** and checks each record for a Callscript object match
3. To determine a match, the Callscript object's ***Get_Answers*** function parses the input questionnaire text top-to-bottom, verifying the question and answers match the lookup table expectations
4. If no match, a useful description of what the discrepancy was is returned and will be the final output
5. If all match, it counts as a match and extracted answers are returned
6. Answers go through hard-coded q&a logic to generate the paragraph 1 sentence at a time

https://user-images.githubusercontent.com/37204126/207978125-cc8286b1-bc01-4e17-8903-906974508d84.mp4

## How to Run
1. pip install regex
2. pip install DateTime
3. optionally do "pip install symspellpy"
4. pip install openpyxl
5. pip install pandas
6. pip install xlsxwriter
6. run the py file. Output will be saved as an Excel file and printed in terminal.
7. Feel free to mess around with the input and/or lookup table between runs. Remember, the goal of this project was to both allow a small (8 chars or 8% of q/a length whichever is shorter) amount of spelling variation and have high security of not generating any paragraph text if there's any question or answer that's not in agreement between all 3 of: input, lookuptable, hard-coded questions/answers.
