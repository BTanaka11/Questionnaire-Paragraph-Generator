import json
import os
import regex
import sys
import time
import datetime as dt
from datetime import date
import glob
import pandas as pd

try:
    from symspellpy import SymSpell, Verbosity
    import pkg_resources
    Correct_Spelling_On = True
    sym_spell = SymSpell(max_dictionary_edit_distance=1, prefix_length=7)
    dictionary_path = pkg_resources.resource_filename(
        "symspellpy", "frequency_dictionary_en_82_765.txt")
    sym_spell.load_dictionary(dictionary_path, term_index=0, count_index=1)
except:
    Correct_Spelling_On = False

lookup_table_filepath = 'LookupTable.xlsx'
input_filepath = 'Input.xlsx'
output_folder_path = 'Output/'

def get_fuzzieness():
    fuzzy = os.environ.get('FUZZIENESS')
    if fuzzy is None:
        fuzzy = 8
    return fuzzy

def Capitalize_Sentences(inp):
    def Uppercase_group(m):
        return m.group(1) + m.group(2).upper()

    return regex.sub(r'([.!?] *)([a-z])', Uppercase_group, inp)


def Correct_Spelling(inp):
    inp = regex.sub(r'([a-zA-Z0-9] *)\n+', r'\1. ', inp)
    if Correct_Spelling_On == True:
        c = []

        def Correct_formatting(inp):
            inp = regex.sub(r' (!|\.|\?|,|\))', r'\1', inp)
            inp = regex.sub(r'(!|\.|\?|,|\()(\w)', r'\1 \2', inp)
            inp = regex.sub(r'\( ', '(', inp)
            return ' '.join(inp.split())

        for word in inp.split():
            if word.isupper() and len(word) > 1:
                c.append(word.lower())
            elif word[0].islower() and word.isalpha():
                candidates = sym_spell.lookup(word, Verbosity.CLOSEST, max_edit_distance=1, include_unknown=True)
                if len(candidates) == 1:
                    c.append(candidates[0].term)

                else:
                    c.append(word)
            else:
                c.append(word)

        return regex.sub(r" ?(' ?|’ ?|:)", r'\1', Correct_formatting(' '.join(c)))
    else:
        return inp

def Update_Dict(D1, D2):
    for key, value in D2.items():
        if key not in D1.keys():
            D1[key] = value
        else:
            Nulls = ['(BLANK-FREE TEXT OPTION PICKED)', '(BLANK-NOTHING AT ALL)', 'Unknown', 'Not Applicable',
                     'Not Specified', 'Not Available']
            if value not in Nulls and D1[key] in Nulls:
                D1[key] = value
    return D1

class Callscript:
    regex_compiled_final_text = regex.compile(r'(^(?:' + regex.escape('***') + '))', flags=regex.MULTILINE | regex.DOTALL | regex.IGNORECASE | regex.WORD)

    def __init__(self, Tab_Name, *Extra_Callscripts):
        self.DF = pd.read_excel(lookup_table_filepath, sheet_name=Tab_Name, keep_default_na=False, dtype={'#': int})
        self.DF['Question'] = self.DF.apply(lambda row: row['Question'].replace('\'', '\"'), axis=1)
        self.DF['Question'] = self.DF.apply(lambda row: regex.sub(r' *- *', '-', row['Question']), axis=1)
        self.DF['Question'] = self.DF.apply(lambda row: row['Question'].replace('“', '"').replace('”', '"'), axis=1)
        self.max_Number = self.DF['#'].max()
        self.min_Number = self.DF['#'].min()
        self.compiled_all_lang = {}
        self.product_line = Tab_Name
        for Q_type in self.DF['Question Type'].drop_duplicates():
            Q_length_calc = [len(Q) for Q in self.DF[self.DF['Question Type'] == Q_type]['Question'].to_list()]
            Q_length_calc = (int(sum(Q_length_calc) / len(Q_length_calc) * .08))
            if Q_length_calc > get_fuzzieness():
                Q_length_calc = str(get_fuzzieness())
            else:
                Q_length_calc = str(Q_length_calc)

            self.compiled_all_lang[Q_type] = regex.compile(r'' + '|'.join(
                ['^(?:' + regex.escape(Q) + '){e<=' + Q_length_calc + ':[^\n]} *\n? ?: ?' for Q in
                 self.DF[self.DF['Question Type'] == Q_type]['Question'].to_list()]),
                                                           flags=regex.MULTILINE | regex.DOTALL | regex.IGNORECASE)

        for col in [col for col in self.DF.columns if 'Response' in col]:
            self.DF[col] = self.DF.apply(lambda row: regex.sub(r' ?- ?', '-', row[col]), axis=1)

        self.Extra_Callscripts = [Callscript(x) for x in Extra_Callscripts]
        self.DF['Free Text Columns'] = self.DF.apply(lambda row: row['Free Text Columns'].split(','), axis=1)

    def Get_Answers(self, AQ, iloc_to_use=0):
        Dict = {}
        AQ_Temp = AQ
        Miss_Type = ''
        for k in range(1, self.max_Number + 2):
            if k != self.max_Number + 1:
                Fuzzy_Match = self.compiled_all_lang[
                    self.DF[self.DF['#'] == k]['Question Type'].iloc[iloc_to_use]].split(AQ_Temp, maxsplit=1)
            else:
                Fuzzy_Match = [AQ_Temp, 1]
            if len(Fuzzy_Match) == 2:
                AQ_Temp = Fuzzy_Match[1]
                if k == 1:
                    regex_final_text = self.regex_compiled_final_text.split(AQ_Temp, maxsplit=1)
                    AQ_Temp = regex_final_text[0]
                    AQ_Surrounding = Fuzzy_Match[0] + ''.join(regex_final_text[1:])
            else:
                Dict.clear()
                if k > 1:
                    Miss_Type = self.DF[self.DF['#'] == k]['Question Type'].iloc[
                                    iloc_to_use] + ' - CALLSCRIPT QUESTION NOT DETECTED!!!'
                break
            if k > 1:
                Answer = Fuzzy_Match[0].strip()
                Answer = regex.sub(r' ?- ?', '-', Answer)
                Free_text_Answer_Found = ''
                Selection_answer_found = False
                for col in [col for col in self.DF[self.DF['#'] == k - 1]['Free Text Columns'].iloc[iloc_to_use] if
                            col]:
                    for Answer_row in self.DF[(self.DF['#'] == k - 1) & (self.DF[col])][col].to_list():
                        answer_match = regex.match(r'^(?:' + regex.escape(Answer_row) + '){e<=' + str(
                            int(len(Answer_row) * .08)) + ':[^\n]} ?\n?(.*)', Answer,
                                                   flags=regex.MULTILINE | regex.DOTALL | regex.IGNORECASE)
                        if answer_match:
                            Answer = answer_match.groups()[0]
                            Free_text_Answer_Found = 'Free Text'
                            break

                    if Free_text_Answer_Found:
                        break

                for col in [col for col in self.DF.columns if
                            'Response' in col and col not in self.DF[self.DF['#'] == k - 1]['Free Text Columns'].iloc[
                                iloc_to_use]]:
                    for Answer_row in [x for x in self.DF[self.DF['#'] == k - 1][col].to_list() if x]:
                        if Answer_row == '(blank)':
                            if not Answer.strip():
                                Selection_answer_found = True
                                if Free_text_Answer_Found:
                                    Free_text_Answer_Found = ''
                                    Answer = '(BLANK-FREE TEXT OPTION PICKED)'
                                else:
                                    Answer = '(BLANK-NOTHING AT ALL)'
                                break

                        if regex.match(r'^(?:' + regex.escape(Answer_row) + '){e<=' + str(
                                int(len(Answer_row) * .08)) + ':[^\n]} ?$', Answer,
                                       flags=regex.MULTILINE | regex.DOTALL | regex.IGNORECASE):
                            Answer = self.DF[(self.DF['Language Translated From'] == 'English') & (self.DF['#'] == k - 1)][col].iloc[
                                iloc_to_use]
                            Selection_answer_found = True
                            Free_text_Answer_Found = ''
                            break
                    if Selection_answer_found:
                        break

                if Free_text_Answer_Found and Answer.strip():
                    Dict[self.DF[self.DF['#'] == k - 1]['Question Type'].iloc[iloc_to_use]] = [Free_text_Answer_Found,
                                                                                               Correct_Spelling(Answer)]

                elif Selection_answer_found:
                    Dict[self.DF[self.DF['#'] == k - 1]['Question Type'].iloc[iloc_to_use]] = Answer
                else:
                    Dict.clear()
                    Miss_Type = self.DF[self.DF['#'] == k - 1]['Question Type'].iloc[
                                    iloc_to_use] + ' - CALLSCRIPT ANSWER NOT DETECTED!!!'
                    break
        if Dict:
            for x in self.Extra_Callscripts:
                Extra_Dict, Extra_Miss_Type, AQ_Surrounding = x.Get_Answers(AQ_Surrounding)
                Dict = Update_Dict(Dict, Extra_Dict)

                if Extra_Miss_Type:
                    Dict['Extra Callscript Miss'] = Dict.pop('Extra Callscript Miss', []) + [
                        x.product_line + ' - ' + Extra_Miss_Type]

            return Dict, Miss_Type, AQ_Surrounding

        else:
            return {}, Miss_Type, AQ

TV_CS = Callscript('TV', 'TV_Dead_Pixel', 'TV_Cracked_Screen')
Dog_Food_CS = Callscript('Dog_Food')
Car_Tire_CS = Callscript('Car_Tire', 'Car_Tire_Leak')

Acronyms = pd.read_excel(lookup_table_filepath, sheet_name='Acronyms', keep_default_na=False)
Acronyms = Acronyms.to_dict(orient='index')

Word_Swaps = pd.read_excel(lookup_table_filepath, sheet_name='Word_Swaps', keep_default_na=False)
Word_Swaps = Word_Swaps.to_dict(orient='index')

def acronym_Correction(inp):
    for val in Word_Swaps.values():
        inp = regex.sub(r"\b{}\b".format(val['Before']), val['After'], inp, flags=regex.IGNORECASE)

    for val in Acronyms.values():
        inp = regex.sub(r"\b{} ?\({}\)".format(val['Full'], val['Acronym']), val['Acronym'], inp,
                        flags=regex.IGNORECASE)
        inp = regex.sub(r"\b{}\b".format(val['Full']), val['Acronym'], inp, flags=regex.IGNORECASE)
        inp = regex.sub(r"\b{}\b".format(val['Acronym']), val['Acronym'].upper(), inp, flags=regex.IGNORECASE)

    for val in Acronyms.values():
        C = len(regex.findall(r"\b{}\b".format(val['Acronym']), inp, flags=regex.IGNORECASE))
        if C == 1:
            inp = regex.sub(r"\b{}\b".format(val['Acronym']), val['Full'], inp, flags=regex.IGNORECASE)
        elif C > 1:
            inp = regex.sub(r"\b{}\b".format(val['Acronym']), val['Full'] + ' (' + val['Acronym'] + ')', inp, count=1,
                            flags=regex.IGNORECASE)

    return inp

def add_period_to_end(inp):
    if inp[-1:] not in ['.', '?', '!', ';', ':']:
        inp = inp + '.'
    return inp

def remove_double_words(inp1, inp2):
    if inp1.split()[-1] == inp2.split()[0]:
        return inp1 + ' '.join(inp2.split()[1:])
    return inp1 + inp2

def text_Generator(x):
    Questionnaire_Text = x['Questionnaire Text'].replace('\r', '\n').replace('&amp;',
                                                                                         'and').replace(
        '&quot;', '"').replace('“', '"').replace('”', '"').replace('—', '-').replace('{',
                                                                                     '(').replace('}',
                                                                                                  ')')
    Questionnaire_Text = regex.sub(r'\n+', '\n', Questionnaire_Text)
    Questionnaire_Text = regex.sub(r' *- *', '-', Questionnaire_Text)

    ED = ''

    for CS in [TV_CS, Dog_Food_CS, Car_Tire_CS]:
        Answers, Miss_Type = CS.Get_Answers(Questionnaire_Text)[0:2]

        if Answers:
            if 'UNK' in str(x['Product Code']):
                Product_Code = 'not provided'
            else:
                Product_Code = str(x['Product Code'])
            if 'UNK' in str(x['Product Batch#']):
                Serial_Number = 'not provided'
            else:
                Serial_Number = str(x['Product Batch#'])

            Product_Family = str(x['Product Family']).title()
            if Product_Family[0] in ['E', 'U', 'I', 'O', 'A']:
                Product_Family = 'an ' + Product_Family
            else:
                Product_Family = 'a ' + Product_Family

            Reporter_Type_Occupation = x['Reporter Type'].lower()

            try:
                Complaint_Receipt_Date = dt.datetime.strptime(x['Complaint Receipt Date'],
                                                              '%Y-%m-%d').strftime('%d-%b-%Y')
            except:
                Complaint_Receipt_Date = dt.datetime.strptime(str(x['Complaint Receipt Date'].date()),
                                                              '%Y-%m-%d').strftime('%d-%b-%Y')

            if str(x['Date of Event']).upper() in ['UNKNOWN', '01/01/0001', '01/01/1900', '01/01/1960']:
                Occurrence_Date = 'unknown'
            else:
                try:
                    Occurrence_Date = dt.datetime.strptime(x['Date of Event'], '%m/%d/%Y').strftime('%d-%b-%Y')
                except:
                    Occurrence_Date = str(x['Date of Event'])

            if CS is TV_CS:

                ED += f'As reported to our company on {Complaint_Receipt_Date} by a {Reporter_Type_Occupation}, there was an issue with {Product_Family} (Product code: {Product_Code} and Batch number: {Serial_Number}).'

                if Answers['Defect'][0] == 'Free Text':
                    ED += ' The issue was: ' + add_period_to_end(
                        Answers['Defect'][1])

                if Answers['Process Step'][0] == 'Free Text':
                    ED += ' The process step during which this occurred was described as: ' + add_period_to_end(
                        Answers['Process Step'][1])
                elif Answers['Process Step'] == 'Unknown':
                    ED += ' The process step during which this occurred was not specified.'
                elif Answers['Process Step'] == 'During setup':
                    ED += ' This occurred during setup.'
                elif Answers['Process Step'] == 'During transportation':
                    ED += ' This occurred during transportation.'
                elif Answers['Process Step'] == 'While watching':
                    ED += ' This occurred while watching.'
                elif Answers['Process Step'] == 'Out of Box':
                    ED += ' This event was an out-of-box-failure.'

                if Answers['Other Products'][0] == 'Free Text':
                    ED += ' Other product (' + Answers['Other Products'][1] + ') was associated with this event.'
                elif Answers['Other Products'] == 'No':
                    ED += ' No other product was associated with this event.'
                elif Answers['Other Products'] == 'Unknown':
                    ED += ' It was unknown if other product was associated with this event.'

                if Answers['Component'][0] == 'Free Text':
                    ED += ' The component involved was described as: ' + add_period_to_end(
                        Answers['Component'][1])
                elif Answers['Component'] == 'Screen':
                    ED += ' The component at fault was the screen.'
                elif Answers['Component'] == 'Screen Supporter':
                    ED += ' The component at fault was the screen stand.'
                elif Answers['Component'] == 'Cable Connector':
                    ED += ' The component at fault was a cable connector.'
                elif Answers['Component'] == 'Software':
                    ED += ' The TV software was at fault.'
                elif Answers['Component'] == 'Remote Control':
                    ED += ' The component at fault was the remote controller.'

                if 'Dead Pixel years used' in Answers.keys():
                    ED += '\n\nAdditional information was provided at the time of the report.'

                    if Answers['Dead Pixel years used'] == 'less than 1 year':
                        ED += ' The TV was used less than a year before this was noticed.'
                    elif Answers['Dead Pixel years used'] == 'between 1 and 3 years':
                        ED += ' The TV was used between 1 and 3 years before this was noticed.'
                    elif Answers['Dead Pixel years used'] == 'over 3 years':
                        ED += ' The TV was used between over 3 years before this was noticed.'
                    elif Answers['Dead Pixel years used'] == 'unknown':
                        ED += ' It was unknown how long the TV was in use before this was noticed.'

                    if Answers['Dead Pixel hours used'] == 'less than 4':
                        ED += ' The reporter claimed less than 4 hours per day average of use.'
                    elif Answers['Dead Pixel hours used'] == 'between 4 and 8':
                        ED += ' The reporter claimed 4-8 hours per day average of use.'
                    elif Answers['Dead Pixel hours used'] == 'over 8':
                        ED += ' The reporter claimed over 8 hours per day average of use.'
                    elif Answers['Dead Pixel hours used'] == 'unknown':
                        ED += ' The reporter did not know how many hours per day on average TV was used.'

                    if Answers['Dead Pixel color'][0] == 'Free Text':
                        ED += ' The color of light the dead pixel emmitted was: ' + add_period_to_end(
                            Answers['Dead Pixel color'][1])
                    elif Answers['Dead Pixel color'] == 'none':
                        ED += ' The dead pixel did not emit any light.'
                    elif Answers['Dead Pixel color'] == 'black':
                        ED += ' The dead pixel emmitted a black light.'
                    elif Answers['Dead Pixel color'] == 'white':
                        ED += ' The dead pixel emmitted a white light.'
                    elif Answers['Dead Pixel color'] == 'blue':
                        ED += ' The dead pixel emmitted a blue light.'

                if 'Cracked Screen Description' in Answers.keys():
                    ED += '\n\nAdditional information was provided at the time of the report.'

                    if Answers['Cracked Screen Description'][0] == 'Free Text':
                        ED += ' The crack was described as: ' + add_period_to_end(
                            Answers['Cracked Screen Description'][1])

                    if Answers['Cracked Screen state'] == 'Does not turn on':
                        ED += ' The TV was unable to turn on.'
                    elif Answers['Cracked Screen state'] == 'Turns on but displays nothing':
                        ED += ' The TV was able to turn on but did not display anything.'
                    elif Answers['Cracked Screen state'] == 'Turns on but displays distorted image':
                        ED += ' The TV was able to turn on but displayed a distorted image.'
                    elif Answers['Cracked Screen state'] == 'Turns on and displays fine image other than cracked area':
                        ED += ' The TV was able to turn on, and the image was fine except at the crack location.'

                    if Answers['Cracked Screen cause'][0] == 'Free Text':
                        ED += ' The cause was known and described as: ' + add_period_to_end(
                            Answers['Cracked Screen cause'][1])

            elif CS is Dog_Food_CS:
                ED += f'A {Reporter_Type_Occupation} contacted our company on {Complaint_Receipt_Date} with a dog food complaint.'
                ED += f' The product code was {Product_Code} and batch number was {Serial_Number}.'

                if Answers['When'] == 'Before':
                    ED += ' The issue was noticed before the dog ate it.'
                elif Answers['When'] == 'After':
                    ED += ' The issue was noticed after the dog ate it.'
                elif Answers['When'] == 'Unknown':
                    ED += ' It was unknown whether the dog ate any of it.'

                if Answers['Breed'][0] == 'Free Text':
                    ED += ' The breed involved was: ' + add_period_to_end(
                        Answers['Breed'][1])
                elif Answers['Breed'] == 'German Shepherd':
                    ED += ' The dog was a German Shepherd.'
                elif Answers['Breed'] == 'Pug':
                    ED += ' The dog was a pug.'
                elif Answers['Breed'] == 'Poodle':
                    ED += ' The dog was a poodle.'
                elif Answers['Breed'] == 'Rottweiler':
                    ED += ' The dog was a rottweiler.'

                if Answers['Quality'][0] == 'Free Text':
                    ED += ' The food was noticably different than usual and described as: ' + add_period_to_end(
                        Answers['Quality'][1])
                elif Answers['Quality'] == 'No':
                    ED += ' Nothing unusual in appearance or odor was noticed with the food.'
                elif Answers['Quality'] == 'Unknown':
                    ED += ' It was unknown if anything unusual was noticed with the food.'

                if Answers['Health'] == 'Fine':
                    ED += ' There were no negative health impacts to the dog.'
                elif Answers['Health'] == 'Fine but refused to eat':
                    ED += ' The dog refused to eat any of it, so there were no health impacts.'
                elif Answers['Health'] == 'Minor Illness':
                    ED += ' The dog ate some of it and became slightly sick.'
                elif Answers['Health'] == 'Serious Illness':
                    ED += ' The dog ate some of it and became seriously sick.'
                elif Answers['Health'] == 'Death':
                    ED += ' The dog died as a result of eating it.'

            elif CS is Car_Tire_CS:
                ED += f'This case was reported to our company by a {Reporter_Type_Occupation} on {Complaint_Receipt_Date}.'
                ED += f' The product involved was {Product_Family}, product code {Product_Code}, batch number {Serial_Number}.'

                if Answers['Issue'][0] == 'Free Text':
                    ED += ' The issue was described as: ' + add_period_to_end(
                        Answers['Issue'][1])

                if Answers['When'] == 'Yes':
                    ED += ' It was noticed while driving.'
                elif Answers['When'] == 'No':
                    ED += ' It was not noticed while driving.'
                elif Answers['When'] == 'Unknown':
                    ED += ' It was unknown if the issue was noticed while driving.'

                if Answers['Milage'] == 'None':
                    ED += ' The tire was brand new.'
                elif Answers['Milage'] == '1-10000':
                    ED += ' The tire had less than 10k miles on it when this was noticed.'
                elif Answers['Milage'] == '10000-30000':
                    ED += ' The tire had between 10 and 30k miles on it when this was noticed.'
                elif Answers['Milage'] == '30000+':
                    ED += ' The tire had over 30k miles on it when this was noticed.'
                elif Answers['Milage'] == 'Unknown':
                    ED += ' It was unknown how many miles the tire had when this was noticed.'

                if Answers['Prior Events'] == 'Accident':
                    ED += ' The car was in an accident shortly prior to the issue, which may have caused/contributed.'
                elif Answers['Prior Events'] == 'very wet roads':
                    ED += ' The car being driven in heavy rain conditions, which may have caused/contributed.'
                elif Answers['Prior Events'] == 'Potholes':
                    ED += ' The car drove over potholes prior to the issue, which may have caused/contributed.'
                elif Answers['Prior Events'] == 'No':
                    ED += ' There was no accident or unusual driving conditions prior to the issue.'

                if 'Tire Leak hole' in Answers.keys():
                    ED += '\n\nAdditional information was provided at the time of the report.'

                    if Answers['Tire Leak hole'][0] == 'Free Text':
                        ED += ' A puncture was visible and described as: ' + add_period_to_end(
                            Answers['Tire Leak hole'][1])
                    elif Answers['Tire Leak hole'] == 'No':
                        ED += ' There was not a visible puncture hole.'
                    elif Answers['Tire Leak hole'] == 'Unknown':
                        ED += ' It was unknown if there was a visible puncture hole.'

                    if Answers['Tire Leak object'][0] == 'Free Text':
                        ED += ' The reporter drove over an object described as: ' + add_period_to_end(
                            Answers['Tire Leak object'][1])
                    elif Answers['Tire Leak object'] == 'No':
                        ED += ' The reporter did not notice driving over anything that may have caused this.'

                    if Answers['Tire Leak maintenance'] == 'Yes':
                        ED += ' Proper tire pressure was maintained throughout the life of the car.'
                    elif Answers['Tire Leak maintenance'] == 'No':
                        ED += ' The reporter admitted to not maintaining proper tire pressure prior to this issue.'
                    elif Answers['Tire Leak maintenance'] == 'Unknown':
                        ED += ' It was unknown if proper tire pressure was maintained prior to this issue.'

            ED = Capitalize_Sentences(ED)
            ED = acronym_Correction(ED)
            ED = regex.sub(r'\n\nAdditional information was provided at the time of the report\.$', '', ED,
                           flags=regex.MULTILINE)

            if 'Extra Callscript Miss' in Answers.keys():
                ED += '\n\n' + '[' + ' AND '.join(Answers['Extra Callscript Miss']) + ' MISSING EXTRA!!!]'

            break
        elif Miss_Type:
            break

    if Miss_Type:
        x['Generated Text'] = f'{Miss_Type}'
    elif ED:
        x['Generated Text'] = ED
    else:
        x['Generated Text'] = '[QUESTIONNAIRE NOT DETECTED!!!]'

    print(f'    GENERATED TEXT FOR ID={x["ID"]}\n{x["Generated Text"]}\n\n')
    return x

df_ora = pd.read_excel(input_filepath)
df_ora['Complaint Receipt Date'] = df_ora['Complaint Receipt Date'].astype('str')
df_ora['Date of Event'] = df_ora['Date of Event'].astype('str')
df_ora.fillna('', inplace=True)
df_ora.drop_duplicates(subset='ID', keep=False, inplace=True)
print(df_ora[df_ora['ID']==3]['Questionnaire Text'].iloc[0])
tic = time.perf_counter()
df_ora = df_ora.apply(text_Generator, axis=1)
toc = time.perf_counter()
print(f"Ran in {(toc - tic) / len(df_ora):0.4f} seconds per record")
excelfile = pd.ExcelWriter(f'{output_folder_path}Output {format(date.today())}.xlsx', engine='xlsxwriter')
df_ora.to_excel(excelfile, sheet_name='output', index=False)
excelfile.close()