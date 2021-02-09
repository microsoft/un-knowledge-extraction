# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.


#%% Imports

import pandas as pd 
import numpy as np
from nltk import word_tokenize
from nltk.corpus import stopwords
from nltk.tokenize import RegexpTokenizer
import re
import string
import spacy
spacy_nlp = spacy.load('en')
import gensim
from scipy import spatial
from collections import Counter
import matplotlib.pyplot as plt
import difflib


stop_words = set(stopwords.words('english'))

current_dir = './UN_Knowledge_Extraction/'
data_dir = current_dir + "Data/"
output_dir = current_dir + "Output/"

UN_DOCS_Paragraphs = pd.read_csv(data_dir + "UN_RES_DOCS_2009_2018.csv").fillna('').reset_index(drop=True)
w2v_google = gensim.models.KeyedVectors.load_word2vec_format(data_dir + 'GoogleNews-vectors-negative300.bin.gz', binary=True)


UNBIS_terms = pd.read_csv(data_dir + "UNBIS_terms.csv", encoding='cp1252')
UNBIS_terms = [term.lower() for term in UNBIS_terms['Term'].unique().tolist()]

SDG_Targets_Indicators = pd.read_csv(data_dir + "SDG_Targets_Indicators.csv", encoding='cp1252')
SDG = list(SDG_Targets_Indicators['SDG'].drop_duplicates())

Targets_SDG_dict = pd.Series(SDG_Targets_Indicators.loc[SDG_Targets_Indicators.Type == 'Targets'].SDG.values,index=SDG_Targets_Indicators.loc[SDG_Targets_Indicators.Type == 'Targets'].Content).to_dict()
Indicators_SDG_dict = pd.Series(SDG_Targets_Indicators.loc[SDG_Targets_Indicators.Type == 'Indicators'].SDG.values,index=SDG_Targets_Indicators.loc[SDG_Targets_Indicators.Type == 'Indicators'].Content).to_dict()

Targets = list(SDG_Targets_Indicators.loc[SDG_Targets_Indicators.Type == 'Targets']['Content'].drop_duplicates())
Indicators = list(SDG_Targets_Indicators.loc[SDG_Targets_Indicators.Type == 'Indicators']['Content'].drop_duplicates())


SDG_Targets_Indicators_High_Frequency_Words = dict()
for SDG in list(SDG_Targets_Indicators.SDG.unique()):
    target = [key for key,value in Targets_SDG_dict.items() if value == SDG]
    indicator = [key for key,value in Indicators_SDG_dict.items() if value == SDG]
    tokenizer = RegexpTokenizer(r'\w+')
    all_words = [w for w in tokenizer.tokenize(' '.join(target + indicator).lower().replace('\t',' ')) if w not in stop_words]
    SDG_Targets_Indicators_High_Frequency_Words[SDG] = Counter(all_words).most_common(10)

for SDG in SDG_Targets_Indicators_High_Frequency_Words.keys():
    print(SDG, SDG_Targets_Indicators_High_Frequency_Words[SDG])


preambular_verb_list = [
        'acknowledging',
        'acting',
        'adhering',
        'affirming',
        'agreeing',
        'alarmed',
        'taking',  ## taking into consideration
        'anxious',
        'appreciating',
        'asserting',
        'attaching',
        'aware',
        'bearing',
        'being',
        'believing',
        'cognizant',
        'commemorating',
        'commending',
        'concerned',
        'concluding',
        'concurring',
        'confident',
        'confirming',
        'conscious',
        'considering',
        'continuing',
        'convinced',
        'deeming', 
        'deploring',
        'disturbed',
        'grieved',
        'perturbed',
        'regretting',
        'shocked',
        'desiring',
        'desirous',
        'deternied',
        'distressed',
        'disturbed',
        'emphasizing',
        'encouraged',
        'endorsing',
        'expressing',
        'faithful',
        'fearing',
        'noting',
        'recalling',
        'gratified',
        'alarmed',
        'guided',
        'having',
        'indignant',
        'holding',
        'hopeful',
        'conformity',  
        'pursuance',  
        'inspired',
        'invoking',
        'opinion',
        'keeping',
        'mindful',
        'observing',
        'outraged',
        'paying',
        'pending',
        'persuaded',
        'realizing',
        'recognizing',
        'recollecting',
        'referring',
        'regretting',
        'reiterating',
        'restating',
        'seeking',
        'sharing',
        'stressing',
        'striving',
        'condemning',
        'taking',
        'trusting',
        'underlining',
        'urging',
        'viewing',
        'warning',
        'welcoming',
        'wishing',
        'preventing',
        ]

operative_verb_list = [
        'accept', 'accepts',
        'recommend', 'recommends',
        'acknowledge', 'acknowledges',
        'address', 'addresses',
        'adopt', 'adopts',
        'proclaim', 'proclaims',
        'affirm', 'affirms',
        'appeal', 'appeals',
        'call', 'calls', 
        'draw', 'draws',  
        'pledge', 'pledges',
        'reiterate', 'reiterates',
        'request', 'requests',
        'agree', 'agrees',
        'decide', 'decides',
        'endorse', 'endorses',
        'invite', 'invites', 
        'note', 'notes',
        'welcome', 'welcomes',
        'amend', 'amends',
        'applaud', 'applauds',
        'appoint', 'appoints',
        'approve', 'approves',
        'assert', 'asserts',
        'assure', 'assures',
        'authorize', 'authorizes',
        'await', 'awaits',
        'believe', 'believes',
        'condemn', 'condemns',
        'censure', 'censures',
        'commend', 'commends',
        'commission', 'commissions',
        'compliment', 'compliments',
        'concur', 'concurs',
        'confirm', 'confirms',
        'congratulate', 'congratulates',
        'consider', 'considers',
        'convey', 'conveys',
        'declare', 'declares',
        'deem', 'deems',
        'appreciate', 'appreciates',
        'deplore', 'deplores',
        'defer', 'defers',
        'demand', 'demands',
        'denounce', 'denounces',
        'deprecate', 'deprecates',
        'designate', 'designates',
        'desire', 'desires',
        'determine', 'determines',
        'direct', 'directs',
        'dissolve', 'dissolves',
        'draw', 'draws',
        'emphasize', 'emphasizes',
        'empower', 'empowers',
        'encourage', 'encourages',
        'entrust', 'entrusts',
        'envisage', 'envisages',
        'establish', 'establishes',
        'exhort', 'exhorts',
        'expect', 'expects',
        'express', 'expresses',
        'extend', 'extends',
        'maintain', 'maintains',
        'support', 'supports',
        'formulate', 'formulates',
        'share', 'shares',
        'reaffirm', 'reaffirms',
        'insist', 'insists',
        'instruct', 'instructs',
        'invite', 'invites',
        'look', 'looks', 
        'make', 'makes',
        'mandate', 'mandates',
        'offer', 'offers',
        'pay', 'pays',
        'propose', 'proposes',
        'realize', 'realizes',
        'reassert', 'reasserts',
        'reassure', 'reassures',
        'recall', 'recalls',
        'recognize', 'recognizes',
        're-emphasize', 're-emphasizes',
        'refer', 'refers',
        'regard', 'regards',
        'register', 'registers',
        'regret', 'regrets',
        'reject', 'rejects',
        'remind', 'reminds',
        'renew', 'renews',
        'resolve', 'resolves',
        'seize', 'seizes',
        'set', 'sets',
        'warn', 'warns',
        'state', 'states',
        'stress', 'stresses',
        'suggest', 'suggests',
        'take', 'takes',
        'transmit', 'transmits',
        'trust', 'trusts',
        'underline', 'underlines',
        'urge', 'urges',
        ]


UN_DOCS_Paragraphs['First_Action_Verb'] = ''
UN_DOCS_Paragraphs['Paragraph_Type'] = ''
UN_DOCS_Paragraphs['Key_Terms'] = [list() for x in range(len(UN_DOCS_Paragraphs.index))]
UN_DOCS_Paragraphs['Referenced_Resolutions'] = [list() for x in range(len(UN_DOCS_Paragraphs.index))]
UN_DOCS_Paragraphs['Referenced_Resolutions_Dates'] = [dict() for x in range(len(UN_DOCS_Paragraphs.index))]
UN_DOCS_Paragraphs['SDG'] = [list() for x in range(len(UN_DOCS_Paragraphs.index))]


for index, row in UN_DOCS_Paragraphs.iterrows():
    if index % 10000 == 0:
        print(index)
    Content = row['Content'].replace('\t',' ')
    Content = ''.join(filter(lambda x:x in string.printable, Content))
    Content = Content.translate(str.maketrans('', '', '(),:;?@{|}~.'))
    Content = Content.translate(str.maketrans('', '', string.digits))
    tokenized_word = word_tokenize(Content.lower())
    Content_space_seperated = " " +  " ".join(tokenized_word) + " "        
    word_count = len(tokenized_word)


    if row['Type'] == 'Paragraph' and word_count >= 10:
        first_action_verb = ''
        try:
            first_action_verb = next(word for word in tokenized_word[:10] if word in preambular_verb_list + operative_verb_list)
        except Exception:
            pass    
        if Content[0].islower() == False:
            UN_DOCS_Paragraphs.loc[index, 'First_Action_Verb'] = first_action_verb
        if first_action_verb in preambular_verb_list and Content[0].islower() == False:
            UN_DOCS_Paragraphs.loc[index, 'Paragraph_Type'] = 'preambular'
        elif first_action_verb in operative_verb_list and Content[0].islower() == False:
            UN_DOCS_Paragraphs.loc[index, 'Paragraph_Type'] = 'operative' 
        elif Content[0].islower() == True:       
            previous_paragraph_types = list(UN_DOCS_Paragraphs.Paragraph_Type[(index-5):(index-1)])  
            previous_paragraph_types_non_empty = [x for x in previous_paragraph_types if x != '']
            if len(previous_paragraph_types_non_empty) >= 1:
                UN_DOCS_Paragraphs.loc[index, 'Paragraph_Type'] = previous_paragraph_types_non_empty[-1]
             
        matching_terms = list(set([term for term in UNBIS_terms if " " + term + " " in Content_space_seperated]))
        matching_terms.sort(key=len, reverse=True)
        key_terms = []
        if matching_terms is not None:
            for i in range(len(matching_terms)):
                if matching_terms[i] in Content:
                    key_terms.append(matching_terms[i])
                    Content = Content.replace(matching_terms[i], '')
        UN_DOCS_Paragraphs.at[index, 'Key_Terms'] = key_terms
        
        Referenced_Resolutions = re.findall(r'resolutions \w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)* .* and all subsequent related resolutions|resolutions \w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)* of [0-9]{1,2} [A-Za-z]{3,9} [0-9]{4}.* and \w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)* of [0-9]{1,2} [A-Za-z]{3,9} [0-9]{4}|resolutions \w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)* and \w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)* of [0-9]{1,2} [A-Za-z]{3,9} [0-9]{4}|resolution \w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)* of [0-9]{1,2} [A-Za-z]{3,9} [0-9]{4}|resolutions \w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)*.* and \w*-*\d+[/]*\d+\s*\(*\w*-*\w*\)* of [0-9]{1,2} [A-Za-z]{3,9} [0-9]{4}|resolutions \w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)*.* and \w*-*\d+[/]*\d+\s*\(*\w*-*\w*\)*|resolution \w*-*\d+[/]*[.]*\d+ \(\w*-*\w*\)|resolution \w*-*\d+[/]*[.]*\d+', Content)
        Referenced_Resolutions_Dates = []
        for referenced_resolution in Referenced_Resolutions:
            referenced_resolution = re.sub(' January ', '/01/', referenced_resolution)
            referenced_resolution = re.sub(' February ', '/02/', referenced_resolution)
            referenced_resolution = re.sub(' March ', '/03/', referenced_resolution)            
            referenced_resolution = re.sub(' April ', '/04/', referenced_resolution)            
            referenced_resolution = re.sub(' May ', '/05/', referenced_resolution)            
            referenced_resolution = re.sub(' June ', '/06/', referenced_resolution)
            referenced_resolution = re.sub(' July ', '/07/', referenced_resolution)
            referenced_resolution = re.sub(' August ', '/08/', referenced_resolution)
            referenced_resolution = re.sub(' September ', '/09/', referenced_resolution)
            referenced_resolution = re.sub(' October ', '/10/', referenced_resolution)
            referenced_resolution = re.sub(' November ', '/11/', referenced_resolution)
            referenced_resolution = re.sub(' December ', '/12/', referenced_resolution)   
            referenced_resolution_split = re.split(',|and', referenced_resolution)
            for resolution in referenced_resolution_split:
                if bool(re.search('resolution\w* (.*) of ([0-9]{1,2}/[0-9]{2}/[0-9]{4})', resolution)):
                    resolution_number = re.findall(r'resolution\w* (.*) of', resolution)[0]
                    date = re.findall(r'of ([0-9]{1,2}/[0-9]{2}/[0-9]{4})', resolution)[0]
                elif bool(re.search('\s*(.*) of ([0-9]{1,2}/[0-9]{2}/[0-9]{4})', resolution)):       
                    resolution_number = re.findall(r'\s*(.*) of', resolution)[0]
                    date = re.findall(r'of ([0-9]{1,2}/[0-9]{2}/[0-9]{4})', resolution)[0]  
                elif bool(re.search('resolution\w* (.*)', resolution)):                         
                    resolution_number = re.findall(r'resolution\w* (.*)', resolution)[0]
                    date = 'NA'    
                elif bool(re.search('\w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)*', resolution)):
                    resolution_number = re.findall(r'\w*-*\d+[/]*[.]*\d+\s*\(*\w*-*\w*\)*', resolution)[0]
                    date = 'NA'
                Referenced_Resolutions_Dates[resolution_number] = date
        UN_DOCS_Paragraphs.at[index, 'Referenced_Resolutions'] = Referenced_Resolutions
        UN_DOCS_Paragraphs.at[index, 'Referenced_Resolutions_Dates'] = Referenced_Resolutions_Dates

        if any(x in tokenized_word for x in ['poverty', 'poor']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('No Poverty')
        elif any(x in Content.lower() for x in ['hunger', 'hungry', 'malnutrition', 'food crisis', 'sufficient food', 'food producers', 'food production', 'food reserves', 'food price', 'food insecurity', 'food security', 'undernutrition']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Zero Hunger')
        elif any(x in tokenized_word for x in ['health', 'well-being', 'mortality', 'disease']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Good Health and Well-Being')
        elif any(x in tokenized_word for x in ['education', 'educational']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Quality Education')
        elif any(x in tokenized_word for x in ['gender equality']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Gender Equality')
        elif any(x in tokenized_word for x in ['water', 'sanitation', 'wastewater']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Clean Water and Sanitation')
        elif any(x in tokenized_word for x in ['energy', 'renewable']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Affordable and Clean Energy')
        elif any(x in tokenized_word for x in ['labour-intensive', 'employment']) or any(x in Content.lower() for x in ['child labour', 'labour rights',  'decent work', 'economic growth', 'economic productivity']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Decent Work and Economic Growth')
        elif any(x in tokenized_word for x in ['industry', 'innovation', 'infrastructure']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Industry, Innovation and Infrastructure')
        elif any(x in tokenized_word for x in ['inequalities', 'inequality']) and (not any(x in Content.lower() for x in ['gender equality'])):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Reduced Inequalities')            
        elif 'sustainable cities' in Content.lower():
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Sustainable Cities and Communities')
        elif any(x in Content.lower() for x in ['consumption and production']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Responsible Consumption and Production')
        elif any(x in Content.lower() for x in ['climate change', 'climate-related', 'natural disaster', 'national disaster', 'local disaster']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Climate Action')
        elif any(x in tokenized_word for x in ['marine', 'fisheries', 'coastal']) or any(x in Content.lower() for x in ['oceans and seas']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Life Below Water') 
        elif any(x in tokenized_word for x in ['biodiversity', 'land ', 'inland', 'species']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Life on Land')     
        elif 'institutions' in tokenized_word and any(x in tokenized_word for x in ['peace', 'justice', 'strong']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Peace, Justice and Strong Institutions')
        elif any(x in tokenized_word for x in ['partner', 'partners', 'partnership', 'partnerships']):
            UN_DOCS_Paragraphs.at[index, 'SDG'].append('Partnerships for the Goals')

w2v_Targets = []
w2v_Indicators = []
Targets_isalpha = []
Indicators_isalpha = []

for i in range(len(Targets)):
    tokenized_word = word_tokenize(Targets[i].lower())
    tokenized_word = [word for word in tokenized_word if len(word) > 1] 
    tokenized_word = [word for word in tokenized_word if word.isalpha()]
    Targets_isalpha.append(' '.join(tokenized_word))
    words_in_vocab = [word for word in tokenized_word if word in w2v_google.vocab]
    w2v_sum = np.sum(w2v_google[words_in_vocab], axis=0)
    #w2v_average = np.average(w2v_google[words_in_vocab], axis=0)    
    w2v_Targets.append(w2v_sum)

for i in range(len(Indicators)):
    tokenized_word = word_tokenize(Indicators[i].lower())
    tokenized_word = [word for word in tokenized_word if len(word) > 1] 
    tokenized_word = [word for word in tokenized_word if word.isalpha()]
    Indicators_isalpha.append(' '.join(tokenized_word))
    words_in_vocab = [word for word in tokenized_word if word in w2v_google.vocab]
    w2v_sum = np.sum(w2v_google[words_in_vocab], axis=0)
    #w2v_average = np.average(w2v_google[words_in_vocab], axis=0)    
    w2v_Indicators.append(w2v_sum)


def Common_Substring(string1, string2):
    substrings = []
    matches = difflib.SequenceMatcher(None, string1, string2).get_matching_blocks()
    for match in sorted(matches, key=lambda x: x[2], reverse=True):  
        substrings.append(string1[match.a:match.a + match.size])
    return substrings


similarity_threshold_target = 0.9
similarity_threshold_indicator = 0.9

UN_DOCS_Paragraphs['Closest_Target'] = [list() for x in range(len(UN_DOCS_Paragraphs.index))]
UN_DOCS_Paragraphs['Closest_Indicator'] = [list() for x in range(len(UN_DOCS_Paragraphs.index))]
UN_DOCS_Paragraphs['Closest_Target_Similarity_Score'] = 0.0
UN_DOCS_Paragraphs['Closest_Indicator_Similarity_Score'] = 0.0

for row_index in range(len(UN_DOCS_Paragraphs)):
    if row_index % 1000 == 0:
        print(row_index)    
    if UN_DOCS_Paragraphs.loc[row_index, 'Type'] == 'Paragraph':
        paragraph = UN_DOCS_Paragraphs.loc[row_index, 'Content']
        tokenized_word = word_tokenize(paragraph.lower().replace('\t',' '))
        tokenized_word = [word for word in tokenized_word if len(word) > 1] 
        tokenized_word = [word for word in tokenized_word if word.isalpha()]
        paragraph_isalpha = ' '.join(tokenized_word)

        similarity_with_target_common_substring = [] 
        for i in range(len(Targets_isalpha)):
            paragraph_target_common_substring = Common_Substring(paragraph_isalpha, Targets_isalpha[i])
            if len(paragraph_target_common_substring) == 0:  
                similarity_with_target_common_substring.append(0.0)
            elif len(paragraph_target_common_substring) >= 1:
                paragraph_target_common_substring = paragraph_target_common_substring[:3] 
                paragraph_target_common_substring_aggregated = ' '.join(paragraph_target_common_substring)
                words_common_substring = paragraph_target_common_substring_aggregated.split()
                words_common_substring = [word for word in words_common_substring if word in tokenized_word]
                words_common_substring_in_vocab = [word for word in words_common_substring if word in w2v_google.vocab]
                if len(words_common_substring_in_vocab) >= 1:
                    w2v_common_substring = np.sum(w2v_google[words_common_substring_in_vocab], axis=0) 
                    similarity_with_target_common_substring.append(1 - spatial.distance.cosine(w2v_Targets[i], w2v_common_substring))
                else: 
                    similarity_with_target_common_substring.append(len(words_common_substring) / len(Targets_isalpha[i].split()))

        similarity_with_indicator_common_substring = [] 
        for i in range(len(Indicators_isalpha)):
            paragraph_indicator_common_substring = Common_Substring(paragraph_isalpha, Indicators_isalpha[i])
            if len(paragraph_indicator_common_substring) == 0: 
                similarity_with_indicator_common_substring.append(0.0)
            elif len(paragraph_indicator_common_substring) >= 1:
                paragraph_indicator_common_substring = paragraph_indicator_common_substring[:3]
                paragraph_indicator_common_substring_aggregated = ' '.join(paragraph_indicator_common_substring)
                words_common_substring = paragraph_indicator_common_substring_aggregated.split()
                words_common_substring = [word for word in words_common_substring if word in tokenized_word]
                words_common_substring_in_vocab = [word for word in words_common_substring if word in w2v_google.vocab]
                if len(words_common_substring_in_vocab) >= 1:
                    w2v_common_substring = np.sum(w2v_google[words_common_substring_in_vocab], axis=0) 
                    similarity_with_indicator_common_substring.append(1 - spatial.distance.cosine(w2v_Indicators[i], w2v_common_substring))
                else:  ## none of the words in common substring are in vocab
                    similarity_with_indicator_common_substring.append(len(words_common_substring) / len(Indicators_isalpha[i].split()))

        UN_DOCS_Paragraphs.loc[row_index, 'Closest_Target_Similarity_Score'] = max(similarity_with_target_common_substring)
        UN_DOCS_Paragraphs.loc[row_index, 'Closest_Indicator_Similarity_Score'] = max(similarity_with_indicator_common_substring)

        similar_target_index = [i for i,similarity in enumerate(similarity_with_target_common_substring) if similarity >= similarity_threshold_target]
        similar_indicator_index = [i for i,similarity in enumerate(similarity_with_indicator_common_substring) if similarity >= similarity_threshold_indicator]

        if ((len(similar_target_index) >= 1) and (max(similarity_with_target_common_substring) >= max(similarity_with_indicator_common_substring))):
            most_similar_target_index = similarity_with_target_common_substring.index(max(similarity_with_target_common_substring))
            most_similar_target = Targets[most_similar_target_index]
            UN_DOCS_Paragraphs.at[row_index, 'Closest_Target'].append(most_similar_target)
            if Targets_SDG_dict[most_similar_target] not in UN_DOCS_Paragraphs.at[row_index, 'SDG']:
                UN_DOCS_Paragraphs.at[row_index, 'SDG'].append(Targets_SDG_dict[most_similar_target])  
        elif ((len(similar_indicator_index) >= 1) and (max(similarity_with_target_common_substring) <= max(similarity_with_indicator_common_substring))):
            most_similar_indicator_index = similarity_with_indicator_common_substring.index(max(similarity_with_indicator_common_substring))
            most_similar_indicator = Indicators[most_similar_indicator_index]
            UN_DOCS_Paragraphs.at[row_index, 'Closest_Indicator'].append(most_similar_indicator)
            if Indicators_SDG_dict[most_similar_indicator] not in UN_DOCS_Paragraphs.at[row_index, 'SDG']:
                UN_DOCS_Paragraphs.at[row_index, 'SDG'].append(Indicators_SDG_dict[most_similar_indicator])



country_list = pd.read_excel(data_dir + "country_list.xlsx").fillna('')
country_names = [country.strip().replace('&', 'and') for country in country_list['Country'].tolist()]

UN_agencies = pd.read_excel(data_dir + "agencies.xlsx").fillna('')
UN_known_orgs = pd.read_excel(data_dir + "un_entities_20191017.xlsx").fillna('')

UN_corporate_names = pd.read_excel(data_dir + "names_A60-72.xlsx").fillna('')
UN_corporate_names = [x for x in UN_corporate_names['Name'] if x not in country_names]
UN_corporate_names = [re.sub("[\(].*?[\)]", "", x).replace('UN', 'United Nations').replace('.','').strip() for x in UN_corporate_names]

additional_un_org_list = [
        'Advisory Committee on Administrative and Budgetary Questions',
        'African Union Mission in Somalia',
        'European Union Rule of Law Mission in Kosovo',
        'Special Political and Decolonization Committee (Fourth Committee)',
        'United Nations Conference on Environment and Development',
        'United Nations Entity for Gender Equality and the Empowerment of Women (UN-Women)',
        'Bretton Woods Institutions',
        'International Tribunal for the Former Yugoslavia',
        'United Nations Assistance Mission in Afghanistan',
        'United Nations Operation in Cte dIvoire',
        'Consultative Group on International Agricultural Research',
        ]

known_un_org_list = list(set(
        UN_agencies['Title'].tolist() 
        + UN_known_orgs['Entity'].tolist()
        + UN_corporate_names
        + additional_un_org_list
        ))
known_un_org_list = [x for x in known_un_org_list if x not in country_names]


known_un_org_list = [org.translate(str.maketrans('', '', ',;:."')) for org in known_un_org_list]
#known_un_org_list = [re.sub(r'[^\x00-\x7F]+',' ', org) for org in known_un_org_list]
known_un_org_list = [''.join([x if x in string.printable else '' for x in org]) for org in known_un_org_list] 
known_un_org_list = [' '.join(w for w in org.split()) for org in known_un_org_list]

known_un_org_w2v = dict()
for org in known_un_org_list:
    words_in_vocab = [word for word in word_tokenize(org.lower()) if word in w2v_google.vocab]
    if len(words_in_vocab) >= 1:
        w2v_sum = np.sum(w2v_google[words_in_vocab], axis=0)
        known_un_org_w2v[org] = w2v_sum
    else:
        known_un_org_w2v[org] = np.asarray([])


key_words_un_org_list = [ 
        'Committee',
        'Council',
        'Conference',
        'Fund',
        'Organization',
        'Entity',
        'Department',
        'Commission',
        'Court',
        'Board',
        'Community',
        'Office',
        'Association',
        'Government',
        'Group',
        'Summit',
        'Subcommittee',
        ]

key_words_not_un_org_list = [
        'Goal',
        'Goals',
        'Agenda',
        'Outcome',
        'Headquarters',
        'Declaration',
        'Account',
        'Implementation',
        'Territory',
        'Territories',
        'Act',
        'Action',
        'Actions',
        'Programme',
        'Agreement',
        'Partnership',
        'Protection of Civilian Persons',
        'Time of War',
        'Framework',
        'Frameworks',
        'Consensus',
        'Convention',
        'Conventions',
        'Related',
        'Resolution',
        'Resolutions',
        'Forum',
        'Meeting',
        'Strategy',
        'Eradicate',
        'General Service',
        'Document',
        'Deconstruction',
        'Status',
        'Statute',
        'Protocol',
        'Protocols',
        'Outcome',
        'Illicit',
        'Session',
        'A/RES/',
        'Movement',
        'Chair',
        'Treatment',
        'Platform',
        'Platforms',
        'Plan',
        'Weapons',
        'National Food Security',
        'Rules',
        'Budget',
        'Principle',
        'Principles',
        'System',
        'Systems',
        'Mechanism',
        'Report',
        'Pact',
        'Compact',
        'Trade',
        'Consequences',
        'United Nations Global Compact',
        'Facility',
        'Covenant',
        'Covenants',
        'Responsible',
        'Treaty',
        'Decade',
        'Wider United Nations',
        'Their',
        'Expert',
        'Personnel',
        'Conservation',
        'Field Service',
        'Information',
        'International Migration and Development',
        'Coordinator',
        'Armistice Line',
        'Further',
        'Day',
        'Week',
        'Month',
        'Year',
        'Criteria',
        'El Nio',
        'Fellowship',
        'Safety of Maritime Navigation',
        'Library',
        'Doha Development Round',
        'Journal',
        'Review',
        'Aid for Trade',
        'Sea',
        'Movement',
        'Zone',
        'International Health Regulations',
        'International Mother',
        'Goodwill Ambassadors',
        'Chronicle',
        'Involuntary Disappearances',
        'Impact',
        'Rapporteur',
        'Rapporteurs',
        'Record',
        'Records',
        'Ministers',
        'Panel',
        'University',
        'Yearbook',
        'Messengers',
        'Terrorism',
        'Dialogue',
        'Officer',
        'Target',
        'Targets',
        'Elimination',
        'Council established',
        'Repair and Assembly',
        'Countries and Peoples',
        'Model Strategies and Practical Measures',
        'Ways and means',
        'Challenge',
        'Network',
        'Safety and Security of Radioactive Sources',
        'Guideline',
        'Guidelines',
        'Parties',
        'Unregulated Fishing',
        'Discrimination',
        'Armed Robbery against Ships',
        'Regular',
        'International Search',
        'Process',
        'Branch',
        'Context',
        'Orthodox Good Friday',
        'Seascape',
        'Regional Security',
        'Cooperation for',
        'Application',
        'Volunteer',
        'Volunteers',
        'Fishing Vessels',
        'Alternative',
        'Green Paper',
        'Holy See',
        'Need of Assistance',
        'Olympic Truce',
        'Mutual Understanding',
        'Tapta',  
        'Census',
        'Sport for Development and Peace',
        'Campaign',
        'Protection of Child Victims of Trafficking',
        'Approach',
        'Service',
        'Commercial Shipping',
        'Reduction of Underwater Noise',
        'Chair',
        'Chairs',
        'Co-Chairs',
        ]


UN_DOCS_Paragraphs['word_cnt'] = 0
UN_DOCS_Paragraphs['Content_clean'] = ''

for index, row in UN_DOCS_Paragraphs.iterrows():
    if index % 10000 == 0:
        print(index)
    Content = row['Content'].replace('\t',' ')
    Content = Content.replace(',',', ') 
    Content = Content.replace(';','; ')     
    Content = Content.replace('.','. ')
    Content = re.sub(r'[0-9]{1,2}.', ' ', Content)
    Content = ''.join([x if x in string.printable else '' for x in Content])
    Content = ' '.join(w for w in Content.split() if not any(x.isdigit() for x in w)) 
    word_cnt = len(Content.split())
    UN_DOCS_Paragraphs.at[index, 'word_cnt'] = word_cnt
    UN_DOCS_Paragraphs.at[index, 'Content_clean'] = Content

UN_DOCS_Paragraphs = UN_DOCS_Paragraphs.sort_values(by=['SourceFile', 'Index'])


UN_DOCS_Resolutions_Content = UN_DOCS_Paragraphs.loc[(UN_DOCS_Paragraphs.Type == 'Paragraph')].groupby(['SourceFile'])['Content'].apply(' '.join).reset_index()
UN_DOCS_Resolutions_Content_clean = UN_DOCS_Paragraphs.loc[(UN_DOCS_Paragraphs.Type == 'Paragraph')].groupby(['SourceFile'])['Content_clean'].apply(' '.join).reset_index()
UN_DOCS_Resolutions = pd.merge(UN_DOCS_Resolutions_Content, UN_DOCS_Resolutions_Content_clean, on='SourceFile')

UN_DOCS_Resolutions['Organization_Names_known'] = [list() for x in range(len(UN_DOCS_Resolutions.index))]
UN_DOCS_Resolutions['Organization_Names_not_from_known_orginal'] = [list() for x in range(len(UN_DOCS_Resolutions.index))]
UN_DOCS_Resolutions['Organization_Names_not_from_known_inferred'] = [list() for x in range(len(UN_DOCS_Resolutions.index))]

for index, row in UN_DOCS_Resolutions.iterrows():
    if index % 100 == 0:
        print(index)
    Content_clean = row['Content_clean']
    known_orgs = [known_org for known_org in known_un_org_list if known_org in Content_clean]
    UN_DOCS_Resolutions.at[index, 'Organization_Names_known'] = known_orgs

    extracted_orgs = list(set([str(element) for element in spacy_nlp(Content_clean).ents if element.label_ == 'ORG']))
    extracted_orgs = [org for org in extracted_orgs if all(char not in org for char in ['_', '/', '.'])]
    for i in range(len(extracted_orgs)):
        extracted_org = extracted_orgs[i].translate(str.maketrans('', '', string.digits)) 
        extracted_org = extracted_org.translate(str.maketrans('', '', ',;:.()')) 
        if extracted_org.lower().startswith('the '): 
            extracted_orgs[i] = extracted_org[4:]
    extracted_orgs = list(set(extracted_orgs))   

    Organization_Names_not_from_known_orginal = []
    for org in extracted_orgs:
        if (
                len(org.split()) > 1 
                and (org not in known_un_org_list)
                and (not org.lower().split()[-1] in stop_words) 
                and ((not any(key_word.lower() in org.lower() for key_word in key_words_not_un_org_list)) or (any(key_word.lower() in org.lower() for key_word in key_words_un_org_list)))
                and (max([org.lower() in known_org.lower() for known_org in known_un_org_list]) == False) 
                #and (max([known_org.lower() in org.lower() for known_org in known_un_org_list]) == False)
                and (max([org.lower().split()[0] in [word for word in operative_verb_list if word.endswith('s')]]) == 0)
                and (max([word in preambular_verb_list for word in org.lower().split()]) == 0)
                and (' of the ' not in org) 
                ):
            Organization_Names_not_from_known_orginal.append(org)  
        elif (
                len(org.split()) > 1 
                and (org not in known_un_org_list)
                and (not org.lower().split()[-1] in stop_words) 
                and ((not any(key_word.lower() in org.lower() for key_word in key_words_not_un_org_list)) or (any(key_word.lower() in org.lower() for key_word in key_words_un_org_list))) 
                and (max([org.lower() in known_org.lower() for known_org in known_un_org_list]) == False)
                #and (max([known_org.lower() in org.lower() for known_org in known_un_org_list]) == False)
                and (max([org.lower().split()[0] in [word for word in operative_verb_list if word.endswith('s')]]) == 0) 
                and (max([word in preambular_verb_list for word in org.lower().split()]) == 0) 
                and (' of the ' in org)  
                ):
            org_split = org.split(' of the ')
            if (org_split[0] not in known_un_org_list) and (len(org_split[0].split()) > 1) and (org_split[1] in known_un_org_list):
                Organization_Names_not_from_known_orginal.append(org_split[0])
            elif (org_split[0] in known_un_org_list) and (org_split[1] not in known_un_org_list) and (len(org_split[1].split()) > 1):
                Organization_Names_not_from_known_orginal.append(org_split[1])
            elif (org_split[0] not in known_un_org_list) and (org_split[1] not in known_un_org_list): 
                Organization_Names_not_from_known_orginal.append(org)

    UN_DOCS_Resolutions.at[index, 'Organization_Names_not_from_known_orginal'] = Organization_Names_not_from_known_orginal
    

    if (len(known_orgs) > 0): 
        for org in UN_DOCS_Resolutions.at[index, 'Organization_Names_not_from_known_orginal']:
            tokenized_word = word_tokenize(org)
            tokenized_word_lower = word_tokenize(org.lower())
            words_in_vocab_lower = [word for word in tokenized_word_lower if word in w2v_google.vocab]
            if (len(words_in_vocab_lower) >= 1):
                org_w2v = np.sum(w2v_google[words_in_vocab_lower], axis=0)
            else:
                org_w2v = np.asarray([])
            common_words_length = []
            w2v_similarity = []
            for known_org in known_orgs:
                known_org_tokenized_word = word_tokenize(known_org)
                known_org_tokenized_word_lower = word_tokenize(known_org.lower())
                known_org_words_in_vocab_lower = [word for word in known_org_tokenized_word_lower if word in w2v_google.vocab]
                if len(known_org_words_in_vocab_lower) >= 1:
                    known_org_w2v = np.sum(w2v_google[known_org_words_in_vocab_lower], axis=0)
                else:
                    known_org_w2v = np.asarray([])
        
                common_words = [word for word in tokenized_word if (word in known_org_tokenized_word and word[0].isupper())]
                common_words_length.append(len(common_words))
                if ((len(org_w2v) == 0) or (len(known_org_w2v) == 0)):
                    w2v_similarity.append(0)
                else:
                    w2v_similarity.append(1 - spatial.distance.cosine(org_w2v, known_org_w2v))
            
            if (max(common_words_length) == 0): 
                UN_DOCS_Resolutions.at[index, 'Organization_Names_not_from_known_inferred'].append((org, org))
            else:
                if (len([l for l in common_words_length if l == max(common_words_length)]) == 1): 
                    known_org_index = common_words_length.index(max(common_words_length))
                    similarity_score = w2v_similarity[known_org_index]
                    known_org = known_orgs[known_org_index]
                    UN_DOCS_Resolutions.at[index, 'Organization_Names_not_from_known_inferred'].append((org, known_org, similarity_score))
                else: 
                    known_org_index = [i for i, x in enumerate(common_words_length) if x == max(common_words_length)]
                    max_similarity_score = max([w2v_similarity[index] for index in known_org_index])
                    max_similarity_known_org_index = w2v_similarity.index(max_similarity_score)
                    known_org = known_orgs[max_similarity_known_org_index]
                    UN_DOCS_Resolutions.at[index, 'Organization_Names_not_from_known_inferred'].append((org, known_org, max_similarity_score))
    

Organization_Names_not_from_known = UN_DOCS_Resolutions['Organization_Names_not_from_known_orginal'].tolist()
Organization_Names_not_from_known = [x for sublist in Organization_Names_not_from_known for x in sublist]
Organization_Names_not_from_known_cnt = Counter(Organization_Names_not_from_known)
Organization_Names_not_from_known_cnt = pd.DataFrame.from_dict(Organization_Names_not_from_known_cnt, orient='index').reset_index()
Organization_Names_not_from_known_cnt = Organization_Names_not_from_known_cnt.rename(columns={'index':'org_names', 0:'count'}).sort_values(by='count', ascending=False).reset_index(drop=True)


UN_DOCS_Paragraphs['Country'] = [list() for x in range(len(UN_DOCS_Paragraphs.index))]
UN_DOCS_Paragraphs['Organization_Names_known'] = [list() for x in range(len(UN_DOCS_Paragraphs.index))]
UN_DOCS_Paragraphs['Organization_Names_not_from_known_orginal'] = [list() for x in range(len(UN_DOCS_Paragraphs.index))]
UN_DOCS_Paragraphs['Organization_Names_not_from_known_inferred'] = [list() for x in range(len(UN_DOCS_Paragraphs.index))]

for index, row in UN_DOCS_Paragraphs.iterrows():
    if index % 10000 == 0:
        print(index)
    SourceFile = row['SourceFile']
    Content_clean = row['Content_clean']
    Organization_Names_known_Resolution = UN_DOCS_Resolutions.loc[UN_DOCS_Resolutions['SourceFile'] == SourceFile , 'Organization_Names_known'].tolist()[0]
    Organization_Names_not_from_known_orginal_Resolution = UN_DOCS_Resolutions.loc[UN_DOCS_Resolutions['SourceFile'] == SourceFile]['Organization_Names_not_from_known_orginal'].tolist()[0]
    Organization_Names_not_from_known_inferred_Resolution = UN_DOCS_Resolutions.loc[UN_DOCS_Resolutions['SourceFile'] == SourceFile]['Organization_Names_not_from_known_inferred'].tolist()[0]
    Country = [country for country in country_names if country.lower() in Content_clean.lower()]
    Organization_Names_known = []
    for org in Organization_Names_known_Resolution:
        if org.lower() in Content_clean.lower():
            Organization_Names_known.append(org)
    Organization_Names_not_from_known_orginal = []
    for org in Organization_Names_not_from_known_orginal_Resolution:
        if org in Content_clean:
            Organization_Names_not_from_known_orginal.append(org)     
    Organization_Names_not_from_known_inferred = []
    if len(Organization_Names_not_from_known_inferred_Resolution) >= 1:
        for org in Organization_Names_not_from_known_inferred_Resolution:
            if org[0] in Content_clean:
                Organization_Names_not_from_known_inferred.append(org)  
    UN_DOCS_Paragraphs.at[index, 'Country'] = Country         
    UN_DOCS_Paragraphs.at[index, 'Organization_Names_known'] = Organization_Names_known
    UN_DOCS_Paragraphs.at[index, 'Organization_Names_not_from_known_orginal'] = Organization_Names_not_from_known_orginal
    UN_DOCS_Paragraphs.at[index, 'Organization_Names_not_from_known_inferred'] = Organization_Names_not_from_known_inferred

UN_DOCS_Paragraphs = UN_DOCS_Paragraphs.drop(columns=['word_cnt', 'Content_clean'])
UN_DOCS_Paragraphs.to_excel(output_dir + 'output_UN_DOCS_paragraph_level.xlsx')

