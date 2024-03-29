# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.


#%% Imports
import pandas as pd 
import re

current_dir = './UN_Knowledge_Extraction/'
data_dir = current_dir + "data/"
output_dir = current_dir + "output/"


UN_DOCS = pd.read_csv(data_dir + "UN_RES_DOCS_2009_2018.csv") 
UN_DOCS_resolution_level = UN_DOCS[['SourceFile']].drop_duplicates().reset_index(drop=True)
UN_DOCS_resolution_level['Resolutuion_Session'] = ''
UN_DOCS_resolution_level['Resolutuion_Agenda_item'] = ''
UN_DOCS_resolution_level['Resolutuion_Number'] = ''
UN_DOCS_resolution_level['Resolutuion_Title'] = ''
UN_DOCS_resolution_level['Resolutuion_Adoption_DateMonthYear'] = ''
UN_DOCS_resolution_level['Resolutuion_Adoption_Day'] = ''
UN_DOCS_resolution_level['Resolutuion_Adoption_Month'] = ''
UN_DOCS_resolution_level['Resolutuion_Adoption_Year'] = ''

for index, row in UN_DOCS_resolution_level.iterrows():
    Resolution_Info = [''] * 5 
    SourceFile = row['SourceFile']    
    SourceFile_info_paragraphs = UN_DOCS.loc[UN_DOCS['SourceFile'] == SourceFile].sort_values(by=['Index']).fillna('').reset_index(drop=True)
    for i in range(len(SourceFile_info_paragraphs)):
        Content = SourceFile_info_paragraphs.loc[i,'Content']
        Type = SourceFile_info_paragraphs.loc[i,'Type']
        if(Resolution_Info[0] == '' and Type == 'Session'): 
            Resolution_Info[0] = Content
        elif(Resolution_Info[1] == '' and Type == 'AgendaItem'): 
                Resolution_Info[1] = Content
        elif(Resolution_Info[2] == '' and Resolution_Info[3] == '' and re.match('(\d+/\d+)\s{0,1}\.\s{0,1}(.*)', Content)):
            Resolution_Info[2] = re.match('(\d+/\d+)\s{0,1}\.\s{0,1}(.*)',Content).groups()[0]
            Resolution_Info[3] = re.match('(\d+/\d+)\s{0,1}\.\s{0,1}(.*)',Content).groups()[1] 
        elif(Resolution_Info[4] == '' and re.match('(.*)on (\d{1,2}\s\w+\s\d{4})$', Content)):
            Resolution_Info[4] = re.match('(.*)on (\d{1,2}\s\w+\s\d{4})$', Content).groups()[1] 
    UN_DOCS_resolution_level.iloc[index]['Resolutuion_Session'] =  Resolution_Info[0]
    UN_DOCS_resolution_level.iloc[index]['Resolutuion_Agenda_item'] = Resolution_Info[1]
    UN_DOCS_resolution_level.iloc[index]['Resolutuion_Number'] = Resolution_Info[2]
    UN_DOCS_resolution_level.iloc[index]['Resolutuion_Title'] = Resolution_Info[3]
    UN_DOCS_resolution_level.iloc[index]['Resolutuion_Adoption_DateMonthYear'] = Resolution_Info[4]
    if (Resolution_Info[4] != ''):
        UN_DOCS_resolution_level.iloc[index]['Resolutuion_Adoption_Day'] = re.match(r'(\d{1,2})\s(\w+)\s(\d{4})', Resolution_Info[4]).groups()[0]
        UN_DOCS_resolution_level.iloc[index]['Resolutuion_Adoption_Month'] = re.match(r'(\d{1,2})\s(\w+)\s(\d{4})', Resolution_Info[4]).groups()[1]   
        UN_DOCS_resolution_level.iloc[index]['Resolutuion_Adoption_Year'] = re.match(r'(\d{1,2})\s(\w+)\s(\d{4})', Resolution_Info[4]).groups()[2]

UN_DOCS_resolution_level.to_excel(output_dir + 'output_UN_DOCS_resolution_level.xlsx')





