#!/usr/bin/env python
# coding: utf-8

# Import the necessary library

from docx2python import docx2python
import re
from bs4 import BeautifulSoup as bs
import json
from pdfminer.high_level import extract_text
import docx2txt
import ntpath
import nltk
import requests
from wordcloud import WordCloud, STOPWORDS 
import matplotlib.pyplot as plt 
import pandas as pd
from collections import Counter
import re
import spacy
from spacy.matcher import Matcher
from nltk.corpus import stopwords
import win32com.client
import os



# We have created a json file to store the keywords of different sections.The following code is reading that json file.

f = open('reserved_words.json',)
reserved_words = json.load(f)
f.close()


# The following function reads the docx files and convert the text of that file into BeautifulSoup object.

def read_docx(path):
    file_ext = ntpath.basename(path)
    filename = file_ext.split('.')
    if filename[1] == 'pdf':
        word = win32com.client.Dispatch("Word.Application")
        word.visible = 0

        mypath = 'C:\\Users\\Admin\\Desktop\\JupyterProjects\\Upmovv\\'
        filename = os.path.basename(path)
        todocx = os.path.abspath(mypath + filename[0:-4] + ".docx")
        wb1 = word.Documents.Open(path)
        file1 = wb1.SaveAs(todocx, FileFormat=16)  # file format for docx
        wb1.Close()
        word.Quit()
        file55 = mypath + (filename[0:-4] + ".docx")
        file123 = docx2python(file55, html=True)
        soup = bs(file123.text)
        return soup

    if filename[1] == 'docx':
        file = docx2python(path, html=True)
        soup = bs(file.text)

        return soup

# The following function takes the BeautifulSoup object as an input and returns the size of all the fonts available in the text. 

def find_font_size(soup):
    fonts = soup.find_all('font')
    font_size = set()
    for font in fonts:
        font_size.add(font['size'])
        
    font_size_list = list(font_size)
    font_size_list.sort(reverse=True)
    
    return font_size_list

	
# The following function takes the BeautifulSoup object and list of font size as an input and returns the list of bold headings from the text using the font size.

def find_header(soup,font_size_list):
    header = []
    for i in font_size_list[0:5]:
        body = soup.find_all('font',attrs={'size':i})
        for bold in body:
            bold_text = bold.find_all('b')
            for x in bold_text:
                header.append(x.text)
    
    return header
        

# This function checks whether the given substring is part of the given string.

def check(string,substring):
    found = False
    for i in range(0,len(substring)):
        if string.lower().find(substring[i])!=-1:
            found = True
    return found


# This function takes the list of the bold header and reserved words as an input and returns the dictionary of matched keys from json files and value from headers.

def title_dict(header,reserved_words):
    title = {}
    
    for a in range(0,len(header)):
        for b in reserved_words:
            d = check(header[a].lower(),reserved_words[b])
            if d == True:
                if b in title:
                    title[b].append(header[a])
                    if header[a].lower() not in reserved_words[b]:
                        reserved_words[b].append(header[a].lower())
                    else:
                        continue
                    
                else:
                    title[b] = [header[a]]
                    if header[a].lower() not in reserved_words[b]:
                        reserved_words[b].append(header[a].lower())
                    else:
                        continue
    
    with open("reserved_words.json", "w") as outfile: 
        json.dump(reserved_words, outfile)                
    
    return title


# This function takes the dict of required sections and generate the score out of 100 as per the sections available in the dict.

def score_generator(required_sections):
    total_score = 0
    if 'contact' in required_sections.keys():
        total_score+=20
    if 'education' in required_sections.keys():
        total_score+=20
    if 'experience' in required_sections.keys():
        total_score+=20
    if 'objective' in required_sections.keys():
        total_score+=20
    if 'skill' in required_sections.keys():
        total_score+=20
        
    return total_score


# This function takes the text and list of headers and returns the sections as a dict using header[i] as a starting point and header[i+1] as ending point.

def generate_section(text,header):
    section = {}
    
    for i in range(0,len(header)):

        if i == len(header)-1:
            start = header[i]
#             end = text[-1]
            section[header[i]]=text[text.index(start):len(text)]
        else:
            start = header[i]
            end = header [i+1]
            section[header[i]]=text[text.index(start):text.index(end)]
    
    return section
    

# First, we have loaded the 'en_core_web_sm' model from spacy.The extract_name() function takes the text as an input and return the available names from the text.

nlp = spacy.load('en_core_web_sm')
matcher = Matcher(nlp.vocab)
def extract_name(resume_text):
    name = []
    nlp_text = nlp(resume_text)
    
    # First name and Last name are always Proper Nouns
    pattern = [{'POS': 'PROPN'}, {'POS': 'PROPN'}]
    
    matcher.add('NAME',[pattern])
    
    matches = matcher(nlp_text)
    
    for match_id, start, end in matches:
        span = nlp_text[start:end]
        name.append(span.text)
        return name

		
# The following function uses regular expression to find the appropriate pattern of phone numbers from the text and returns the list of all available phone numbers from the text.


PHONE_REG = re.compile(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]')
def extract_phone_number(input_text):
    phone_number = []
    phone = re.findall(PHONE_REG, input_text)

    if phone:
        number = ''.join(phone[0])

        if input_text.find(number) >= 0 and len(number) < 16:
            phone_number.append(number)
            return phone_number
    return None


# The following function uses regular expression to find the appropriate pattern of Email address from the text and returns the list of all available email ids from the text.

EMAIL_REG = re.compile(r'[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+')
def extract_emails(input_text):
    return re.findall(EMAIL_REG, input_text)


# The following function uses regular expression to find the appropriate pattern of linkedin profile from the text and returns the list of all available profile links from the text.

LINKED_REG = re.compile(r'https\:\/\/www\.linkedin+\.[a-zA-Z0-9/~\-_,&=\?\.;]+[^\.,\s<]')

def extract_linkedin(input_text):
    return re.findall(LINKED_REG, input_text)


# The following function checks whether the 'contact' section is available in the text. It takes the first section of the generated section and find that any of the contact details is available or not. If fond, then function returns True otherwise it returns False. 

def check_contact_info(sections,header):
    contact_found = False
    if extract_name(sections[header[0]]):
        contact_found = True
    elif extract_emails(sections[header[0]]):
        contact_found = True
    elif extract_linkedin(sections[header[0]]):
        contact_found = True
    elif extract_phone_number(sections[header[0]]):
        contact_found = True
    return contact_found


# This function returns the only required sections (such as 'contact','objective','experience','education','skills').

def generate_required_sections(sections,header,title):
    required_section = {}
    
    if check_contact_info(sections,header):
        required_section['contact'] = sections[header[0]]
    for i in title:
        value = []
        for j in title[i]:
            value.append(sections[j])
        value="".join(value)
        required_section[i]=value
    
    return required_section


# The main function (starting point of the program execution).

if __name__=='__main__':
    soup = read_docx('sample/resume.docx')
    font_size = find_font_size(soup)
    header = find_header(soup,font_size)
    
    text = soup.text
    title = title_dict(header,reserved_words)
    sections = generate_section(text,header)
    required_section = generate_required_sections(sections,header,title) 
    check_contact = check_contact_info(sections,header)
    score = score_generator(required_section)
    
#     print(sections)
#     print(header)
#     print('-------------------------------------------------------------------------')
    print('Title: ',title)
#     print('-------------------------------------------------------------------------')
    
#     print('-------------------------------------------------------------------------')
#     print('Sections: ',sections)
#     print('-------------------------------------------------------------------------')
    print('Sections Available: ',required_section.keys())
    print('-------------------------------------------------------------------------')
    for i in required_section:
        print(i)
        print(required_section[i])
        print('========================')
    print('-------------------------------------------------------------------------')
    print('Score: ',score)
#     print(check_contact)



