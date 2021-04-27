#!/usr/bin/env python
# coding: utf-8

# In[87]:


from docx2python import docx2python
import re
from bs4 import BeautifulSoup as bs

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


# In[104]:


reserved_words = {'contact':['contact','name','email','linkedin'],'objective':['objective','executive summary'],
                  'education':['education','qualification']
                  ,'experience':['exp.','experience','work summary','demonstrated'],'skill':['skill','expertise']}


# In[105]:


def read_docx(path):
    file = docx2python(path,html=True)
    soup = bs(file.text)
    
    return soup


# In[106]:


def find_font_size(soup):
    fonts = soup.find_all('font')
    font_size = set()
    for font in fonts:
        font_size.add(font['size'])
        
    font_size_list = list(font_size)
    font_size_list.sort(reverse=True)
    
    return font_size_list


# In[107]:


def find_header(soup,font_size_list):
    header = []
    for i in font_size_list[0:5]:
        body = soup.find_all('font',attrs={'size':i})
        for bold in body:
            bold_text = bold.find_all('b')
            for x in bold_text:
                header.append(x.text)
    
    return header
        


# In[108]:


def check(string,substring):
    found = False
    for i in range(0,len(substring)):
        if string.lower().find(substring[i])!=-1:
            found = True
    return found


# In[109]:


def title_dict(header,reserved_words):
    title = {}
    

    for a in range(0,len(header)):
        for b in reserved_words:
            d = check(header[a].lower(),reserved_words[b])
            if d == True:
                if b in title:
                    title[b].append(header[a])
                else:
                    title[b] = [header[a]]
    return title


# In[110]:


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


# In[111]:


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
    


# In[112]:


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


EMAIL_REG = re.compile(r'[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+')
def extract_emails(input_text):
    return re.findall(EMAIL_REG, input_text)


LINKED_REG = re.compile(r'https\:\/\/www\.linkedin+\.[a-zA-Z0-9/~\-_,&=\?\.;]+[^\.,\s<]')

def extract_linkedin(input_text):
    return re.findall(LINKED_REG, input_text)


# In[113]:


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


# In[114]:


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


# In[139]:


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
#     print('Title: ',title)
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


# In[ ]:




