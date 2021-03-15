from pdfminer.high_level import extract_text
import docx2txt
import ntpath
import nltk
import requests
from wordcloud import WordCloud, STOPWORDS 
import matplotlib.pyplot as plt 
import pandas as pd
from collections import Counter

# nltk.download('stopwords')
df = pd.read_csv('skills_new.csv')
SKILLS_DB = df.values

def extract_text_from_file(path):
    file_ext = ntpath.basename(path)
    filename = file_ext.split('.')
    if filename[1]=='pdf':
        return extract_text(path)
    if filename[1]=='docx':
        txt = docx2txt.process(path)
        if txt:
            return txt.replace('\t', ' ')
        return None

    
def extract_skills(input_text):
    stop_words = set(nltk.corpus.stopwords.words('english'))
    word_tokens = nltk.tokenize.word_tokenize(input_text)

    # remove the stop words
    filtered_tokens = [w for w in word_tokens if w not in stop_words]

    # remove the punctuation
    filtered_tokens = [w for w in word_tokens if w.isalpha()]

    # generate bigrams and trigrams (such as artificial intelligence)
    bigrams_trigrams = list(map(' '.join, nltk.everygrams(filtered_tokens, 2, 3)))

    # we create a list to keep the results in.
    found_skills = []

    # we search for each token in our skills database
    for token in filtered_tokens:
        if token.lower() in SKILLS_DB:
            found_skills.append(token.lower())

    # we search for each bigram and trigram in our skills database
    for ngram in bigrams_trigrams:
        if ngram.lower() in SKILLS_DB:
            found_skills.append(ngram.lower())
            

    return found_skills

def generate_wordcloud(skills):
    comment_words = Counter(skills)
    stopwords = set(STOPWORDS)
#     print(comment_words)
    
    wordcloud = WordCloud(width = 800, height = 600, 
                          stopwords = stopwords,
                          background_color ='white',
                         min_font_size = 6,font_step=1).generate_from_frequencies(comment_words)
    
    plt.figure(figsize = (10, 10), facecolor = None) 
    plt.imshow(wordcloud,interpolation='bilinear') 
    plt.axis("off") 
    plt.tight_layout(pad = 0) 
  
    plt.show()
##################################################################################################

if __name__ == '__main__':
    text = extract_text_from_file('hitesh_resume.pdf')
    skills = extract_skills(text)
    wordcloud = generate_wordcloud(skills)
