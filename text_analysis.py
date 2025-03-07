import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords

nltk.download('punkt')
nltk.download('stopwords')

# Load positive and negative words
with open(r"C:\Users\messi\Desktop\text_analysis.py\positive-words.txt", "r") as f:
    positive_words = set(f.read().split())

with open(r"C:\Users\messi\Desktop\text_analysis.py\negative-words.txt", "r") as f:
    negative_words = set(f.read().split())

stop_words = set(stopwords.words('english'))

# Function to extract article text
def extract_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, "html.parser")
        title = soup.find('title').text.strip()
        paragraphs = soup.find_all('p')
        article_text = " ".join([p.text for p in paragraphs])
        return title, article_text
    except:
        return None, None

# Function to perform text analysis
def analyze_text(text):
    words = word_tokenize(text)
    words = [word.lower() for word in words if word.isalpha() and word not in stop_words]
    sentences = sent_tokenize(text)
    
    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(words) + 0.000001)
    
    avg_sentence_length = len(words) / len(sentences)
    complex_word_count = sum(1 for word in words if len(re.findall(r'[aeiouAEIOU]', word)) > 2)
    percentage_complex_words = complex_word_count / len(words)
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    avg_words_per_sentence = len(words) / len(sentences)
    syllables_per_word = sum(len(re.findall(r'[aeiouAEIOU]', word)) for word in words) / len(words)
    personal_pronouns = len(re.findall(r'\b(I|we|my|ours|us)\b', text, re.IGNORECASE))
    avg_word_length = sum(len(word) for word in words) / len(words)
    
    return [positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length,
            percentage_complex_words, fog_index, avg_words_per_sentence, complex_word_count, len(words),
            syllables_per_word, personal_pronouns, avg_word_length]

# Load input data
input_data = pd.read_excel("Input.xlsx")
output_columns = ["URL_ID", "URL", "POSITIVE SCORE", "NEGATIVE SCORE", "POLARITY SCORE", "SUBJECTIVITY SCORE",
                  "AVG SENTENCE LENGTH", "PERCENTAGE OF COMPLEX WORDS", "FOG INDEX", "AVG NUMBER OF WORDS PER SENTENCE",
                  "COMPLEX WORD COUNT", "WORD COUNT", "SYLLABLE PER WORD", "PERSONAL PRONOUNS", "AVG WORD LENGTH"]
output_data = []

for index, row in input_data.iterrows():
    if index == 0:  # Only process the first URL
        url_id, url = row['URL_ID'], row['URL']
        title, text = extract_text(url)
        if text:
            with open(f"{url_id}.txt", "w", encoding="utf-8") as f:
                f.write(title + "\n" + text)
            metrics = analyze_text(text)
            output_data.append([url_id, url] + metrics)
        else:
            output_data.append([url_id, url] + [None] * 13)

# Save output
output_df = pd.DataFrame(output_data, columns=output_columns)
output_df.to_excel("Output.xlsx", index=False)
print("Processing complete. Output saved to Output.xlsx.")
