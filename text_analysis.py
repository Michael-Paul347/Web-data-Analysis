import pandas as pd
import requests
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
import re
import os

# Ensure necessary NLTK data packages are downloaded
nltk.download('punkt')
nltk.download('stopwords')

# Load stop words into a set for quick lookup
stopwords = set(nltk.corpus.stopwords.words('english'))
base_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the stopwords directory path
stopwords_dir = os.path.abspath(os.path.join(base_dir, 'StopWords'))

for stopword_file in os.listdir(stopwords_dir):
    file_path = os.path.join(stopwords_dir, stopword_file)
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
        for word in file.read().split():
            stopwords.update(file.read().splitlines())

# Function to load words from a file into a dictionary, excluding stop words
positive_words = {'positive': []}
negative_words = {'negative': []}

# Construct the positive and negative words file paths
positive_words_path = os.path.abspath(os.path.join(base_dir, 'MasterDictionary', 'positive-words.txt'))
negative_words_path = os.path.abspath(os.path.join(base_dir, 'MasterDictionary', 'negative-words.txt'))

# Read and update the positive words list in the dictionary
with open(positive_words_path, 'r', encoding='utf-8', errors='ignore') as file:
    for word in file.read().splitlines():
        if word and word not in stopwords:
            positive_words['positive'].append(word)

# Read and update the negative words list in the dictionary
with open(negative_words_path, 'r', encoding='utf-8', errors='ignore') as file:
    for word in file.read().splitlines():
        if word and word not in stopwords:
            negative_words['negative'].append(word)

def extract_article_text(url):
    # Send a request to the URL
    response = requests.get(url)
    
    # Check if the request was successful
    if response.status_code != 200:
        raise Exception(f"Failed to fetch the page. Status code: {response.status_code}")
    
    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(response.content, 'html.parser')
    
    title = soup.find('h1').get_text(strip=True) + '\n\n' if soup.find('h1') else ''

    # Find the target div
    content_div = soup.find('div', class_='td-post-content tagdiv-type')
    
    # Extract the required text from the specified elements
    extracted_text = ""
    
    # Extract all the <h1>, <p>, and <ul> tags
    for element in content_div.find_all(['h1', 'p', 'ul','ol']):
        if element.name == 'ul' or element.name == 'ol' :
            for li in element.find_all('li'):
               if  li.get_text(separator=' ', strip=True) not in extracted_text:
                extracted_text += "- " + li.get_text(separator=' ', strip=True) + "\n"
               else:
                   
                extracted_text += "" 
                   
            extracted_text += "\n"  # Add an extra newline after a list for separation
        else:
            extracted_text += element.get_text(separator=' ', strip=True) + "\n\n"
    
    # Remove duplicate entries
    lines = extracted_text.split('\n')
    unique_lines = []
    for line in lines:
        if line not in unique_lines:
            unique_lines.append(line)
    
    return title  + '\n\n'.join(unique_lines).strip()  # Remove trailing newlines

def clean_text(text):
    words = word_tokenize(text)
    return [word for word in words if word.isalnum() and word.lower() not in stopwords]

def calculate_positive_score(words):
    pos = sum(1 for word in words if word in positive_words['positive'])
    return round(pos, 3)

def calculate_negative_score(words):
    neg = sum(1 for word in words if word in negative_words['negative'])
    return round(neg, 3)

def calculate_polarity_score(positive_score, negative_score):
    pol = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    return round(pol, 3)

def calculate_subjectivity_score(positive_score, negative_score, total_words):
    sub = (positive_score + negative_score) / (total_words + 0.000001)
    return round(sub, 3)

def calculate_average_sentence_length(text):
    sentences = sent_tokenize(text)
    total_words = len(word_tokenize(text))
    avg_sentence_len = total_words / len(sentences)
    return round(avg_sentence_len, 2)

def calculate_percentage_complex_words(words):
    complex_words = [word for word in words if count_syllables(word) > 2]
    complex_per = (len(complex_words) / len(words)) * 100
    return round(complex_per, 4)

def calculate_fog_index(average_sentence_length, percentage_complex_words):
    fog_i = 0.4 * (average_sentence_length + percentage_complex_words)
    return round(fog_i, 2)

def count_syllables(word):
    word = word.lower()
    vowels = "aeiou"
    count = 0
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("es") or word.endswith("ed"):
        count -= 1
    if count == 0:
        count += 1
    return count

def count_personal_pronouns(text):
    pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, re.I)
    return len(pronouns)

def calculate_average_word_length(words):
    avg_word_len = sum(len(word) for word in words) / len(words)
    return round(avg_word_len, 2)

# Read input file
input_df = pd.read_excel('Input.xlsx')
results = []

# Process each URL
for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    try:
        article_text = extract_article_text(url)
        cleaned_words = clean_text(article_text)
        positive_score = calculate_positive_score(cleaned_words)
        negative_score = calculate_negative_score(cleaned_words)
        polarity_score = calculate_polarity_score(positive_score, negative_score)
        subjectivity_score = calculate_subjectivity_score(positive_score, negative_score, len(cleaned_words))
        avg_sentence_length = calculate_average_sentence_length(article_text)
        percentage_complex_words = calculate_percentage_complex_words(cleaned_words)
        fog_index = calculate_fog_index(avg_sentence_length, percentage_complex_words)
        complex_word_count = len([word for word in cleaned_words if count_syllables(word) > 2])
        word_count = len(cleaned_words)
        syllable_per_word = round(sum(count_syllables(word) for word in cleaned_words) / len(cleaned_words), 2)
        personal_pronouns = count_personal_pronouns(article_text)
        avg_word_length = calculate_average_word_length(cleaned_words)

        results.append([
            url_id, url, positive_score, negative_score, polarity_score, subjectivity_score,
            avg_sentence_length, percentage_complex_words, fog_index, complex_word_count,
            word_count, syllable_per_word, personal_pronouns, avg_word_length
        ])
    except Exception as e:
        print(f"Failed to process URL {url}: {e}")

# Create DataFrame for output
output_columns = [
    'URL_ID', 'URL', 'Positive Score', 'Negative Score', 'Polarity Score', 'Subjectivity Score',
    'Avg Sentence Length', 'Percentage of Complex Words', 'Fog Index', 'Complex Word Count',
    'Word Count', 'Syllable Per Word', 'Personal Pronouns', 'Avg Word Length'
]
output_df = pd.DataFrame(results, columns=output_columns)

# Save output to Excel file
output_df.to_excel('Output.xlsx', index=False)

# Save results to text files
os.makedirs('ExtractedArticles', exist_ok=True)
for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    try:
        with open(f'ExtractedArticles/{url_id}.txt', 'w', encoding='utf-8') as f:
            f.write(extract_article_text(row['URL']))
        print(f"Extracted and analyzed content from {url}")
    except Exception as e:
        print(f"Failed to extract content from {url}: {e}")

print("Analysis completed successfully and results saved to Output.xlsx")