import pandas as pd
import os
from urllib.request import urlopen
from bs4 import BeautifulSoup
import nltk
nltk.download('stopwords')
import requests
import re


stop_words = set()

# Specify the path to the stop words directory
stop_words_dir = 'stop_words'

# Iterate over all files in the stop words directory
for filename in os.listdir(stop_words_dir):
    filepath = os.path.join(stop_words_dir, filename)
    
    # Read the stop words from the file
    with open(filepath, 'r') as f:
        for line in f:
            # Ignore text after '|'
            word = line.split('|')[0].strip().lower()
            stop_words.add(word)

# Download the required nltk packages
nltk.download('punkt')

# Load positive and negative words
positive_words = set(open('positive_words.txt').read().splitlines())
negative_words = set(open('negative_words.txt').read().splitlines())

# Load input file
input_file = pd.read_excel('input.xlsx')

# Create a new directory for the articles
if not os.path.exists('articles'):
    os.makedirs('articles')

# Create output dataframe
output = pd.DataFrame(columns=['URL_ID', 'URL', 'Positive Score', 'Negative Score', 'Polarity Score', 'Subjectivity Score', 'Average Sentence Length', 'Average Number of Words Per Sentence', 'Percentage of Complex Words', 'Complex Word Count', 'Fog Index','Word Count','Syllable Per Word','Personal Pronouns','Average Word Length'])

# Loop through each URL in the input file and extract the article text and title
for i in range(len(input_file)):
    url_id = input_file['URL_ID'][i]
    url = input_file['URL'][i]
    print(f"Processing {url}")
    
    # Load the webpage and extract the article text
    res = requests.get(url)
    soup = BeautifulSoup(res.content, 'html.parser')
    article_title = ''
    if soup.find('h1') is not None:
        article_title = soup.find('h1').text.strip()
   
    article_text = ''
    for p in soup.find_all('p'):
        article_text += p.text.strip() + '\n\n\n'

    
    # Save the article as a text file with the URL_ID as the filename         
    filename = f'articles/{url_id}.txt'
    with open(filename, 'w', encoding='utf-8') as file:         #storing text into string format with (context manager)
        file.write(f'{article_title}\n\n\n\n{article_text}')

    # Print the filename for confirmation
    print(f'Saved {filename}')

    
    # Calculate positive score
    words = nltk.word_tokenize(article_text.lower())

    #remove stop word
    words = [word for word in words if word not in stop_words]

    # Initialize the Positive Score to 0
    pos_score = 0

    # Calculate the Positive Score
    for word in words:
        if word in positive_words:
            pos_score += 1
    
    # Initialize the Negative Score to 0
    neg_score = 0

    # Calculate the Negative Score
    for word in words:
        if word in negative_words:
            neg_score -= 1

    # Multiply the Negative Score by -1 to make it a positive number
    neg_score *= -1
        
    # Calculate polarity score
    polarity_score = (pos_score - neg_score) / ((pos_score + neg_score) + 0.000001) 
    
    # Calculate the Total Number of Words
    total_words = len(words)

    # Calculate the Subjectivity Score
    subjectivity_score = (pos_score + neg_score) / (total_words + 0.000001)

    
    # Calculate average sentence length
    sentences = nltk.sent_tokenize(article_text)
    num_sentences = len(sentences)
    num_words = len(words)
    avg_sentence_length = num_words / num_sentences if num_sentences != 0 else 0
    
    # Calculate average number of words per sentence
    avg_words_per_sentence = num_words / num_sentences if num_sentences != 0 else 0
    
    # Calculate percentage of complex words
    complex_words = [word for word in words if len(word) > 2 and word not in nltk.corpus.stopwords.words('english')] 
    percent_complex_words = len(complex_words) / len(words) if len(words) != 0 else 0
    

    # Calculate average word length
    total_word_length = sum(len(word) for word in words)
    average_word_length = total_word_length / len(words) if len(words) != 0 else 0

    # Calculate word count after cleaning text
    cleaned_words = [word for word in words if word not in nltk.corpus.stopwords.words('english') and word.isalpha()]
    word_count = len(cleaned_words)


    # Calculate Syllable Count Per Word
    def count_syllables(word):              
        vowels = 'aeiouy'
        count = 0
        if word[0] in vowels:
            count += 1
        for index in range(1, len(word)):
            if word[index] in vowels and word[index - 1] not in vowels:     #This is because a vowel sound in a word usually marks the beginning of a new syllable
                count += 1
                if word.endswith('es') or word.endswith('ed'):
                    count -= 1
        if count == 0:
            count = 1
        return count

    syllable_count_per_word = [count_syllables(word) for word in cleaned_words]


    # Calculate complex word count
    complex_word_count = len(complex_words)
    
    # Calculate fog index
    fog_index = 0.4 * (avg_sentence_length + percent_complex_words)

    

    # Define the personal pronouns regex pattern                                              
    personal_pronouns_pattern = re.compile(r'\b(i|we|my|ours|us)\b', flags=re.IGNORECASE)

    # Calculate personal pronoun count
    personal_pronoun_count = len(re.findall(personal_pronouns_pattern, article_text))


    
    # Add the results to the output dataframe
    output = output.append({'URL_ID': url_id, 'URL': url, 'Positive Score': pos_score, 'Negative Score': neg_score, 
                            'Polarity Score': polarity_score, 'Subjectivity Score': subjectivity_score, 
                            'Average Sentence Length': avg_sentence_length, 'Average Number of Words Per Sentence': avg_words_per_sentence,
                            'Percentage of Complex Words':percent_complex_words, 'Complex Word Count':complex_word_count, 'Fog Index': fog_index,
                            'Word Count':word_count,'Syllable Per Word':syllable_count_per_word,'Personal Pronouns':personal_pronoun_count,'Average Word Length':average_word_length },ignore_index=True)

# Save the output dataframe to a new excel file
output.to_excel('output.xlsx', index=False)
