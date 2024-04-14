import glob
import docx
import re
import nltk
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk import pos_tag
from termcolor import colored
from nltk.collocations import *

# Download necessary NLTK data
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('stopwords')
nltk.download('wordnet')  # For synonyms, antonyms, hyponyms, meronyms

# Define patterns for each sub-category of cohesive devices
personal_pronouns_pattern = r'\b(I|me|you|he|him|she|her|it|we|us|they|them)\b'
possessive_pronouns_pattern = r'\b(my|your|his|her|its|our|your|their)\b'
interrogative_pronouns_pattern = r'\b(who|whom|whose|which|what)\b'
possessive_adjectives_pattern = r'\b(my|your|his|her|its|our|your|their)\b'
demonstratives_pattern = r'\b(this|that|these|those)\b'
comparatives_pattern = r'\b(more|less|better|worse|greater|fewer)\b'
definite_articles_pattern = r'\b(the)\b'
additives_pattern = r'\b(and|also|furthermore|moreover)\b'
adversatives_pattern = r'\b(but|however|although|nevertheless)\b'
causals_pattern = r'\b(because|since|therefore|thus)\b'
temporals_pattern = r'\b(when|while|as soon as|before|after)\b'
continuatives_pattern = r'\b(then|next|besides|likewise)\b'
repetition_pattern = r'\b(\w+)\s+\1\b'  # Exact repetition

# Function to count occurrences of each pattern and return the actual words
def count_pattern(text, pattern):
    found = re.findall(pattern, text)
    # return len(found), found
    
    # Option 2: Convert dict_items to a list (if you need a list of matches)
    return len(list(found)), list(found)  # Count and return list of matches

# Function to count nouns using NLTK's POS tagger
def count_nouns(text):
    words = word_tokenize(text)
    words = [word for word in words if word.lower() not in stopwords.words('english')]
    tagged_words = pos_tag(words)
    nouns = [word for word, tag in tagged_words if tag.startswith('NN')]
    return len(nouns), nouns

# Function to count repetition (and synonyms/antonyms/hyponyms/meronyms if using WordNet)
def count_repetition(text):
    word_counts = {}  # To store counts of each unique word

    for word in word_tokenize(text.lower()):  # Convert to lowercase
        if word not in stopwords.words('english') and word != '.':
            if word in word_counts:
                word_counts[word] += 1
            else:
                word_counts[word] = 1

    repetitions = {word: count for word, count in word_counts.items() if count > 1}  # Filter repeated words

    return sum(repetitions.values()), repetitions  # Return overall count and repeated words
# Function to count occurrences of substitution using POS tags
def count_substitution(text):
    words = word_tokenize(text)
    tagged_words = pos_tag(words)
    # Define POS tags related to substitution (e.g., pronouns can substitute nouns)
    substitution_verbal_tags = ['VB', 'VBD', 'VBG', 'VBN', 'VBP', 'VBZ']  # Verbs
    substitution_clausal_tags = ['IN', 'DT']  # Prepositions, determiners
    substitution_nominal_tags = ['NN', 'NNS', 'NNP', 'NNPS']  # Nouns
    verbal_count = sum(1 for word, tag in tagged_words if tag in substitution_verbal_tags)
    clausal_count = sum(1 for word, tag in tagged_words if tag in substitution_clausal_tags)
    nominal_count = sum(1 for word, tag in tagged_words if tag in substitution_nominal_tags)
    return verbal_count, clausal_count, nominal_count


# Function to count occurrences of ellipses
def count_ellipses(text):
    # Ellipses are often context-dependent and may require manual rules or ML models
    # Here's a simple placeholder for ellipses count
    ellipses_count = text.count('...')
    # Placeholder for categorization
    return ellipses_count, 0, 0, 0  # Return counts for verbal, clausal, and nominal (0 for now)

# Function to count occurrences of reiteration and categorize them into synonyms, antonyms, hyponyms, and meronyms
def categorize_reiteration(repeated_words, text):
    # Placeholder for categorization using WordNet
    synonyms = set()
    antonyms = set()
    hyponyms = set()
    meronyms = set()
    for word in repeated_words:
        # Ensure the word is present in the document's text
        if word in text:
            synsets = nltk.corpus.wordnet.synsets(word)
            for synset in synsets:
                for lemma in synset.lemmas():
                    # Filter synonyms
                    if lemma.name() in text:
                        synonyms.add(lemma.name())
                    for antonym in lemma.antonyms():
                        # Filter antonyms
                        if antonym.name() in text:
                            antonyms.add(antonym.name())
                for hyponym in synset.hyponyms():
                    # Filter hyponyms
                    if hyponym.name() in text:
                        hyponyms.add(hyponym.name())
                for meronym in synset.part_meronyms() + synset.substance_meronyms() + synset.member_meronyms():
                    # Filter meronyms
                    if meronym.name() in text:
                        meronyms.add(meronym.name())
    return synonyms, antonyms, hyponyms, meronyms

# Modified function to count occurrences of reiteration and use the new categorization function
def count_reiteration(text):
    # This function can be complex as it involves semantic analysis
    # For simplicity, let's count just the repetition of the same word
    words = word_tokenize(text.lower())
    word_counts = nltk.FreqDist(words)
    # Find repeated words
    repeated_words = [word for word, count in word_counts.items() if count > 1]
    # Categorize reiteration into synonyms, antonyms, hyponyms, and meronyms
    synonyms, antonyms, hyponyms, meronyms = categorize_reiteration(repeated_words, text)
    return len(repeated_words), synonyms, antonyms, hyponyms, meronyms
# Function to find collocations in text

def find_collocations(text):
    tokens = word_tokenize(text)
    # Filter out periods
    tokens = [token for token in tokens if token != '.']
    finder = BigramCollocationFinder.from_words(tokens)
    # Apply frequency filter to remove infrequent collocations
    finder.apply_freq_filter(3)
    bigram_measures = nltk.collocations.BigramAssocMeasures()
    collocations = finder.nbest(bigram_measures.pmi, 10)
    
    # Count occurrences of collocations
    collocations_count = [finder.ngram_fd[collocation] for collocation in collocations]
    
    # Total count of collocations
    total_collocations_count = sum(collocations_count)
    
    return collocations, total_collocations_count



# Function to process a single composition
def process_composition(file_path):
    doc = docx.Document(file_path)
    text = "\n".join([para.text for para in doc.paragraphs])
    
    # Count occurrences of each cohesive device and collect actual words
    personal_pronouns_count, personal_pronouns = count_pattern(text, personal_pronouns_pattern)
    possessive_pronouns_count, possessive_pronouns = count_pattern(text, possessive_pronouns_pattern)
    interrogative_pronouns_count, interrogative_pronouns = count_pattern(text, interrogative_pronouns_pattern)
    possessive_adjectives_count, possessive_adjectives = count_pattern(text, possessive_adjectives_pattern)
    demonstratives_count, demonstratives = count_pattern(text, demonstratives_pattern)
    comparatives_count, comparatives = count_pattern(text, comparatives_pattern)
    definite_articles_count, definite_articles = count_pattern(text, definite_articles_pattern)
    additives_count, additives = count_pattern(text, additives_pattern)
    adversatives_count, adversatives = count_pattern(text, adversatives_pattern)
    causals_count, causals = count_pattern(text, causals_pattern)
    temporals_count, temporals = count_pattern(text, temporals_pattern)
    continuatives_count, continuatives = count_pattern(text, continuatives_pattern)
    repetition_count, repetitions = count_repetition(text)

  
    substitution_verbal_count, substitution_clausal_count, substitution_nominal_count = count_substitution(text)
    ellipses_count, ellipses_verbal_count, ellipses_clausal_count, ellipses_nominal_count = count_ellipses(text)
    reiteration_count, synonyms, antonyms, hyponyms, meronyms = count_reiteration(text)
    collocations, total_collocations_count = find_collocations(text)


    # Count nouns and collect nouns
    nouns_count, nouns = count_nouns(text)
    
    # Print the results for this composition
    print(f"Results for {file_path}:")
    # Print each device's results with repeated words printed once along with their counts
    print_devices_results("Personal Pronouns", personal_pronouns_count, personal_pronouns)
    print_devices_results("Possessive Pronouns", possessive_pronouns_count, possessive_pronouns)
    print_devices_results("Interrogative Pronouns", interrogative_pronouns_count, interrogative_pronouns)
    print_devices_results("Possessive Adjectives", possessive_adjectives_count, possessive_adjectives)
    print_devices_results("Nouns", nouns_count, nouns)
    print_devices_results("Demonstratives", demonstratives_count, demonstratives)
    print_devices_results("Comparatives", comparatives_count, comparatives)
    print_devices_results("Definite Articles", definite_articles_count, definite_articles)
    print_devices_results("Additives", additives_count, additives)
    print_devices_results("Adversatives", adversatives_count, adversatives)
    print_devices_results("Causals", causals_count, causals)
    print_devices_results("Temporals", temporals_count, temporals)
    print_devices_results("Continuatives", continuatives_count, continuatives)
    print_devices_results("Repetitions", repetition_count, repetitions.items())

    print_devices_results("Substitution - Verbal", substitution_verbal_count, [])
    print_devices_results("Substitution - Clausal", substitution_clausal_count, [])
    print_devices_results("Substitution - Nominal", substitution_nominal_count, [])

    print_devices_results("Ellipses", ellipses_count, [])
    print_devices_results("Ellipses - Verbal", ellipses_verbal_count, [])
    print_devices_results("Ellipses - Clausal", ellipses_clausal_count, [])
    print_devices_results("Ellipses - Nominal", ellipses_nominal_count, [])

    # print_devices_results("Reiteration", reiteration_count, [])
    print_devices_results("Synonyms", len(synonyms), synonyms)
    print_devices_results("Antonyms", len(antonyms), antonyms)
    print_devices_results("Hyponyms", len(hyponyms), hyponyms)
    print_devices_results("Meronyms", len(meronyms), meronyms)
    # Print collocations
    collocations, total_collocations_count = find_collocations(text)
    print(colored(f"Collocations: {total_collocations_count}", 'cyan'))  # Print total count of collocations
    for i, collocation in enumerate(collocations):
        if i < len(collocations) - 1:
            print(colored(f"{' '.join(collocation)}", 'cyan'), end=', ')
        else:
            print(colored(f"{' '.join(collocation)}", 'cyan'))

def print_devices_results(device_name, device_count, device_words):
    print(colored(f"{device_name}: {device_count}", 'blue'), end=" ")
    if device_words:
        if isinstance(device_words, set):
            unique_words = device_words
        else:
            unique_words = set(device_words)
        for word in unique_words:
            count = device_words.count(word) if isinstance(device_words, list) else 1
            print(colored(f"({word}({count}))", 'blue'), end=" ")
    print()
    




# Main function to process all compositions in a directory
def process_all_compositions(directory_path):
    print(f"Processing files in: {directory_path}")
    files_processed = 0
    for file_path in glob.glob(f"{directory_path}/*.docx"):
        try:
            process_composition(file_path)
            files_processed += 1
        except Exception as e:
            print(f"An error occurred while processing {file_path}: {e}")
    if files_processed == 0:
        print("No .docx files processed. Check the directory path and file permissions.")

# Example usage
directory_path = "path/to/directory"

process_all_compositions(directory_path)