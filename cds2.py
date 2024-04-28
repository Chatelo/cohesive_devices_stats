import glob
import docx
import os
import re
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import nltk
import pandas as pd
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk import pos_tag
from termcolor import colored
from nltk.collocations import BigramCollocationFinder

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
  word_counts = {}
  for word in found:
    if word in word_counts:
      word_counts[word] += 1
    else:
      word_counts[word] = 1
  return word_counts

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

    # Define patterns and corresponding function calls
    patterns_functions = {
        "Personal Pronouns": personal_pronouns_pattern,
        "Possessive Pronouns": possessive_pronouns_pattern,
        "Interrogative Pronouns": interrogative_pronouns_pattern,
        "Possessive Adjectives": possessive_adjectives_pattern,
        "Nouns": None,  # Count nouns separately
        "Demonstratives": demonstratives_pattern,
        "Comparatives": comparatives_pattern,
        "Definite Articles": definite_articles_pattern,
        "Additives": additives_pattern,
        "Adversatives": adversatives_pattern,
        "Causals": causals_pattern,
        "Temporals": temporals_pattern,
        "Continuatives": continuatives_pattern,
        "Repetitions": repetition_pattern,
        # Add other patterns as needed
    }

    # Initialize dictionary to store counts and words for each pattern
    results = {pattern_name: {"Count": 0, "Words": [], "Word Frequencies": {}} for pattern_name in patterns_functions.keys()}
    total_pattern_count = 0  # Initialize a variable to store total count

    # Process each pattern
    for pattern_name, pattern in patterns_functions.items():
        if pattern:
            word_counts = count_pattern(text, pattern)  # Use modified function
            total_pattern_count = sum(word_counts.values())  # Update total count
            results[pattern_name]["Count"] = total_pattern_count
            results[pattern_name]["Words"] = list(word_counts.keys())
            results[pattern_name]["Word Frequencies"] = word_counts

    # Count nouns separately
    nouns_count, nouns = count_nouns(text)
    results["Nouns"]["Count"] = nouns_count
    results["Nouns"]["Words"] = nouns
    results["Nouns"]["Counts"] = [1] * len(nouns)  # Initialize counts to 1

    # Construct and return a dictionary of results for this composition
    doc_results = {
        "File": file_path,
    }

    # Add counts and words for each pattern to the doc_results dictionary
    for pattern_name, data in results.items():
        doc_results[pattern_name + " Count"] = data["Count"]
        # Create a dictionary to store the count of each word
        # word_count_dict = {word: count for word, count in zip(data["Words"], data["Counts"])}
        # Join words and counts for words occurring more than once
        words_with_counts = [f"{word}({count})" if count > 1 else word for word, count in data["Word Frequencies"].items()]
        doc_results[pattern_name + " Words"] = ", ".join(words_with_counts)

    return doc_results




# Main function to process all compositions in a directory
def process_all_compositions(directory_path):
    print(f"Processing files in: {directory_path}")
    results_list = []
    for file_path in glob.glob(f"{directory_path}/*.docx"):
        try:
            doc_results = process_composition(file_path)
            results_list.append(doc_results)
        except Exception as e:
            print(f"An error occurred while processing {file_path}: {e}")
    
    # Create a DataFrame from the results list
    results_df = pd.DataFrame(results_list)
    # Extract just the file name without the directory path or extension
    results_df['File'] = results_df['File'].apply(lambda x: os.path.splitext(os.path.basename(x))[0])
    
    # Export the DataFrame to an Excel file
    excel_file_path = "results.xlsx"
    results_df.to_excel(excel_file_path, index=False)
    print(f"Results exported to {excel_file_path}")


# Example usage
directory_path = "/path/to/your/file(s)-directory/"
process_all_compositions(directory_path)
