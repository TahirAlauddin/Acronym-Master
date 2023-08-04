import spacy
from spacy.matcher import Matcher
from fuzzywuzzy import fuzz
import re

# Define function to multiply last character in a string if it's followed by a digit.
# For example, 'A3' becomes 'AAA', 'B2' becomes 'BB' etc.
def replicate_last_char(input_str):
    # Search for a pattern that ends with a letter followed by one or more digits
    match = re.search(r"([a-zA-Z])(\d+)$", input_str)
    if match:
        # If pattern is found, extract the letter and digit
        items = match.groups()
        # Return the string with last digit characters replaced by the replicated last letter
        return input_str[:-len(items[1])] + items[0] * (int(items[1]) - 1)

# Define function to remove specific symbols from a text
def remove_symbols(text):
    # Define symbols to be removed
    symbols = ",()’"
    # For each symbol, replace it with an empty string in the text
    for symbol in symbols:
        text = text.replace(symbol, "")
    # Return text after removing symbols
    return text


# Define function to extract potential abbreviations from a SpaCy document
def extract_abbreviations(doc, matcher):
    # Define a set to store potential abbreviations
    potential_abbreviations = set()

    # Apply the matcher to the document
    matches = matcher(doc)
    # For each match
    for match_id, start, end in matches:
        # Extract the matching text
        potential_abbreviation = doc[start:end].text
        # If the abbreviation does not contain a space, add it to the set
        if " " not in potential_abbreviation:
            potential_abbreviations.add(potential_abbreviation)

    # Return the set of potential abbreviations
    return potential_abbreviations



# Function to get candidate expansions for a given abbreviation from the document
def get_candidate_expansions(abbreviation, doc):
    # Initialize an empty list to hold the candidate expansions
    candidates = []

    # Iterate over each token in the document
    for token in doc:
        # If the token text matches the abbreviation
        if token.text == abbreviation:
            # Look for potential full forms before the abbreviation
            # Calculate start and end indices for potential full form by going 
            # a few tokens before and after the abbreviation token's index.
            start = max(0, token.i - len(abbreviation) - 1)
            end = max(0, token.i + len(abbreviation) + 1)
            # Append potential full forms from before the abbreviation to candidates list
            candidates.append(doc[start:token.i])
            # Append potential full forms from after the abbreviation to candidates list
            candidates.append(doc[token.i:end])

    # Return the list of candidate expansions
    return candidates

# Function to check if a sublist exists in a list at a certain position
def sublist_exists(uppercase_chars, abbr_chars, pos, matched_indx):
    # If the sublist is empty, return the current position and matched index
    if len(uppercase_chars) == 0:
        return pos, matched_indx
    
    # If position is -1, check if the sublist equals the end of the list
    if pos == -1:
        start_pos = len(abbr_chars) - len(uppercase_chars)
        end_pos = len(abbr_chars)
        result = abbr_chars[start_pos:end_pos] == uppercase_chars
        # If sublist equals end of list, adjust position and matched index
        if result:
            matched_indx += len(uppercase_chars)
            pos -= len(uppercase_chars)
        return pos, matched_indx

    # If position is not -1, check if sublist equals a slice of the list at the given position
    start_pos = pos - len(uppercase_chars) + 1
    end_pos = pos + 1
    result = abbr_chars[start_pos:end_pos] == uppercase_chars
    # If sublist equals the slice of the list, adjust position and matched index
    if result:
        matched_indx += len(uppercase_chars)
        pos -= len(uppercase_chars)
    return pos, matched_indx


# Function to determine if a potential full form is a valid expansion of an abbreviation
def is_full_form(abbreviation, potential_full_form, threshold):
    # Clean abbreviation by removing certain special characters
    for i in ['@', '&', "/", "\\"]:
        abbreviation = abbreviation.replace(i, "")
    # Convert the cleaned abbreviation into a list of characters
    abbr_chars = list(abbreviation)
    # Split the potential full form into a list of words
    full_form_words = str(potential_full_form).split()

    matched_indx = 0
    full_form = []
    index = -1

    # Iterate through the full form words in reverse
    for full_form_word in reversed(full_form_words):
        # Extract uppercase characters from the word
        uppercase_chars = [char for char in full_form_word if char.isupper()]
        # Check if the sublist of uppercase characters exists at a certain position in the abbreviation
        index, matched_indx = sublist_exists(uppercase_chars, abbr_chars, index, matched_indx)
        # Append the word to the list of full form words
        full_form.append(full_form_word)
        # If the abbreviation is fully matched, break the loop
        if matched_indx == len(abbr_chars):
            break

    # If no full form words were found, return None
    if not full_form:
        return None

    # Reconstruct the full form string by reversing the list of words and joining them with spaces
    full_form = " ".join(reversed(full_form))

    # Capitalize the first letter of each word in the full form, excluding "and"
    candidate_caps = ' '.join(word[0].upper() + word[1:] if word != 'and' else word for word in full_form.split())
    # Keep only the uppercase characters in the candidate full form
    candidate_caps = "".join([char for char in candidate_caps if char.isupper()])

    # If the abbreviation contains a number, replicate the last character of the abbreviation
    if any(char.isdigit() for char in abbreviation):
        abbreviation = replicate_last_char(abbreviation)

    # Calculate the fuzzy match score between the abbreviation and the candidate full form
    match_score = fuzz.ratio(abbreviation, candidate_caps)
    # If the match score exceeds the threshold, return the full form
    if match_score >= threshold:
        return full_form

    # If no full form passed the match score threshold, return None
    return None


# Function to get the definition of abbreviations in a document
def get_abbreviations_definition(doc, matcher, threshold):
    # Get a set of potential abbreviations from the document
    potential_abbreviations = extract_abbreviations(doc, matcher)
    # Initialize an empty dictionary to hold the candidate expansions for each abbreviation
    abbreviations = dict()

    # For each potential abbreviation, get its candidate expansions from the document
    abbreviations_full_forms = dict()
    for potential_abbreviation in potential_abbreviations:
        abbreviations[potential_abbreviation] = get_candidate_expansions(potential_abbreviation, doc)

    # For each abbreviation and its candidate expansions
    for abbreviation in abbreviations:
        # For each candidate expansion of the abbreviation
        for potential_abbreviation in abbreviations[abbreviation]:
            # Check if the candidate expansion is a valid full form of the abbreviation
            full_form = is_full_form(abbreviation, potential_abbreviation, threshold)
            # If a valid full form is found, add it to the dictionary of full forms
            if full_form:
                abbreviations_full_forms[abbreviation] = full_form
        
    # Return the dictionary of abbreviations and their full forms
    return abbreviations_full_forms

def get_abbreviations(text, signal):
    # Load the Spacy English model
    nlp = spacy.load("en_core_web_sm")
    
    # Emit signal
    signal.emit(40)
    
    # Initialize a Matcher with the shared vocabulary
    matcher = Matcher(nlp.vocab)

    # Define a pattern to match capitalized words
    abbreviation_pattern1 = [{"IS_UPPER": True}]

    # Define a pattern to match abbreviations like 'Ph.D.'
    abbreviation_pattern2 = [{"TEXT": {"REGEX": r'\b[A-Za-z]+\.[A-Za-z\.]*'}}]

    # Add the patterns to the matcher
    matcher.add("Abbreviation1", [abbreviation_pattern1])
    matcher.add("Abbreviation2", [abbreviation_pattern2])

    # Emit signal
    signal.emit(50)
    
    # Remove certain symbols from the text
    text = remove_symbols(text)

    # Emit signal
    signal.emit(60)
    
    # Process the text with the Spacy model
    doc = nlp(text)

    # Emit signal
    signal.emit(65)
    
    # Get a dictionary of abbreviations and their full forms from the processed text
    dicto = get_abbreviations_definition(doc, matcher, 80)
    # Print each abbreviation and its full form

    # Emit signal
    signal.emit(70)
    
    result = {}
    for i in dicto:
        result[i] = dicto[i]
        # print(i, dicto[i])
    
    # Emit signal
    signal.emit(80)
    
    return result
    

def main():
    # Define the input text
    text = """To assist Commander, US Fleet Forces Command’s (USFFC) cybersecurity initiatives that support training and equipping combat forces, executing command and control (C2) activities, performing operational planning, and executing joint missions, XYZ, Inc. (XYZ) is pleased to respond to USFFC’s solicitation for Navy Risk Management Support. As the incumbent contractor providing these services to USFFC today, XYZ is uniquely positioned to continue our support to the command in its Risk Management Framework (RMF) Assessment and Authorization (A&A) efforts because of our extensive experience across the Department of the Navy’s (DON) major Cybersecurity programs and efforts of Ph.D., BSCS and M.O.O.S Dr."""

    get_abbreviations(text)

if __name__ == "__main__":
    main()