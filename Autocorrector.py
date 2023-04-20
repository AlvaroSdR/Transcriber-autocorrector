from docx import Document
from textblob import TextBlob
from nltk.tokenize import sent_tokenize
import string

# Load the .docx file
doc = Document('text.docx')

# Loop through all paragraphs in the .docx file
for paragraph in doc.paragraphs:
    # Split the paragraph into words
    words = paragraph.text.split()
    # Loop through all words in the paragraph
    for i in range(len(words)):
        # Use TextBlob to perform parts-of-speech tagging
        blob = TextBlob(words[i])
        # If the word is a noun or proper noun, skip autocorrection
        if len(blob.tags) > 0 and blob.tags[0][1] in ['NN', 'NNP']:
            continue
        # Replace the word with its corrected version
        words[i] = str(blob.correct())

    # Join the corrected words back into a paragraph
    corrected_paragraph = ' '.join(words)

    # Split the corrected paragraph into sentences
    sentences = sent_tokenize(corrected_paragraph)

    # Loop through all sentences in the paragraph
    for i in range(len(sentences)):
        # Remove leading/trailing white spaces
        sentence = sentences[i].strip()
        # Add a punctuation mark if the sentence does not end with one
        if not sentence.endswith(tuple(string.punctuation)):
            sentence += '.'
        # Replace the original sentence text with the corrected sentence text
        sentences[i] = sentence

    # Join the corrected sentences back into a paragraph
    corrected_paragraph = ' '.join(sentences)
    # Replace the original paragraph text with the corrected paragraph text
    paragraph.text = corrected_paragraph

# Save the corrected .docx file
doc.save('output.docx')
