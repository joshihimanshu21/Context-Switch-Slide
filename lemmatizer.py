from nltk.stem import WordNetLemmatizer
import warnings
warnings.filterwarnings('ignore')
from pptx import Presentation
from nltk.corpus import wordnet
from nltk.tokenize import word_tokenize 
from nltk.stem import WordNetLemmatizer
import nltk 
from nltk.corpus import stopwords 
from autocorrect import Speller


def get_wordnet_pos(treebank_tag):

    if treebank_tag.startswith('J'):
        return wordnet.ADJ
    elif treebank_tag.startswith('V'):
        return wordnet.VERB
    elif treebank_tag.startswith('N'):
        return wordnet.NOUN
    elif treebank_tag.startswith('R'):
        return wordnet.ADV
    else:
        return None # for easy if-statement

def func(txt):  
    tokenized = txt
    stop_words = set(stopwords.words('english'))
    wordslist = nltk.word_tokenize(tokenized)
    tagged = nltk.pos_tag(wordslist)
    spell = Speller(lang='en')
    words = [(spell(w),tag) for w,tag in tagged if not w in stop_words]  
    return words

def preprocessing(sentence):
    sentence = sentence.lower()
    sentence = func(str(sentence))
    temp = []
    l = WordNetLemmatizer()
    for word,tag in sentence:
        wntag = get_wordnet_pos(tag)
        if wntag is not None:
            word = l.lemmatize(word,pos=wntag)
        temp.append(word)
    sentence = temp
    return sentence

if __name__ == "__main__":
    lemmatizer = WordNetLemmatizer()

    sentence = input("Enter a sentence:")

    print(preprocessing(sentence))

# print(lemmatizer.lemmatize("better",pos='a'))


