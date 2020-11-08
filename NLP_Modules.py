import gensim
import numpy as np
#nltk.download('punkt')
from nltk.tokenize import word_tokenize, sent_tokenize, LineTokenizer

def get_text_similiarity(ref_text, query_text, dummy_sentence):
    ''' Setup reference text string statistic '''
    sentences_tokens = LineTokenizer().tokenize(ref_text)  # tokenize sentences by new line
    unique_sentences_tokens = np.unique(sentences_tokens).tolist()  # remove duplicate tokenized sentences

    if len(unique_sentences_tokens) == 1:
        unique_sentences_tokens.append(dummy_sentence)

    # convert every words [as each item elements in list] in each sentences [as list array] into list object
    create_sentences_list = [[w.lower() for w in word_tokenize(word)] for word in unique_sentences_tokens]

    # assign each unique words of a List array to dictionary unique id
    unique_word_dictionary = gensim.corpora.Dictionary(create_sentences_list)

    # assign frequency of occurence of each words for each sentences as weightage by dictionary unique id
    corpus = [unique_word_dictionary.doc2bow(sentence) for sentence in create_sentences_list]

    # calculate TF*IDF scores, the higher the score, the rarer the words appears in the whole context
    # formula: Wt,d = TFt,d log (N/DFt)
    # TFt,d is the number of occurrences of t in document d
    # DFt is the number of documents containing the term t
    # N is the total number of documents [sentences] in the corpus [Corpus is a large collection of text]
    tf_idf = gensim.models.TfidfModel(corpus)

    '''quantify % similiarity of a text string to the reference text string '''
    sims = gensim.similarities.Similarity('/Users/YewLoung/PycharmProjects/SE_Datasheet_Scrape/Index_Matrix',
                                          corpus=tf_idf[corpus],
                                          num_features=len(unique_word_dictionary))

    query_sentences_tokens = LineTokenizer().tokenize(query_text)  # tokenize sentences by new line
    query_unique_sentences_tokens = np.unique(query_sentences_tokens).tolist()  # remove duplicate tokenized sentences

    for word in query_unique_sentences_tokens:
        # convert every words [as each item elements in list] in each sentences [list array] into list object
        query_sentence = [w.lower() for w in word_tokenize(word)]

        # assign each unique words to dictionary unique id
        query_sentence_bow = unique_word_dictionary.doc2bow(query_sentence)

    # perform a similarity query against the corpus
    query_sentence_tf_idf = tf_idf[query_sentence_bow]

    sum_of_sims = (np.sum(sims[query_sentence_tf_idf], dtype=np.float32))
    percentage_of_similarity = round(float((sum_of_sims / len(create_sentences_list)) * 100))

    return percentage_of_similarity


#print(get_text_similiarity('Cadmiumfree', 'Cadmiumfree'))

