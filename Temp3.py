import nltk
import gensim
import numpy as np
#nltk.download('punkt')
from nltk.tokenize import word_tokenize, sent_tokenize, LineTokenizer

textX = 'Mars is the fourth planet in our solar system. ' \
        'It is the second-smallest planet in the Solar System after Mercury. ' \
        'Saturn is yellow planet.'

textY = 'Mars is the fourth planet in our solar system\n' \
        'It is the second-smallest planet in the Solar System after Mercury\n' \
        'Saturn is yellow planet\nIt is the second-smallest planet in the Solar System after Mercury\n' \
        'Mars is the fourth planet in our solar system\n'

textY = 'I love reading Japanese novels. My favorite Japanese writer is Tanizaki Junichiro\n' \
        'Natsume Soseki is a well-known Japanese novelist and his Kokoro is a masterpiece\n' \
        'American modern poetry is good\n'

#query_text = 'Saturn is the sixth planet from the Sun'

query_text = 'Mars is the fourth planet in our solar system ' \
             'It is the second-smallest planet in the Solar System after Mercury ' \
             'Saturn is yellow planet But, Mars is smaller than Saturn'

query_text = 'Japan has some great novelists. Who is your favorite Japanese writer?'

query_text = 'I love reading Japanese novels. My favorite Japanese writer is Tanizaki Junichiro'

#file_docs = []
#with open('demofile.rtf') as f:
#    tokens = sent_tokenize(f.read())
#    for line in tokens:
#        file_docs.append(line)
#print("Number of documents:", len(file_docs))
#print(gen_docs)


''' Setup reference text string statistic'''
#sentences_tokens = sent_tokenize(textX)  # tokenize sentences by full stop
sentences_tokens = LineTokenizer().tokenize(textY)  # tokenize sentences by new line
unique_sentences_tokens = np.unique(sentences_tokens).tolist()  # remove duplicate tokenized sentences

if len(unique_sentences_tokens) == 1:
    unique_sentences_tokens.append('XXXXX')

# convert every words [as each item elements in list] in each sentences [as list array] into list object
gen_docs = [[w.lower() for w in word_tokenize(word)] for word in unique_sentences_tokens]

# assign each unique words of a List array to dictionary unique id
dictionary = gensim.corpora.Dictionary(gen_docs)

# assign frequency of occurence of each words for each sentences weightage by dictionary unique id
corpus = [dictionary.doc2bow(gen_doc) for gen_doc in gen_docs]

print(sentences_tokens, '\n')
print(unique_sentences_tokens, '\n')

print(gen_docs, '\n')
print(dictionary.token2id, '\n')
print(corpus, '\n')

# calculate TF*IDF scores, the higher the score, the rarer the words appears in the whole context
# formula: Wt,d = TFt,d log (N/DFt)
# TFt,d is the number of occurrences of t in document d
# DFt is the number of documents containing the term t
# N is the total number of documents [sentences] in the corpus [Corpus is a large collection of text]
tf_idf = gensim.models.TfidfModel(corpus)
for doc in tf_idf[corpus]:
    print([[dictionary[id], np.around(freq, decimals=2)] for id, freq in doc])
    print(doc)
print('\n')


''' quantify % similiarity of a text string to the reference text string '''
sims = gensim.similarities.SparseMatrixSimilarity(corpus=tf_idf[corpus], num_features=len(dictionary))

#sims = gensim.similarities.Similarity('/Users/YewLoung/PycharmProjects/SE_Datasheet_Scrape/Index_Matrix',
#                                      corpus=tf_idf[corpus],
#                                     num_features=len(dictionary))

query_sentences_tokens = LineTokenizer().tokenize(query_text)  # tokenize sentences by new line
query_unique_sentences_tokens = np.unique(query_sentences_tokens).tolist()  # remove duplicate tokenized sentences


for word in query_unique_sentences_tokens:
    # convert every words [as each item elements in list] in each sentences [list array] into list object
    query_gen_docs = [w.lower() for w in word_tokenize(word)]

    # assign each unique words to dictionary unique id
    query_doc_bow = dictionary.doc2bow(query_gen_docs)


# perform a similarity query against the corpus
query_doc_tf_idf = tf_idf[query_doc_bow]

# print(document_number, document_similarity)
print('Comparing Result:', sims[query_doc_tf_idf], '\n')

sum_of_sims = (np.sum(sims[query_doc_tf_idf], dtype=np.float32))
max_of_sims = round(np.amax(sims[query_doc_tf_idf]), 1)
print(sum_of_sims, '\n')

percentage_of_similarity = round(float((sum_of_sims / len(gen_docs)) * 100))
print(percentage_of_similarity, [sims[query_doc_tf_idf][i] for i in range(len(sims[query_doc_tf_idf]))], '\n')
print(max_of_sims)