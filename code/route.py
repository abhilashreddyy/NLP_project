
import azure
import json
import configparser
import os
import time
from _thread import start_new_thread
import common_code
import google

def analysis_sentiment_and_keyword(data,max_questions,dimensions,sheets,thread_message):

    config = configparser.ConfigParser()
    config.read('config.ini')
    empty_cells = []
    sol,keys,no_of_sol,sol1,jsontext,wordcloud_text = [],[],[],[],[],[]
    key = config['azure_params']['key']
    sentiment_serviceurl = config['azure_params']['sentiment_serviceurl']
    keywords_serviceurl = config['azure_params']['keywords_serviceurl']
    count = 0

    for i in range(max_questions) :
        no_of_sol.append(len(data[i]))

    for i in range(max_questions) :
        json_data = {}
        json_data['documents'] = [] ;
        count = 0
        for j in range(no_of_sol[i]):
            if data[i][j] is None :
                data[i][j] = "data unavailable"
                empty_cells.append((i,j))


            json_data['documents'].append({})
            json_data['documents'][j]['language'] = 'en'
            json_data['documents'][j]['id'] = count
            json_data['documents'][j]['text'] = data[i][j]
            count += 1 ;

        jsontext.append(json_data)
    tim = time.time()
    sentiment_info,keywords_info = azure.send_nlp_request_azure([sentiment_serviceurl,keywords_serviceurl],jsontext,key,thread_message)
    tim = time.time()-tim
    print(thread_message + ' : time taken for keyword and sentimet analysis : ' + str(tim))
    temp = empty_cells[:]
    for i in range(len(sentiment_info)) :
        for j in range(len(sentiment_info[i])) :
            if temp == [] :
                break
            if temp[0][0] == i and temp[0][1] == j :
                temp.pop(0)
                sentiment_info[i][j]['score'] = .5
                keywords_info[i][j]['keyPhrases'] = []



    azure.output_exel_with_question(sentiment_info,max_questions,dimensions,sheets,"sentiment",thread_message)
    azure.output_exel_with_question(keywords_info,max_questions,dimensions,sheets,"keyword",thread_message)
    azure.generate_piechart_for_azure_sentiment(sentiment_info,thread_message)

    for word in keywords_info :
        string = ''
        for word1 in word :
            for word2 in word1['keyPhrases'] :
                string = string + word2 + " "
        wordcloud_text.append(string)

    common_code.generate_word_cloud(wordcloud_text,"keywordwise",thread_message)

    print('Started extracting most relevant keywords')

    gradient = int(config['related_word_extraction_params']['gradient'])
    treshold = float(config['related_word_extraction_params']['treshold'])
    simplified_keywords_info = []
    question = [['environment','technical','security','passion','intelligence','customer','satisfaction'],['environment','technical','security','passion','intelligence','customer','satisfaction'],['environment','technical','security','passion','intelligence','customer','satisfaction'],['environment','technical','security','passion','intelligence','customer','satisfaction']]
    for i,word1 in enumerate(keywords_info) :
        simplified_keywords_info.append([])
        for word2 in word1:
            for word3 in word2['keyPhrases'] :
                simplified_keywords_info[i].append(word3)
    print
    q_wordcloud,extra_keys = azure.question_related_keys_extraction(simplified_keywords_info,question,gradient,treshold,thread_message)

    common_code.generate_word_cloud(q_wordcloud,"including_question",thread_message)
    print('Completed extracting most relevant keywords')


def grouping_topic(data,max_questions,thread_message) :

    config = configparser.ConfigParser()
    question_nos = []
    config.read('config.ini')
    wordcloud_text = []
    key = config['azure_params']['key']
    topicgrouping_serviceurl = config['azure_params']['topicgrouping_serviceurl']
    count = 0
    for i,question in enumerate(data) :
        #print(len(question))
        if len(question) > 100 :
            question_nos.append(i)
        else :
            print("cannot send question " ,i+1, " for analysis due to insufficent data . minimum 100 answers are required for analysis")
    print('Started grouping topics')
    response = azure.send_grouping_azure_req([data[no] for no in question_nos],topicgrouping_serviceurl,key,thread_message)
    for i in response :
        string = ''
        for j in i :
            string = string + j['keyPhrase'] + " "
        wordcloud_text.append(string)
    if wordcloud_text != [] :
        common_code.generate_word_cloud(wordcloud_text,"topicwise",thread_message,eligible_qts = question_nos)
    print('sucessfully generated word cloud from grouped topics')

def most_rel(data,value,thread_message) :
    print("started processing most relevant answers")
    most_rel = azure.extract_most_relevant(data,value,thread_message)
    azure.write_to_file(most_rel,data,thread_message)
    print("sucessfully completed extraction of most relevant answers")


def route_function() :
    data,max_questions,dimensions,sheets = common_code.read_questions()
    value = [['environment','technical','security','passion','intelligence','customer','satisfaction'],['friendly','together','understanding','attitude'],['resource','time'],['clear','understanding','undertaking','review']]
    #try :
        #start_new_thread(most_rel,(data,value,"azure_most_rel_thread",))
        #start_new_thread(analysis_sentiment_and_keyword,(data,max_questions,dimensions,sheets,"azure_sentiment&keyword_thread",))
        #start_new_thread(grouping_topic,(data,max_questions,"azure_grouping_topic_thread",))
    most_rel(data,value,"azure_most_rel_thread")
    analysis_sentiment_and_keyword(data,max_questions,dimensions,sheets,"azure_sentiment&keyword_thread")
    grouping_topic(data,max_questions,"azure_grouping_topic_thread")
    #google.analysis_sentiment_and_keyword(data,max_questions,dimensions,sheets)
    #except :
    #    print('unable to start thread !!!')

if __name__ == '__main__' :
    try:
        config = configparser.ConfigParser()
        config.read('config.ini')
        os.remove(config['azure_params']['output_xlsx_file'])
    except OSError:
        pass

    route_function()

