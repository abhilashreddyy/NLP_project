
import azure
import json
import configparser
import os

def analysis_sentiment_and_keyword():

    config = configparser.ConfigParser()
    config.read('config.ini')
    empty_cells = []
    value = [['environment','technical','security','passion','intelligence','customer','satisfaction'],['friendly','together','understanding','attitude'],['resource','time'],['clear','understanding','undertaking','review']]
    sol,keys,no_of_sol,sol1,jsontext,wordcloud_text = [],[],[],[],[],[]
    data,max_questions,dimensions,sheets = azure.read_questions()
    most_rel = azure.extract_most_relevant(data,value)
    azure.write_to_file(most_rel,data)
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

    sentiment_info,keywords_info = azure.send_nlp_request_azure([sentiment_serviceurl,keywords_serviceurl],jsontext,key)
    temp = empty_cells[:]
    for i in range(len(sentiment_info)) :
        for j in range(len(sentiment_info[i])) :
            if temp == [] :
                break
            if temp[0][0] == i and temp[0][1] == j :
                temp.pop(0)
                sentiment_info[i][j]['score'] = .5
                keywords_info[i][j]['keyPhrases'] = []



    azure.output_exel_with_question(sentiment_info,max_questions,dimensions,sheets,"sentiment")
    azure.output_exel_with_question(keywords_info,max_questions,dimensions,sheets,"keyword")
    azure.generate_piechart_for_azure_sentiment(sentiment_info)

    for word in keywords_info :
        string = ''
        for word1 in word :
            for word2 in word1['keyPhrases'] :
                string = string + word2 + " "
        wordcloud_text.append(string)

    azure.generate_word_cloud(wordcloud_text,"keywordwise")

    value = ['environment','technical','security','passion','intelligence','customer','satisfaction']


    '''gradient = 3
    treshold = .5
    simplified_keywords_info = []
    question = [['environment','technical','security','passion','intelligence','customer','satisfaction'],['environment','technical','security','passion','intelligence','customer','satisfaction'],['environment','technical','security','passion','intelligence','customer','satisfaction'],['environment','technical','security','passion','intelligence','customer','satisfaction']]
    for i,word1 in enumerate(keywords_info) :
        simplified_keywords_info.append([])
        for word2 in word1:
            for word3 in word2['keyPhrases'] :
                simplified_keywords_info[i].append(word3)


    q_wordcloud,extra_keys = azure.question_related_keys_extraction(simplified_keywords_info,question,gradient,treshold)

    azure.generate_word_cloud(q_wordcloud,"including_question")'''







def grouping_topic() :

    sol,keys,no_of_sol,wordcloud_text = [],[],[],[]
    data,max_questions,dimensions,sheets = azure.read_questions()
    key = config['azure_params']['key']
    topicgrouping_serviceurl = config['azure_params']['topicgrouping_serviceurl']
    count = 0
    response = azure.send_grouping_azure_req([data[0]],topicgrouping_serviceurl,key)
    for i in response :
        string = ''
        for j in i :
            string = string + j['keyPhrase'] + " "
        wordcloud_text.append(string)

    azure.generate_word_cloud(wordcloud_text,"topicwise")



if __name__ == '__main__' :
    try:
        os.remove('analysis.xlsx')
    except OSError:
        pass

    analysis_sentiment_and_keyword()
    #grouping_topic()
