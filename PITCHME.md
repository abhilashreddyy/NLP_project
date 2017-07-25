
### Text Analysing Application using
### Natural Language Processing(NLP)

---
## What is this application about ?
+++
* Organizations typically conduct a lot of surverys . To minimize the amount of works Surveys are of the form that get the users to chosse a number as an answer so that it can be easily processed. But a lot of qualitative / behavioural surveys also need to be conducted in an organization and expect the users to type in answers .
  * Questions and Answers are collated as excel files since they are easier to process
  * Each sheet typically consists of questions (as columns ) and answers by multiple people (as rows )
  * Multiple surveys are stored as  sheets
+++
* Time consuming to read all the text and come to a conclusion
* This application attempts to provide a better intution /summary of the answers at a glance.
* Uses azure and google NLP cloud services

+++
Screen shot of input data
![Sentiment Pie Chart](/images/data.png)
---
## What this application does ?
+++
### 1. Detects Sentiment of an answer . Appends a column to the Excel sheet with sentiment

### 2. Detects the keywords  of an answers and adds a column to the Excel sheet with the keyword
+++
### 3. Generates a Word Cloud of Keywords for each question ( based on the keyword and Topic )

<img src="images/wordcloud.png" width="500" height = "300" float = 'right'>

this is an image generated from answers about question
What is professional excellence ?


+++
### 4. Generates Pie charts illustrating sentiment of data
![Sentiment Pie Chart](/images/piechart.png)

+++
### 5. Generates a csv file ralted to most relevant answers for a set of keywords
---
### Snap shot of the project
![Projct Snapshot](/images/project_screen_shot.png)


+++
* Duration  
  * 8 weeks

* Number of people
  * 1
---
### Functional Objectives
* Sentiment analysis
* Key-words extraction
* Topic detection
* Extract most relevant answers for given keywords ( by using the data of the question )
+++
### Other Objectives
* Provide a test bed to evaluate the Google Cloud NLP and Azure NLP Cloud Services
* Evaluate the feasibility of not using local version (e.g) Spacy / Stanford NLP
* Evaluate the possibility of offering this as a service for Office 365 Forms
---
## Features of Google NLP
+++
* Sentiment Analysis
<table>
  <tbody>
    <tr>
      <th>Input</th>
      <th>Output</th>
    </tr>
    <tr>
      <td><pre>{
  "encodingType": "UTF8",
  "document": {
    "type": "PLAIN_TEXT",
    "content": "Enjoy your vacation!"
  }
}</pre></td>
      <td><pre>{
  "documentSentiment": {
    "magnitude": 0.8,
    "score": 0.8
  },
  "language": "en",
  "sentences": [
    {
      "text": {
        "content": "Enjoy your vacation!",
        "beginOffset": 0
      },
      "sentiment": {
        "magnitude": 0.8,
        "score": 0.8
      }
    }
  ]
}</pre></td>
    </tr>
  </tbody>
</table>
+++
* Entity analysis
```json
  {
  "entities": [
    {
      "name": "Lawrence of Arabia",
      "type": "WORK_OF_ART",
      "metadata": {
        "mid": "/m/0bx0l",
        "wikipedia_url": "http://en.wikipedia.org/wiki/Lawrence_of_Arabia_(film)"
      },
      "salience": 0.75222147,
      "mentions": [
        {
          "text": {
            "content": "Lawrence of Arabia",
            "beginOffset": 1
          },
          "type": "PROPER"
        },
        {
          "text": {
            "content": "film biography",
            "beginOffset": 39
          },
          "type": "COMMON"
        }
      ]
    },
```
+++
* Syntax Analysis
```json
   {
      "text": {
        "content": "is",
        "beginOffset": 35
      },
      "partOfSpeech": {
        "tag": "VERB",
        "mood": "INDICATIVE",
        "number": "SINGULAR",
        "person": "THIRD",
        "tense": "PRESENT",
      },
      "dependencyEdge": {
        "headTokenIndex": 7,
        "label": "ROOT"
      },
      "lemma": "be"
    },
```
---
## Features of Azure NLP
+++
* Input for azure services
```
{
     "documents": [
         {
             "language": "en",
             "id": "1",
             "text": "First document"
         },
         ...
         {
             "language": "en",
             "id": "100",
             "text": "Final document"
         }
     ]
 }
```
+++
* Sentiment Analysis
```
{
       "documents": [
         {
             "id": "1",
             "score": "0.934"
         },
         ...
         {
             "id": "100",
             "score": "0.002"
         },
     ]
 }
```
+++
* Key phrases Extraction
```
{
       "documents": [
         {
             "id": "1",
             "keyPhrases": ["key phrase 1", ..., "key phrase n"]
         },
         ...
         {
             "id": "100",
             "keyPhrases": ["key phrase 1", ..., "key phrase n"]
         },
     ]
 }
```
+++
### Topic Detection
* Send request to azure services
* Responds with an URL representing the location where computation of topic detection is performed
* Send get request to URL
* Responds with detected Topics 
+++
<table>
  <tbody>
    <tr>
      <th>Input</th>
      <th>Output</th>
    </tr>
    <tr>
      <td><pre>{
     "documents": [
         {
             "id": "1",
             "text": "First document"
         },
         ...
         {
             "id": "100",
             "text": "Final document"
         }
     ],
     "stopWords": [
         "issue", "error", "user"
     ],
     "stopPhrases": [
         "Microsoft", "Azure"
     ]
 }
</pre></td>
      <td><pre>{
        "topics" : [{
            "id" : "string",
            "score" : "number",
            "keyPhrase" : "string"
        }],
        "topicAssignments" : [{
            "documentId" : "string",
            "topicId" : "string",
            "distance" : "number"
        }],
        "errors" : [{
            "id" : "string",
            "message" : "string"
        }]
    }</pre></td>
    </tr>
  </tbody>
</table>
---

* Request cloud services for sentiment of each response for a survey and append sentiment to respective response in excel filename
* Topics detected from given text corpus are used to generate word cloud
* Keywords extracted are used for generating word cloud and extracting most relevant answers
---
### Objectives achieved with azure services
* Uses azure *sentiment analysis*(piechart)
* Uses azure key word detection for *extracting keywords*(word cloud)
* Uses azure *topic detection* for generating topic of each sentence (word cloud)
  * keys generated avoids noise in image
* also uses WUP similarity for generating Word-Cloud related to keywords in question​

+++
### But !!!
* topic detection *takes 7 to 8 minutes* for returning result
* Some part of request does not support asynchronous requests
* Extracting keywords related to question take 6 to 7 seconds of compilation

+++
Word Cloud generated from topics
![Sentiment Pie Chart](images/topicwise.png)
+++

Word Cloud generated from keys
![Sentiment Pie Chart](/images/keywordwise1.png)
+++

Word Cloud generated from keywords related to question
![Sentiment Pie Chart](/images/wordcloud.png)

---
### Objectives achieved with google services
* Uses google *sentiment analysis*
* Uses google *syntax analysis*
  * Returns every word saperately tagged with
    * parts of speech (tag,mood,tense etc.)
    * dependency edge of syntax tree
    * lemma of each word
    * Selected parts of speech(noun,verb) are used for extracting keywords  
* Google NLP service also have *entity analysis* which extracts *common nouns* from text
+++

### Word Cloud generated from google cloud data

![Sentiment Pie Chart](/images/googlewordcloud.png)

---
### comparison of google and azure services

google | azure
-------|-------
Returns a dictionary object | returns a json formated string
takes 2.5-2.6 seconds for performing sentiment and keyword analysis | takes 2.2-2.3 seconds for performing sentiment and keyword analysis
Do not have this features | takes 7-8 + no-of-questions minutes for topic detection
+++
google | azure
-------|-------
best for combined analysis of entire file content | best for seperate analysis (eg :if format is like one question many answers)
entity recognition | do not contain this feature




---

### Asynchronous nature of code
* (best ans most intresting part of application)
* First I sent http request to REST services **one by one**
  * as a result it **took 30 seconds** for each time compile and run
* Used libraries **asyncio,aiohttp** and some features of python3 *yield from*
* When i made it **asynchronous** it took **less than 3 seconds** for compile and run
+++

<img src="/images/sync.jpg" alt="Sync Request" style="width: 800px; height: 350;"/>

+++

<img src="/images/async_explanation.jpg" alt="Sync Request" style="width: 450px; height: 600;"/>

+++

<img src="/images/async.jpg" alt="Sync Request" style="width: 800px; height: 350;"/>

---
# Design of Application
+++
###Pseudo Code
* Reading from excel file .
  * Initally checking dimensions of file .
  * read entire file column wise .
* Send async request to cloud services .
* Generate Word cloud .
  * Uisng keywords
* Generate piechart .
  * count positive negative and neutral sentences
  * project them on piechart
* Write sentimet and keyword info to excel file .
*
---
###Reading from file
* decides no of columns by considering no of headings
* decides no of rows by considering no of numbers in first column
* returns 2d list of strings and dimensions of each sheet and name of sheets in workbook(excel file)
* next slide consists of code chunk which performs this operation

+++?code=azure.py&lang=python
@[37-47](loads parameters and file required to read)
@[48](finds dimensions of sheet)
@[21-33](finds dimensions of sheet)
@[49-57](returns data in sheet)

---

###sending async request to azure cloud service
* Trigger to sends all requests
* stores their memory locations
* waits until all responds


+++?code=azure.py&lang=python
@[179-191](loads required parameters)
@[192-195](transfer control for sending async request and location)
@[172-176](async function)
@[201-202](waits until response)
@[203-221](filter required data and handle exceptions)

---
## generating word cloud and pie chart
+++?code=azure.py&lang=python
@[142-149](word cloud generation)   
@[151-169](piechart generation)

---
### Remaining processing
* Same thing done with topic detection
  * but a little complicated(code)
* Extracted most similar keys in answers from a set of initally provided keys
* Same thing done with google API with some changes in format of input and output

---
## Future scope
* If integrated with Google or Office 365 online forums
  * Can directly analyse online discussion forums
* Can add a function for gathering most relevant reviews or answers for given keywords
* Pythom has largest support for natural language processing libraries (Spacy,NLTK etc)
