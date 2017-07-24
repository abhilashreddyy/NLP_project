
### Text Analysing Application using
### Natural Language Processing(NLP)

---
## What is this application about ?
+++
* Surveys consists of large amount of text
  * I this case we have an excel file
  * File consists of sheets
  * each file consists of questions and multiple answers by multiple people
* Time consuming to read all the text and come to a conclusion
* This application gives better intution over data in less time
* Uses azure and google NLP cloud services

+++
Screen shot of input data
![Sentiment Pie Chart](/images/data.png)
---
## What this application can do ?
+++
### 1. generates a excel file which is similar to provided excel file but sentiment and keywords are appended to data
+++
### 2. Generates Word Cloud from text

<img src="images/wordcloud.png" width="500" height = "300" float = 'right'>

this is an image generated from answers about question
What is professional excellence ?


+++
### 3. Generates Pie charts illustrating sentiment of data
![Sentiment Pie Chart](/images/piechart.png)

+++
### 4. Generates a csv file ralted to most relevant answers for a set of keywords
---
### Objectives
* Sentiment analysis
* Key-words extraction
* Topic detection
* extract most relevant answers for given keywords

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
WC generated from topics
![Sentiment Pie Chart](images/topicwise.png)
+++

WC generated from keys
![Sentiment Pie Chart](/images/keywordwise1.png)
+++

WC generated from keywords related to question
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

### WC generated from google cloud data

![Sentiment Pie Chart](/images/googlewordcloud.png)

---
### comparison of google andd azure services

google | azure
-------|-------
Returns a dictionary object | returns a json formatted string
takes 2.5-2.6 seconds for performing sentiment and keyword analysis | takes 2.2-2.3 seconds for performing sentiment and keyword analysis
Do not have this features | takes 7-8 + no-of-questions minutes for topic detection
+++
google | azure
-------|-------
best for combined analysis of entire file content | best for seperate analysis (eg :if format is like one question many answers)
entity recognition | do not contain this feature




---
### Why this application written in python ?
* as the code is written in python
  * It have good scope for furthur development
  * Pythom has largest support for natural language processing libraries (Spacy,NLTK etc)
  * Very easy to use on integrating this application with google docs or microsoft online excel sheets

+++

### Asynchronous nature of code
* (best ans most intresting part of application)
* First I sent http request to REST services **one by one**
  * as a result it **took 30 seconds** for each time compile and run
* Used libraries **asyncio,aiohttp** and some features of python3 *yield from*
* When i made it **asynchronous** it took **less than 3 seconds** for compile and run

---
# Design of Application
---

###reading from file
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
@[192-193](transfer control for sending async request and location)
@[172-176](async function)
@[199-200](waits until response)
@[201-219](filter required data and handle exceptions)

---
## generating word cloud and pie chart
+++?code=azure.py&lang=python
@[142-149](word cloud generation)   
@[151-169](piechart generation)

---
### Remaining processing
* Same thing done with topic detection
  * but a little complicated(code)
* Then extrcated related keywords in answers from given keys
* Same thing done with google API with some changes in format of input and output

---
## Future scope
* If integrated with online google or microsoft excel sheets
  * Can directly analyse online discussion forums
* Can add a function for gathering most relevant reviews or answers for given keywords
