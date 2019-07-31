
import xlsxwriter, pickle, re, nltk, csv, sklearn
import pandas as pd
from nltk.stem import SnowballStemmer
from nltk.corpus import stopwords
%matplotlib inline
import matplotlib.pyplot as plt
import numpy as np
from sklearn import model_selection, preprocessing, linear_model, naive_bayes, metrics, svm
from sklearn.feature_extraction.text import CountVectorizer, TfidfTransformer,TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.svm import SVC, LinearSVC
from sklearn.metrics import classification_report, f1_score, accuracy_score, confusion_matrix
from sklearn.pipeline import Pipeline 
import os
import xlsxwriter, pickle, re, nltk, csv, sklearn
import pandas as pd
from nltk.stem import SnowballStemmer
from nltk.corpus import stopwords
%matplotlib inline
import matplotlib.pyplot as plt

data = pd.read_csv(r"C:\Users\SG185314\Desktop\Education\py\software.csv",encoding='latin-1',index_col = 'WO Nbr / SR Nbr')
data.shape              # dimention of data
list(data)              # column names 
                        # merging the columns 
df = data[["Remark","flage1"]]
                            #df = df1.drop(['WO Nbr / SR Nbr'], axis = 1)
df = df.dropna()
df.shape
#df.groupby('Flag1').describe()
print(df['Remark'].apply(lambda x: len(x.split(' '))).sum())

########################################################################

from nltk.tokenize import RegexpTokenizer
tokenizer = RegexpTokenizer(r'\w+')                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             

################## REmoving all the unwanted numbers
stemmer = SnowballStemmer('english')
words = stopwords.words("english")

df["cleaned_Comments"] = df['Remark'].apply(lambda x: " ".join([stemmer.stem(i) for i in re.sub("[^a-zA-Z]"," ", x).split() if i not in words]).lower())
print(df['cleaned_Comments'].apply(lambda x: len(x.split(' '))).sum())

### I have fitted the tfidf model in below code on complete data set

from sklearn.feature_extraction.text import TfidfTransformer  

tfidfconverter = TfidfVectorizer(ngram_range = (1,1), stop_words = "english",sublinear_tf = True, max_features = 10000)


x = tfidfconverter.fit_transform(df["cleaned_Comments"])


train_X, test_X, train_Y, test_Y = model_selection.train_test_split(x,df["flage1"],test_size=0.3)

from sklearn import svm
clf = svm.SVC(gamma='scale')
clf.fit(train_X[0:100000,:], train_Y[0:100000])

train_X[0:5000,:]

prediction = clf.predict(test_X)
accuracy_score(test_Y,prediction)

#####################################################################################
###
###     Saving a model and writing a output in .xlsx  file
###
#####################################################################################

a = df.loc[y_test.index.tolist(),]
a.shape

classification = model.predict(X_test)
classification = pd.DataFrame(classification)
classification = classification.set_index(a.index)
a.index
classification.index
list(a)
data = pd.concat([a[['raw_Comments']], classification], axis = 1)
data.head
data.shape

#data.to_excel("Software_Calls_Classification.xlsx")

os.chdir(r"C:\Users\SG185314\Desktop\Education\py\Software")


trftran = tfidfconverter.fit(df["cleaned_Comments"])
y = trftran.fit_transform(df["cleaned_Comments"])

with open('text_classifier', 'wb') as picklefile:  
    pickle.dump(clf, picklefile)

with open('text_classifier', 'rb') as training_model:  
    model = pickle.load(training_model)

with open('tfidf_transform', 'wb') as picklefile:  
    pickle.dump(trftran, picklefile)

with open('tfidf_transform', 'rb') as training_model:  
    transformation = pickle.load(training_model)


workbook = xlsxwriter.Workbook('Software_Calls_Classification.xlsx')
worksheet = workbook.add_worksheet("Final Data")

bold = workbook.add_format({'bold': True})
worksheet.write(0 , 0, 'WO Nbr / SR Nbr',bold)
worksheet.set_column(0 , 0, 20)
worksheet.write(0 , 1, 'raw_Comments',bold)
worksheet.set_column(0 , 1, 100)
worksheet.write(0 , 2, 'Classification',bold)
worksheet.set_column(0 , 2, 30)
for i in range(1,data.shape[0]):
    worksheet.write(i , 0, data.index[i])
    worksheet.write(i , 1, data.iloc[i,0])
    worksheet.write(i , 2, data.iloc[i,1])
worksheet = workbook.add_worksheet()
workbook.close()