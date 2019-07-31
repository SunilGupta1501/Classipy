######################################
#
#       Software calls classification final
#
######################################
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
#from sklearn.grid_search import GridSearchCV
#from sklearn.cross_validation import StratifiedKFold, cross_val_score, train_test_split 
from sklearn.tree import DecisionTreeClassifier 
from sklearn.feature_selection import SelectKBest, chi2
#from textblob import TextBlob
import os
import xlwt
data = pd.read_csv("C:\\Users\\SG185314\\Desktop\\Education\\py\\Software\\Software_Calls.csv",encoding='latin-1',index_col = 'WO Nbr / SR Nbr')
                        # merging the columns 
#data["raw_Comments"] = data['First Remark 1'] + ' ' + data['First Remark 2'] + ' ' +data['First Remark 3'] + ' ' + data['Last Remark 1']+ ' ' + data['Last Remark 2']+ ' ' + data['Last Remark 3']
df1 = data[['raw_Comments']]
df1 = df1.dropna()
df1.iloc[1:10,0]
df = df1.dropna()
#print(df['raw_Comments'].apply(lambda x: len(x.split(' '))).sum())


from nltk.tokenize import RegexpTokenizer
tokenizer = RegexpTokenizer(r'\w+')

################## REmoving all the unwanted numbers
stemmer = SnowballStemmer('english')
words = stopwords.words("english")

df["cleaned_Comments"] = df['raw_Comments'].apply(lambda x: " ".join([stemmer.stem(i) for i in re.sub("[^a-zA-Z]"," ", x).split() if i not in words]).lower())
print(df['cleaned_Comments'].apply(lambda x: len(x.split(' '))).sum())

os.chdir(r"C:\Users\SG185314\Desktop\Education\py\Software")
################  Recall the both transformation as well as Model
with open('tfidf_transform', 'rb') as transfering_tfidf:  
    transformation = pickle.load(transfering_tfidf)

with open('text_classifier', 'rb') as training_model:  
    model = pickle.load(training_model)
#################################################################    

y = transformation.transform(df["cleaned_Comments"])

classification = model.predict(y)
classification = pd.DataFrame(classification)
classification = classification.set_index(df.index)
output_data = pd.concat([df[['raw_Comments']], classification], axis = 1)
output_data.shape

import xlwt

workbook = xlwt.Workbook()  
sheet = workbook.add_sheet("Classification") 
  
# Specifying style for the special cell 
style = xlwt.easyxf('font: bold 1') 

# Specifying style for the special cell in the date and time formate
for i in range(0,output_data.shape[0]):
    sheet.write(i+1, 0, str(output_data.index[i]))
    sheet.write(i+1, 1, output_data.iloc[i,0])
    sheet.write(i+1, 2, output_data.iloc[i,1])
## Naming the column
sheet.write(0, 0, 'WO Nbr / SR Nbr', style)    
sheet.write(0, 1, "Remark", style)    
sheet.write(0, 2, "Flage", style)    
 # adjust the column width for each cell       
first_col = sheet.col(1);first_col.width = 256*30
sec_col = sheet.col(2);sec_col.width = 256*20
zero_col = sheet.col(0);zero_col.width = 256*20
##  closing the sheet
workbook.save("Software-classification.xls")

############################################################################################
os.getcwd()
'''

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

import os
os.getcwd()

'''




import xlwt

def SVM_CLA():
    os.chdir(r"C:\\Users\\SG185314\\Desktop\\Education\\py\\Software")
    data = pd.read_csv("Software_Calls.csv",encoding='latin-1',index_col = 'WO Nbr / SR Nbr')
                        # merging the columns 
    	#data["raw_Comments"] = data['First Remark 1'] + ' ' + data['First Remark 2'] + ' ' +data['First Remark 3'] + ' ' + data['Last Remark 1']+ ' ' + data['Last Remark 2']+ ' ' + data['Last Remark 3']
    df1 = data[['raw_Comments']]
    df1 = df1.dropna()
    df1.iloc[1:10,0]
    df = df1.dropna()
    #print(df['raw_Comments'].apply(lambda x: len(x.split(' '))).sum())
    
    
    from nltk.tokenize import RegexpTokenizer
    tokenizer = RegexpTokenizer(r'\w+')
    
    ################## REmoving all the unwanted numbers
    stemmer = SnowballStemmer('english')
    words = stopwords.words("english")
    
    df["cleaned_Comments"] = df['raw_Comments'].apply(lambda x: " ".join([stemmer.stem(i) for i in re.sub("[^a-zA-Z]"," ", x).split() if i not in words]).lower())
    #	print(df['cleaned_Comments'].apply(lambda x: len(x.split(' '))).sum())
    
    os.chdir(r"C:\Users\SG185314\Desktop\Education\py\Software")
    ################  Recall the both transformation as well as Model
    with open('tfidf_transform', 'rb') as transfering_tfidf:  
    	transformation = pickle.load(transfering_tfidf)
    
    with open('text_classifier', 'rb') as training_model:  
        model = pickle.load(training_model)
    #################################################################    
    
    y = transformation.transform(df["cleaned_Comments"])
    
    classification = model.predict(y)
    classification = pd.DataFrame(classification)
    classification = classification.set_index(df.index)
    output_data = pd.concat([df[['raw_Comments']], classification], axis = 1)
    
    workbook = xlwt.Workbook()  	
    sheet = workbook.add_sheet("Classification") 
      
    # Specifying style for the special cell 
    style = xlwt.easyxf('font: bold 1') 
    
    # Specifying style for the special cell in the date and time formate
    for i in range(0,output_data.shape[0]):
    	sheet.write(i+1, 0, str(output_data.index[i]))
    	sheet.write(i+1, 1, output_data.iloc[i,0])
    	sheet.write(i+1, 2, output_data.iloc[i,1])
    	## Naming the column
    sheet.write(0, 0, 'WO Nbr / SR Nbr', style)    
    sheet.write(0, 1, "Remark", style)    
    sheet.write(0, 2, "Flage", style)    
     # adjust the column width for each cell       
    first_col = sheet.col(1);first_col.width = 256*30
    sec_col = sheet.col(2);sec_col.width = 256*20
    zero_col = sheet.col(0);zero_col.width = 256*20
    ##  closing the sheet
    workbook.save("Software-classification.xls")

SVM_CLA()