import pandas as pd
import numpy as np
import pandas as pd
import xlwings as xw
import os
from datetime import datetime
from tqdm import tqdm
import xlwt as xw
####Classifier
def SVM_CLA(Path):
    print(Path)
    os.chdir(Path)
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
    #    print(df['cleaned_Comments'].apply(lambda x: len(x.split(' '))).sum())
    
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

    
def PCI(Path):
        os.chdir(Path)
       
        import xlwt as xw
        ### import .xlsb file from excel
##        app = xw.App()
##        book = xw.Book('Parent_chiled.xlsb')
##        sheet = book.sheets('Sheet1')
##        df = sheet.range('A1').options(pd.DataFrame, expand='table').value
##        book.close()
##        app.kill()
       
        df = pd.read_csv('Parent_chiled.csv')
 
        ###  Soet the columns as a required way
        df.iloc[:,0]
        df.sort_values(['ATM ID','SD & T', 'ED &T'], ascending=[True,True, False])
        df.columns = ['ATM_ID',"start_date", "end_date"]
        ATM_ID = df.ATM_ID.unique()
        df["Duration"] = (df["end_date"]-df["start_date"]).dt.seconds + ((df["Duration"].dt.days*24)*3600)
        df["PC"] = ""
        df.shape
        ### Define a temp data frame for the storing the output
        data = df.iloc[[0],:]
       
        ### Finding a parent,child and intersection ticket with respect to updated down time duration
        for ID in tqdm(ATM_ID):
             df1 = df[df['ATM_ID']== ID]
             if(df1.shape[0]>0):
                 for i in range(0,df1.shape[0]-1):
                     for k in range((i+1),df1.shape[0]):
                         if((df1.iloc[i,2]>= df1.iloc[k,1]) and (df1.iloc[i,2]>= df1.iloc[k,2])):
                             if(df1.iloc[i,4] == ""):
                                 df1.iloc[i,4] = "Parent"
                                 df1.iloc[k,4] = "Child"
                                 df1.iloc[k,3] = (df1.iloc[k,1] - df1.iloc[k,1]).dt.seconds + ((df["Duration"].dt.days*24)*3600)
                             else:
                                 df1.iloc[k,4] = "Child"
                                 df1.iloc[k,3] = (df1.iloc[k,1] - df1.iloc[k,1]).dt.seconds + ((df["Duration"].dt.days*24)*3600)
                         if((df1.iloc[i,2]> df1.iloc[k,1]) and (df1.iloc[i,2]< df1.iloc[k,2])):
                             if(df1.iloc[i,4] == ""):
                                 df1.iloc[i,4] = "Parent Intersect"
                                 df1.iloc[k,4] = "Intersect"
                                 df1.iloc[k,3] = (df1.iloc[k,2] - df1.iloc[i,2]).dt.seconds + ((df["Duration"].dt.days*24)*3600)
                             else:
                                 df1.iloc[k,4] = "Intersect"
                                 df1.iloc[k,3] = (df1.iloc[k,2] - df1.iloc[i,2]).dt.seconds + ((df["Duration"].dt.days*24)*3600)
                       
                 data = data.append(df1)
             else:
                 data = data.append(df1)
       
        ### Making a user define function for the time calculation
        def sec_to_hours(seconds):
            a=str(seconds//3600)        
            b=str((seconds%3600)//60)
            if len(b)!=2:
                b = str(0)+b
            c=str((seconds%3600)%60)
            if len(c)!=2:
                c = str(0)+c
            d=["{}:{}:{}".format(a, b, c)]
            return d
       
            
        ### Converting time into a formate of "37:40:00"
        j = 1
        for i in Total_Time:
            df["Duration"].loc[j]= sec_to_hours(i)
            j=j+1
       
        data.iloc[:, [0,1,2,4]]
       
       
        #### Calling a excel workbook and renaming excel sheet
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Pareent Child Identification")
         
        # Specifying style for the special cell
        style = xlwt.easyxf('font: bold 1')
       
        # Specifying style for the special cell in the date and time formate
        date_format = xlwt.XFStyle()
        date_format.num_format_str = 'dd/mm/yyyy hh:mm:ss'
       
        time_format = xlwt.XFStyle()
        time_format.num_format_str = 'hh:mm:ss'
         
        # Specifying column
        for i in tqdm(range(0,data.shape[0])):
            sheet.write(i+1, 0, data.iloc[i,0])
            sheet.write(i+1, 1, data.iloc[i,1],date_format)
            sheet.write(i+1, 2, data.iloc[i,2],date_format)
            sheet.write(i+1, 3, str(data.iloc[i,3]))
            sheet.write(i+1, 4, data.iloc[i,4], style)
        ## Naming the column
        sheet.write(0, 0, "ATM ID", style)   
        sheet.write(0, 1, "Start Date & time", style)   
        sheet.write(0, 2, "End Date & time", style)   
        sheet.write(0, 3, "Duration", style)   
        sheet.write(0, 4, "PC Class", style)   
         # adjust the column width for each cell      
        first_col = sheet.col(1);first_col.width = 256*20
        sec_col = sheet.col(2);sec_col.width = 256*20
        zero_col = sheet.col(0);zero_col.width = 256*12
        ##  closing the sheet
        workbook.save("PCI-Sample.xls")


##Popup complete msg
def popupmsg(msg):
    NORM_FONT= ("Verdana", 10)
    popup = tk.Tk()
    popup.wm_title("!")
    label = ttk.Label(popup, text=msg, font=NORM_FONT)
    label.pack(side="top", fill="x", pady=10)
    B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
    B1.pack()
    popup.mainloop()

    
from tkinter import *

global fields

fields=['Insert Path:']

##Extract entries from userform

def fetch(entries):
   global path
   StoreP(entries)
   x =1
   ReptList=[]
   print(entries)
   for entry in entries:
##       if x==1: 
           Path = entry[1].get()
##           x=x+1
##       elif x==2:
##           Password  = entry[1].get()
##           x=x+1
##       elif (entry[1].get())!= "":
##           ReptList.append(entry[1].get())
##   path= Uname
   SVM_CLA(Path)

def fetch_1(entries):
   global path
   StoreP(entries)
   x =1
   ReptList=[]
   print(entries)
   for entry in entries:
##       if x==1: 
           Path = entry[1].get()
##           x=x+1
##       elif x==2:
##           Password  = entry[1].get()
##           x=x+1
##       elif (entry[1].get())!= "":
##           ReptList.append(entry[1].get())
   PCI(Path)

##Create fields      
def makeform(root, fields):
   entries = []
   for field in fields:
      row = Frame(root)
      if field == 'r':
          lab = Label(row, width=15, anchor='w',justify=LEFT)
      else:
          lab = Label(row, width=15, text=field, anchor='w',justify=LEFT)

      if field=='Password:':
          ent = Entry(row,show='*',width=30)
      else:
          ent = Entry(row,width=30)
      row.pack(side=TOP, fill=X, padx=7, pady=7)
      lab.pack(anchor=W)
      ent.pack(side=RIGHT, expand=YES, fill=X)
      entries.append((field, ent))
   return entries
   

## Add report field
Rpt=0
def insterreport(r):
    Add= makeform(root,r)
    Add=Add[0]
    ents.append(Add)

##Close window 
def close_window():
    root.destroy()

##Store entries for next time
def StoreP(entries):
    PN="LastSave"
    List=[]
    x =1

    for entry in entries:
        if (entry[1].get())!="":
            List.append(entry[1].get())
    
    Write={}
    Write[PN]=List
    fileW= open("Profiles.txt",'w')
    fileW.write(str(Write))

## Extract and add last form    
def ExtractP(entries):
    import ast
    file = open("Profiles.txt")
    PD =file.read()
    if PD!= "":
        LastSave= ast.literal_eval(PD)
        print(LastSave)
        x =0
        ReptList=[]
        #print(entries)
        for i in range(len(LastSave["LastSave"])-2):
            insterreport('r')

        for entry in entries:
            entry[1].insert(0,str(LastSave["LastSave"][x]))
            x=x+1
            
## Main form processing and view setup
if __name__ == '__main__':
   import win32gui, win32con
   from PIL import ImageTk, Image
   import os
   dir_path = "C:\\Users\\SG185314\\Desktop\\Education\\VBA\\Macro\\Distribution List"
   #os.path.join(os.path.join(os.environ['USERPROFILE']))
   #dir_path = os.path.join(__file__)
   Minimize = win32gui.GetForegroundWindow()
   win32gui.ShowWindow(Minimize, win32con.SW_MINIMIZE)
   root = Tk()
   root.title("Classipy")
   top = Frame(root)
   bottom = Frame(root)
   top.pack(side=TOP)
   bottom.pack(side=BOTTOM, fill=BOTH, expand=True)
   
   img = Image.open(os.path.join(dir_path,"NCR_logo.ico"))
   img = img.resize((25, 20), Image.ANTIALIAS)
   img = ImageTk.PhotoImage(img)
   label = Label(root, image = img,text="")
   label.pack(in_=top,side = LEFT)

   label = Label(root, text="  Welcome to Classipy", font = ('Raleway',12),fg = 'blue')
   label.pack(in_=top,side=TOP)

   mainframe = Frame(root)
   ents = makeform(root, fields)
 
   root.bind('<Return>', (lambda event, e=ents: fetch(e)))

   label = Label(root, text="\nSelect from report below:", font = ('Raleway',12),fg = 'blue')
   label.pack(side=TOP)
   ExtractP(ents)
    ## Report Names field
   Add = Button(root, text='Classifier',
          command=(lambda e=ents:fetch(e)), height = 2, width = 10)
   Add.pack(in_=bottom, side=RIGHT)
   
   
   b1 = Button(root, text='PCI',
          command=(lambda e=ents: fetch_1(e)), height = 2, width = 10)
   
   b2 = Button(root, text='Quit', command= close_window, height = 2, width = 10)
   b2.pack(in_=bottom, side=LEFT, padx=5, pady=5)
   b1.pack(in_=bottom, side=RIGHT, padx=5, pady=5)
   root.mainloop()


