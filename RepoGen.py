#!/usr/bin/env python
# coding: utf-8

# In[31]:


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os


# In[32]:


data = pd.read_excel("Dummy_Data.xlsx", sheet_name='Sheet1')
#data.drop(axis=0, index=0, inplace=True)
data.columns = data.iloc[0]
data = data[1:]
#data.head(5)


# In[33]:


data['Your score'].replace(to_replace='0', value=0, inplace=True)
#grouped = data.groupby(['Full Name '],axis=1)
std1 = data[data['Registration Number'] == data.iloc[0]['Registration Number']]
std2 = data[data['Registration Number'] == data.iloc[25]['Registration Number']]
std3 = data[data['Registration Number'] == data.iloc[50]['Registration Number']]
std4 = data[data['Registration Number'] == data.iloc[75]['Registration Number']]
std5 = data[data['Registration Number'] == data.iloc[100]['Registration Number']]

sum1 = std1["Your score"].values.sum()
sum2 = std2["Your score"].values.sum()
sum3 = std3["Your score"].values.sum()
sum4 = std4["Your score"].values.sum()
sum5 = std5["Your score"].values.sum()
average = np.mean([sum1, sum2, sum3, sum4, sum5])
average


# In[34]:


cols = std1.columns
#cols


# In[35]:


dirname = os.path.abspath('')
dfs = [std1, std2, std3, std4, std5]
#print(cols[2], cols[3], cols[4], cols[5], cols[6], cols[7], cols[8], cols[9], cols[10], cols[11], cols[12], end="")
#print("")
stdno = 0
for stds in dfs :
    stdno = stdno+1
    fname = stds.iloc[0][cols[2]]
    lname = stds.iloc[0][cols[3]]
    fullname = stds.iloc[0][cols[4]]
    regno = stds.iloc[0][cols[5]]
    grade = stds.iloc[0][cols[6]]
    sname = stds.iloc[0][cols[7]]
    gender = stds.iloc[0][cols[8]]
    dob = stds.iloc[0][cols[9]]
    city = stds.iloc[0][cols[10]]
    dot = stds.iloc[0][cols[11]]
    cnt = stds.iloc[0][cols[12]]
    repo = stds[cols[13:19]]
    result = stds.iloc[0][19]
    repo.fillna('', inplace=True)
    score = stds["Your score"].values.sum()
    #repo.reset_index(drop=True, inplace=True)
    #print(fname+"        "+lname+"       "+ fullname, regno, grade, sname, gender, dob, city, dot, cnt)
    #print(repo)
    #print(tabulate(repo, headers=repo.columns, tablefmt="fancy_grid"))
    print("Student No: ", stdno)
    print("")
    
    print("Personal Details: ")

    
    fig1, (ax1, ax2, ax3, ax4) = plt.subplots(4, figsize=(10,10), gridspec_kw={'height_ratios':[0.2,0.2,1.5,0.2]})
    fig1.suptitle('Science Olympiad Exam', fontsize=16)
    ax1.axis('off')
    ax1.axis('tight')
    ax1.set_title("Personal Details")
    tab1 = ax1.table(cellText=[[fname, lname, fullname, gender, dob, city, cnt]], colLabels=["First Name", "Last Name", "Full Name", "Gender", "Date of birth", "City of residence", "Country of residence"], loc='center', cellLoc='center')
    tab1.auto_set_font_size(False)
    tab1.auto_set_column_width(col=list(range(7))) 
    
    
    ax2.axis('off')
    ax2.axis('tight')
    ax2.set_title("Exam Details")
    print("Exam Details: ")
    tab2 = ax2.table(cellText=[[regno, dot, sname]], colLabels=["Registration No.",  "Date and time of exam", "School Name"], loc='center', cellLoc='center')
    tab2.auto_set_font_size(False)
    tab2.auto_set_column_width(col=list(range(3)))
    
       
    #Printing score card
    #hide axes 
    ax3.axis('off')
    ax3.axis('tight')
    ax3.set_title("Score Card")
    print("Score Card: ")
    tab3 = ax3.table(cellText=repo.values, colLabels=repo.columns, loc='center', cellLoc='center')
    tab3.auto_set_font_size(False)
    tab3.auto_set_column_width(col=list(range(len(repo.columns))))

    im = plt.imread(dirname + '\Pics_for_assignment\{}_{}.jpg'.format(fname, lname))
    newax = fig1.add_axes([0.8,0.7,0.16,0.15], anchor='NE', zorder=1)
    newax.imshow(im)
    newax.axis('off')
    
    ax4.axis('off')
    ax4.axis('tight')
    tab4 = ax4.table(cellText=[[result]], colLabels=["Final Result"], loc='center', cellLoc='center')
    tab4.auto_set_font_size(False)
    tab4.auto_set_column_width(col=list(range(2)))
    
    fig1.tight_layout()
    fig1.patch.set_visible(False)
    
    
    
    fig2, (ax5, ax6) = plt.subplots(1,2,figsize=(10,5))
    #ax.axis('off')
    ax5.get_yaxis().set_visible(False)
    r1 = ax5.bar(1, score, 0.1, align='center')
    r2 = ax5.bar(1.4, average, 0.1, align='center')
#     ax.bar_label(rects1, padding=3)
#     ax.bar_label(rects2, padding=3)
    ax5.set_title("Score comparison")
    ax5.set_xticks([1,1.4])
    ax5.set_xticklabels(["Your Score", "Global Average"])
    ax5.bar_label(r1, padding=3)
    ax5.bar_label(r2, padding=3)
    
    cor = len(repo[repo["Outcome (Correct/Incorrect/Not Attempted)"]=="Correct"].values)
    incor = len(repo[repo["Outcome (Correct/Incorrect/Not Attempted)"]=="Incorrect"].values)
    unat = len(repo[repo["Outcome (Correct/Incorrect/Not Attempted)"]=="Unattempted"].values)
    explode = (0.1, 0, 0)
    labels=["Correct", "Incorrect", "Unattempted"]
    colours = {'Correct': 'C2',
           'Incorrect': 'C1',
           'Unattempted': 'C0'}
    ax6.set_title("Question Analysis")
    ax6.pie([cor, incor, unat], explode=explode, labels=labels, colors=[colours[key] for key in labels],autopct='%1.1f%%',
        shadow=True, startangle=90)
    ax6.axis('equal')

    filePath = dirname + '\Report_Cards\{}_{}.pdf'.format(fname, lname)
    
    if os.path.exists(filePath):
        os.remove(filePath)
    pp = PdfPages(filePath)
    pp.savefig(fig1)
    pp.savefig(fig2)
    pp.close()
    
    plt.show()

    
    
    


# In[ ]:





# In[ ]:





# In[ ]:




