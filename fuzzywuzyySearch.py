import tkinter as tk
import re
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import pandas as pd
from tkinter import *
from tkinter import filedialog
app = tk.Tk()
entry = tk.Entry(app)
app['bg'] = '#49A'
app.title('Edunomics search')
search_label=Label(app, text='Enter the description', font=('bold', 14))
search_label.grid(row=0, column=0)
entry.grid(row=0, column=1, pady=20)
result_list=Text(app, height=20, width=80, border=5)
result_list.grid(row=1, column=0, columnspan=3, rowspan=6)
#scrollbar = Scrollbar(app)
#scrollbar.grid(row=1, column=3)
#attach scrollbar to listbox
#result_list.config(yscrollcommand=scrollbar.set)
#scrollbar.config(command=result_list.yview)
df = pd.read_excel('./Mappedsku.xlsx')
def outercover():
    global entry
    query = entry.get()
    #getting dimensions of input
    pattern = re.compile('\d{3,4}X\d{3,4}-\d{1,2}')
    l = []
    for string in pattern.findall(query):
        l.append(string.strip())
    dimension = ' '.join(map(str, l))
    features = query.replace(dimension,'')
    #print(dimension)
    #print(features)
    def search(row):
        prename = row['Company Dscription']
        p = re.compile('\d{3,4}X\d{3,4}-\d{1,2}')
        l = []
        for string in p.findall(prename):
            l.append(string.strip())
        name = ' '.join(map(str, l))
        return fuzz.token_sort_ratio(name.lower(), dimension.lower())
    df1=(df[df.apply(search, axis=1)==100])
    def get_ratio(row):
        prename = row['Company Dscription']
        p=re.compile('\d{3,4}X\d{3,4}-\d{1,2}')
        l = []
        for string in p.findall(prename):
            l.append(string.strip())
        postname = ' '.join(map(str, l))
        name=prename.replace(postname, '')
        return fuzz.token_sort_ratio(name.lower(), features.lower())
    df2=df1[df1.apply(get_ratio,axis=1) > 70]
    df2.index+=2
    result_list.delete('1.0',tk.END)
    result_list.insert(tk.END, str(df2.iloc[:,:]))
search_btn = Button(app, text ='search', width = 12, font = ('bold',14), command = outercover)
search_btn.grid(row=0, column=2, pady=5)
app.geometry('700x400')
app.resizable(0, 0)
app.mainloop()