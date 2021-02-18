# -*- coding: utf-8 -*-
"""
Created on Sun Jan 17 17:40:41 2021

@author: Kay
"""

import pandas as pd

def redraft(file):
    df = pd.read_excel(file) #Read the excel file in as a dataframe
    df = df.sort_values(by=['Draft Class','WS'], ascending=False, ignore_index=True) #Sort by Draft Class & Prime Wins
    counts = df['Draft Class'].value_counts()
    counts = counts.sort_index(ascending = False)
    print(counts)
    
    redraft = []
    for draftClass in range(len(counts)):
        for pick in range(counts.iloc[draftClass]):
            redraft.append(pick+1)
    
    df['ReDraft Position'] = redraft
    df.to_excel('Redraft.xlsx')
