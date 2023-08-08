# -*- coding: utf-8 -*-
"""
Created on Aug 07  2023

@author: Edward Sazonov
"""


index_text = '<!DOCTYPE html>\n <html>\n <body>\n <h1>Real - time database view </h1>\n'
for pat_idx in range(len(participants_to_find)):
    index_text = index_text + '<p><a href="' + participants_to_find[pat_idx] + '.htm">' + participants_to_find[
        pat_idx] + '</a></p>\n'
index_text = index_text + '</body>\n</html>\n'

try:
    filename = 'wwwroot/index.html'
    with open(filename, 'w') as file:
        file.write(index_text)
except:
    print("Cannot write to file")
