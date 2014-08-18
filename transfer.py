import numpy as np
import pandas as pd
import string
import re

Template = "Template-Type: ReDIF-Paper 1.0"

log_file_name = "report.log"
log_file = open(log_file_name, "w")

excel_file_name = "sample_list.xlsx"
table = pd.read_excel(excel_file_name, 'Sheet1', index_col = None, na_values = ['NA'])

table = table.sort(['Year'])
current_year = 0
current_paper_count = 0

for idx in xrange(len(table)):
	flag = False
	current_entry = table[idx: (idx + 1)]
	cur_idx = current_entry.index[0]
	# Time
	if pd.isnull(current_entry.loc[cur_idx]['Year']):
		flag = True
		log_file.write("Original Entry " + str(current_entry.loc[cur_idx]['Index']) + " year data is blank.\n")
	elif (current_entry.loc[cur_idx]['Year'] != current_year):
		if (current_year != 0):
			current_file.close()
		current_year = current_entry.loc[cur_idx]['Year']
		current_file = open(str(current_year) + '.rdf', "w")
		current_paper_count = 0
	
	# Title
	if pd.isnull(current_entry.loc[cur_idx]['Title']):
		flag = True
		log_file.write("Original Entry " + str(current_entry.loc[cur_idx]['Index']) + " title is blank.\n")
	else:
		title = current_entry.loc[cur_idx]['Title']
	
	# Author
	if pd.isnull(current_entry.loc[cur_idx]['Author']):
		flag = True
		log_file.write("Original Entry " + str(current_entry.loc[cur_idx]['Index']) + " author is blank.\n")
	else:
		authors = current_entry.loc[cur_idx]['Author'].split(";") # authors are split with semicolon
		authors_name_first = []
		authors_name_last = []
		for i in xrange(len(authors)):
			authors[i] = authors[i].lstrip().rstrip()
			name_split = re.split(r'[\s,]+', authors[i]) # split whitespace and comma
			authors_name_last.append(name_split[-1])
			authors_name_first.append(string.join(name_split[:-1]))
	
	# URL
	if pd.isnull(current_entry.loc[cur_idx]['File URL']):
		flag = True
		log_file.write("Original Entry " + str(current_entry.loc[cur_idx]['Index']) + " URL is blank.\n")
	else:
		file_url = current_entry.loc[cur_idx]['File URL']
	
	if (not flag): # Then all necessary information are available
		current_paper_count = current_paper_count + 1
		current_file.write(Template + "\n")
		current_file.write("Title: " + title + "\n")
		for i in xrange(len(authors)):
			current_file.write("Author-Name: " + authors[i] + "\n")
			current_file.write("X-Author-Name-First: " + authors_name_first[i] + "\n")
			current_file.write("X-Author-Name-Last: " + authors_name_last[i] + "\n")
		current_file.write("File-URL:\n")
		current_file.write(file_url + "\n")
		current_file.write("File-Format: application/pdf\n")
		current_file.write("Cureation-Date: " + str(current_year) + "\n")
		current_file.write("Handle: RePEc:ste:nystbu:" + str(current_year % 1000) + "-" + str(current_paper_count).zfill(2) + "\n")
		current_file.write("\n")
		log_file.write("Original Entry " + str(current_entry.loc[cur_idx]['Index']) + " processed successfully.\n")
		
current_file.close()
log_file.close()