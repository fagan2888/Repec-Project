"""
Author: Peifan Wu
Last update: 19th August, 2014
"""

import numpy as np
import pandas as pd
import string
import re

Template = "Template-Type: ReDIF-Paper 1.0"

log_file_name = "report.log"
log_file = open(log_file_name, "w")

excel_file_name = "sample_list.xlsx"
df = pd.read_excel(excel_file_name, 'Sheet1', index_col = None)

# Drop empty lines
df = df.dropna(how = 'all')

years = []
for idx in df.index:
	flag = False
	for i in xrange(len(years)):
		if (years[i] == df.loc[idx]['Year']):
			flag = True
			break
	if (not flag):
		#if isinstance(df.loc[idx]['Year'], float):
		cur_number = df.loc[idx]['Year'].astype('int')
		years.append(cur_number)
			
years.sort()
years = np.array(years, dtype = 'int')

# Generate file headers
file_list = []
index_list = []
for year in years:
	current_file_name = str(year) + ".rdf"
	current_file = open(current_file_name, "w")
	file_list.append(current_file)
	index_list.append(np.empty(0))

# Loop over each row in the table
All_correct = True
for idx in df.index:
	flag = False
	# Date
	if pd.isnull(df.loc[idx]['Year']):
		flag = True
		log_file.write("Original Entry " + str(idx) + " year data is blank.\n")
	elif (not flag):
		year_index = years.searchsorted(df.loc[idx]['Year'].astype('int'))
	
	# Index for current year
	if pd.isnull(df.loc[idx]['Index']):
		flag = True
		log_file.write("Original Entry " + str(idx) + " index data is blank.\n")
	elif (not flag):
		current_index = df.loc[idx]['Index'].astype('int')
		position = index_list[year_index].searchsorted(current_index)
		if (position < len(index_list[year_index])):
			if (index_list[year_index][position] == current_index):
				flag = True
				log_file.write("Original Entry " + str(idx) + " has duplicated index.\n")
		if (not flag):
			index_list[year_index] = np.insert(index_list[year_index], position, current_index)
		
	# Title
	if pd.isnull(df.loc[idx]['Title']):
		flag = True
		log_file.write("Original Entry " + str(idx) + " title is blank.\n")
	elif (not flag):
		title = df.loc[idx]['Title']
	
	# Author
	if pd.isnull(df.loc[idx]['Author']):
		flag = True
		log_file.write("Original Entry " + str(idx) + " author is blank.\n")
	elif (not flag):
		authors = df.loc[idx]['Author'].split(";") # authors are split with semicolon
		authors_name_first = []
		authors_name_last = []
		for i in range(len(authors)):
			authors[i] = authors[i].lstrip().rstrip()
			name_split = re.split(r'[\s,]+', authors[i]) # split whitespace and comma
			authors_name_last.append(name_split[-1])
			authors_name_first.append(" ".join(name_split[:-1]))
	
	# URL
	if pd.isnull(df.loc[idx]['File URL']):
		flag = True
		log_file.write("Original Entry " + str(idx) + " URL is blank.\n")
	elif (not flag):
		# TODO: add some regular expression here to check whether URL is legal 
		file_url = df.loc[idx]['File URL']
	
	if (not flag): # Then all necessary information are available
		file_list[year_index].write(Template + "\n")
		file_list[year_index].write("Title: " + title + "\n")
		for i in range(len(authors)):
			file_list[year_index].write("Author-Name: " + authors[i] + "\n")
			file_list[year_index].write("X-Author-Name-First: " + authors_name_first[i] + "\n")
			file_list[year_index].write("X-Author-Name-Last: " + authors_name_last[i] + "\n")
		file_list[year_index].write("File-URL:\n")
		file_list[year_index].write(file_url + "\n")
		file_list[year_index].write("File-Format: application/pdf\n")
		file_list[year_index].write("Creation-Date: " + str(years[year_index]) + "\n")
		file_list[year_index].write("Handle: RePEc:ste:nystbu:" + str(years[year_index] % 1000) + "-" + str(current_index).zfill(2) + "\n")
		file_list[year_index].write("\n")
		# log_file.write("Original Entry " + str(idx) + " processed successfully.\n")
	else:
		All_correct = False

if (All_correct):
	log_file.write("All entries are processed successfully.\n")
else:
	log_file.write("Problems are listed above.\n")
		
# Close all the generated files
for file in file_list:
	file.close()
# Close log file
log_file.close()