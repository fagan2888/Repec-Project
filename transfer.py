"""
Author: Peifan Wu
Last update: 21st August, 2014
"""

import codecs
import numpy as np
import pandas as pd
import string
import re

def url_valid(url):
    regex = re.compile(
        r'^https?://'  # http:// or https://
        r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+[A-Z]{2,6}\.?|'  # domain...
        r'localhost|'  # localhost...
        r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' # ...or ip
        r'(?::\d+)?'  # optional port
        r'(?:/?|[/?]\S+)$', re.IGNORECASE)
    return regex.search(url)

Template = u"Template-Type: ReDIF-Paper 1.0"

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
		# if isinstance(df.loc[idx]['Year'], int):
		years.append(df.loc[idx]['Year'])
			
years.sort()
years = np.array(years, dtype = 'int')

# Generate file headers
file_list = []
index_list = []
for year in years:
	current_file_name = str(year) + ".rdf"
	current_file = codecs.open(current_file_name, encoding = 'utf-8', mode = "w")
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
		# if isinstance(df.loc[idx]['Year'], int):
		year_index = years.searchsorted(df.loc[idx]['Year'])
		# else:
		#	flag = True
		#	log_file.write("Original Entry " + str(idx) + " year data has wrong format.\n")
	
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
		if url_valid(df.loc[idx]['File URL']):
			file_url = df.loc[idx]['File URL']
		else:
			flag = True
			log_file.write("Original Entry " + str(idx) + " URL is invalid.\n")
	
	if (not flag): # Then all necessary information are available
		file_list[year_index].write(Template + u"\n")
		file_list[year_index].write(u"Title: " + title + u"\n")
		for i in range(len(authors)):
			file_list[year_index].write(u"Author-Name: " + authors[i] + u'\n')
			file_list[year_index].write(u"X-Author-Name-First: " + authors_name_first[i] + u'\n')
			file_list[year_index].write(u"X-Author-Name-Last: " + authors_name_last[i] + u'\n')
		file_list[year_index].write(u"File-URL:\n")
		file_list[year_index].write(file_url + u"\n")
		file_list[year_index].write(u"File-Format: application/pdf\n")
		file_list[year_index].write(u"Creation-Date: " + str(years[year_index]) + u"\n")
		file_list[year_index].write(u"Handle: RePEc:ste:nystbu:" + str(years[year_index] % 1000) + "-" + str(current_index).zfill(2) + u"\n")
		file_list[year_index].write(u"\n")
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