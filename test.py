import onepy

notebooks = onepy.getHierarchy()

#Prints the firt notebook
print (notebooks[0])


#Print all the sections of the first notebook
print(notebooks[0]["sections"])

#Print all the sections of the first sectiongroup in the first notebook
print (notebooks[0]["sectionGroups"][0]["sections"])


#Print all the pages in the first section of the first notebook
print (notebooks[0]["sections"][0]["pages"])

#Print the name of every page in every section in the first notebook
for section in notebooks[0]["sections"]:
    for page in section["pages"]:
        print(page["name"])



