import json
import os
import shutil
import sys

import requests
from bs4 import BeautifulSoup
from colorama import Back, Fore, Style, init
from progress.bar import Bar

# Initialize colorama for colored terminal (for Windows)
init()
print(Fore.BLACK + Back.CYAN + 'E' + Style.RESET_ALL + Fore.WHITE + 'xcel ', end='')
print(Fore.BLACK + Back.CYAN + 'F' + Style.RESET_ALL + Fore.WHITE + 'unction ', end='')
print(Fore.BLACK + Back.CYAN + 'T' + Style.RESET_ALL + Fore.WHITE + 'ranslator' + Fore.CYAN)

# Create progress bar
bar = Bar('Initializing', fill='█', suffix='%(percent)d%%')

# Clean html folder if it exists
if os.path.exists('html'):
    shutil.rmtree('html')

# Clean json folder if it exists
if os.path.exists('json'):
    shutil.rmtree('json')

# Create html folder, json folder, and database file
os.mkdir('html')
os.mkdir('json')

# Progress bar for cleaning and creating external files
bar.next()

# Collect data from the internet
responseObject = requests.get('http://dolf.trieschnigg.nl/excel/index.php')

# Set encoding format
responseObject.encoding = 'UTF-8'

# Save index.html in html folder
file = open(os.path.join('html', 'index.html'), 'w')
file.writelines(responseObject.text)
file.close()

# Progress bar for collecting index.html
bar.next()

# Parse index.html
soup = BeautifulSoup(open(os.path.join('html', 'index.html')), 'html.parser')

# Progress bar for parsing index.html
bar.next()

# Count the number of languages
dictionary = {}
languages = []
jsonfiles = []
for content in soup.find_all('th'):
    dictionary[content.get('langid')] = {}
    languages.append(content.get('langid'))

# Progress bar for counting the number of languages
bar.next()

# Save all html files in html folder
for content in soup.find_all('th'):
    dictionary[content.get('langid')]['name'] = content.get_text()
    dictionary[content.get('langid')]['langid'] = content.get('langid')
    if content.find('a') != None:
        dictionary[content.get('langid')]['link'] = 'http://dolf.trieschnigg.nl/excel/' + content.find('a').get('href')
        dictionary[content.get('langid')]['json'] = content.get('langid') + '.json'
        responseObject = requests.get(dictionary[content.get('langid')]['link'])
        # Set encoding format
        responseObject.encoding = 'UTF-8'
        jsonfiles.append(content.get('langid'))
        # Import Excel functions from html files in html folder
        file = open(os.path.join('html', content.get('langid') + '.html'), 'w')
        file.writelines(responseObject.text)
        file.close()
        # Create language specific json files
        file = open(os.path.join('json', content.get('langid') + '.json'), 'w')
        file.close()
    else:
        dictionary[content.get('langid')]['link'] = 'None'
        dictionary[content.get('langid')]['json'] = 'None'
    # Progress bar for all html files and json files
    bar.next(3)
file = open(os.path.join('json', 'structure.json'), 'w')
file.write(json.dumps(dictionary))
file.close()

# Save all Excel functions in index.json
dictionary.clear()
for content in soup.find_all('tr'):
    dictionary[content.find('td').get_text()] = {}
for row in soup.find_all('tr'):
    count = 0
    for content in row.find_all('td'):
        dictionary[row.find('td').get_text()
                   ][languages[count]] = content.get_text()
        count = count + 1
file = open(os.path.join('json', 'index.json'), 'w')
file.write(json.dumps(dictionary))
file.close()

# Progress bar for saving all Excel functions
bar.next()

# Save Excel functions and descriptions in language specific json files
for jsonfile in jsonfiles:
    soup = BeautifulSoup(
        open(os.path.join('html', jsonfile + '.html')), 'html.parser')
    dictionary.clear()
    for content in soup.find_all('tr'):
        dictionary[content.find_all('td')[2].get_text()] = {}
    for content in soup.find_all('tr'):
        dictionary[content.find_all('td')[2].get_text()][jsonfile] = content.find_all('td')[0].get_text()
        dictionary[content.find_all('td')[2].get_text()][jsonfile +
                                                         '_description'] = content.find_all('td')[1].get_text()
        dictionary[content.find_all('td')[2].get_text()]['en_description'] = content.find_all('td')[3].get_text()
    file = open(os.path.join('json', jsonfile + '.json'), 'w')
    file.write(json.dumps(dictionary))
    file.close()
    # Progress bar for each language specific json file
    bar.next(3)

# Clean html folder if it exists
if os.path.exists('html'):
    shutil.rmtree('html')

# Destroy progress bar
bar.next(bar.remaining)
bar.clearln()
bar.finish()

# Show language list
file = open(os.path.join('json', 'structure.json'), 'r')
dictionary = json.loads(file.read())
file.close()
print(Fore.WHITE + Back.BLUE + str('Language').ljust(20) + str('Code').ljust(10), end='')
print(str('Language').ljust(20) + str('Code').ljust(10), end='')
print(str('Language').ljust(20) + str('Code').ljust(10) + Style.RESET_ALL)
print(Fore.BLUE + '-' * 90)
count = 0
for content in dictionary:
    if dictionary[content]['json'] == 'None' and dictionary[content]['langid'] != 'en':
        print(Fore.WHITE + str((dictionary[content]['name']).split('/')[0]).ljust(20), end='')
        print(Fore.WHITE + str(dictionary[content]['langid']).ljust(10), end='')
    else:
        print(Fore.GREEN + str((dictionary[content]['name']).split('/')[0]).ljust(20), end='')
        print(Fore.MAGENTA + str(dictionary[content]['langid']).ljust(10), end='')
    if count == 2:
        print()
        print(Fore.BLUE + '-' * 90)
        count = 0
    else:
        count = count + 1
print()
print(Fore.BLUE + '-' * 90)
print(Fore.WHITE + Back.BLUE + str('Note: There is no description for languages in white.').ljust(90))
print(Style.RESET_ALL)

# Ask for languages for search and result, and check the validity
print(Fore.CYAN + '>> Code of language for ' + Fore.WHITE + Back.CYAN +
      'search' + Style.RESET_ALL + Fore.CYAN + ': ' + Fore.YELLOW, end='')
searchLanguage = str(input()).lower()[0:2]
while searchLanguage not in dictionary:
    print(Fore.RED + 'It seems that there is no match for the code...please try again:)')
    print(Fore.CYAN + '>> Code of language for ' + Fore.WHITE + Back.CYAN +
          'search' + Style.RESET_ALL + Fore.CYAN + ': ' + Fore.YELLOW, end='')
    searchLanguage = str(input()).lower()[0:2]
print(Fore.CYAN + '>> Code of language for ' + Fore.WHITE + Back.CYAN +
      'result' + Style.RESET_ALL + Fore.CYAN + ': ' + Fore.YELLOW, end='')
resultLanguage = str(input()).lower()[0:2]
while searchLanguage not in dictionary:
    print(Fore.RED + 'It seems that there is no match for the code...please try again:)')
    print(Fore.CYAN + '>> Code of language for ' + Fore.WHITE + Back.CYAN +
          'result' + Style.RESET_ALL + Fore.CYAN + ': ' + Fore.YELLOW, end='')
    resultLanguage = str(input()).lower()[0:2]
searchJsonFile = dictionary[searchLanguage]['json']
resultJsonFile = dictionary[resultLanguage]['json']
print()

# Show instructions in terminal
print(Fore.CYAN + 'Instructions')
print(Fore.CYAN + '==> Press letter(s) with Enter to see close functions')
print(Fore.CYAN + '==> Press QUIT with Enter to quit at any time')
print(Style.RESET_ALL)

# Import language specific data from json files
searchDictionary = {}
resultDictionary = {}
file = open(os.path.join('json', 'index.json'), 'r')
searchDictionary = json.loads(file.read())
file.close()
if resultLanguage == 'en':
    file = open(os.path.join('json', 'fr.json'), 'r')
    resultDictionary = json.loads(file.read())
    file.close()
elif resultJsonFile == 'None':
    resultDictionary = None
else:
    file = open(os.path.join('json', resultJsonFile), 'r')
    resultDictionary = json.loads(file.read())
    file.close()

# Loop for searching and showing results
functions = []
indexFunction = ''
while True:
    functions.clear()
    print(Fore.YELLOW + '>> ', end='')
    userInput = str(input()).upper()
    if userInput == 'QUIT':
        sys.exit()
    for function in searchDictionary:
        if userInput == searchDictionary[function][searchLanguage]:
            functions.clear()
            functions.append(searchDictionary[function][searchLanguage])
            indexFunction = searchDictionary[function]['en']
            break
        elif str(searchDictionary[function][searchLanguage]).find(userInput) == 0:
            functions.append(searchDictionary[function][searchLanguage])
            indexFunction = searchDictionary[function]['en']
    if len(functions) < 1:
        print(Fore.RED + 'It seems that there is no match for the letter(s)...please try again:)')
    elif len(functions) != 1:
        print(Fore.CYAN + str(functions))
    else:
        print(Fore.MAGENTA + searchLanguage.upper() + ': ' + Fore.GREEN + str(functions[0]))
        print(Fore.MAGENTA + resultLanguage.upper() + ': ' + Fore.GREEN +
              str(searchDictionary[indexFunction][resultLanguage]))
        if resultDictionary != None:
            if resultLanguage == 'en' and indexFunction in resultDictionary:
                print(Fore.MAGENTA + 'DESCRIPTION: ' + Fore.GREEN +
                      str(resultDictionary[indexFunction]['en_description']))
            elif indexFunction in resultDictionary:
                print(Fore.MAGENTA + 'DESCRIPTION: ' + Fore.GREEN +
                      str(resultDictionary[indexFunction][resultLanguage + '_description']))
