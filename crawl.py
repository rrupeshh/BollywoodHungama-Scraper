import requests
from bs4 import BeautifulSoup
import os
import urllib.request
import xlsxwriter

alphabets = ['A','B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
             'V', 'W', 'X', 'Y', 'Z', 'NUM']
alphabets_pages = [
    791, 418, 719, 196, 149, 393, 492,
    203, 471, 770, 315, 1040, 384, 64, 719, 36,
    526, 1248, 610, 80, 130, 118, 5, 143, 130, 87
]

class CastDetails:
    
    def __init__(self):
        self.workbook = xlsxwriter.Workbook('movies_details.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.row = 0
        self.column = -1

    def increaseRow(self):
        self.row += 1


    def getCastNames(self, individualLink, movieName):
        individualLink_source_code = requests.get(individualLink)

        individualLink_plain_text = individualLink_source_code.text

        soup = BeautifulSoup(individualLink_plain_text, 'html.parser')

        newList = []
        for link in soup.findAll('li', {'itemprop':'actors'}):
            actors = link.a.text
            indexOfDots = actors.find('...')
            actorNames = actors[0:indexOfDots]
            actorNames = actorNames.lower()
            actorNames = actorNames.replace(' ', '-')
            lengthOfActorNames = len(actorNames)
            actorNames = actorNames[1:lengthOfActorNames - 1]
            newList.append(actorNames)

        joinString = ','
        actorNamesListString = joinString.join(newList)
        print(actorNamesListString) 
        print('-----------------------')

        self.worksheet.write(self.row, self.column + 9, actorNamesListString)

        detailsList = []

        length = len(soup.select('ul.crew-list > li'))

        for i in range(length):
            x = soup.select('ul.crew-list > li')[i].get_text(strip=True)
            detailsList.append(x)

        # print(detailsList)

        y = ['Banner', 'Release Date','Genre','Music Director','Producer','Language', 'Director','Editor','Cinematographer','Screenplay','Shooting Location','Production Designer','Lyricist','Dialogue','Sound','Writer','Costume Designer','Co-producer','Publicity PRO','Music Company','Playback Singers']
        
        self.worksheet.write(self.row, self.column + 1, movieName)

        rate = soup.findAll('span', {'class':'ML15'})
        if len(rate) == 0:
            rating = 0
        else:
            for r in rate:
                rating = r.text
                rating = rating.replace('/5', '')

        print(rating)

        self.worksheet.write(self.row, self.column + 4, rating)


        for yVal in y:
            for xVal in detailsList:
                if yVal in xVal:
                    xVal = xVal.replace(yVal,'')

                    if yVal == 'Banner':
                        print('Banner  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 5, xVal)

                    if yVal == 'Release Date':
                        print('Release Date  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 2, xVal)

                    if yVal == 'Genre':
                        print('Genre  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 7, xVal)

                    if yVal == 'Music Director':
                        print('Music Director  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 22, xVal)

                    if yVal == 'Producer':
                        print('Producer  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 8, xVal)

                    if yVal == 'Language':
                        print('Language  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 23, xVal)

                    if yVal == 'Director':
                        print('Director  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 24, xVal)

                    if yVal == 'Editor':
                        print('Editor  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 26, xVal)

                    if yVal == 'Cinematographer':
                        print('Cinematographer  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 27, xVal)

                    if yVal == 'Screenplay':
                        print('Screenplay  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 28, xVal)

                    # if yVal == 'Shooting Location':
                    #     print('Shooting Location  : ' + xVal)
                    #     self.worksheet.write(self.row, self.column + 5, xVal)

                    if yVal == 'Production Designer':
                        print('Production Designer  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 33, xVal)

                    if yVal == 'Lyricist':
                        print('Lyricist  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 25, xVal)

                    if yVal == 'Dialogue':
                        print('Dialogue  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 29, xVal)

                    if yVal == 'Sound':
                        print('Sound  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 31, xVal)

                    if yVal == 'Writer':
                        print('Writer  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 10, xVal)

                    if yVal == 'Costume Designer':
                        print('Costume Designer  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 13, xVal)

                    if yVal == 'Co-producer':
                        print('Co-producer  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 21, xVal)

                    if yVal == 'Publicity PRO':
                        print('Publicity PRO  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 14, xVal)

                    if yVal == 'Music Company':
                        print('Music Company  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 32, xVal)

                    if yVal == 'Playback Singers':
                        print('Playback Singers  : ' + xVal)
                        self.worksheet.write(self.row, self.column + 35, xVal)

def getImageLink(individualLink, fileName, castDetailsObj):
    individualLink_source_code = requests.get(individualLink)

    individualLink_plain_text = individualLink_source_code.text

    soup = BeautifulSoup(individualLink_plain_text, 'html.parser')

    for link in soup.findAll('img', {'class':'bh-fixed-cover-image image'}):
        imageLink = link['src']
        print('Image Link:' + imageLink)

        if not os.path.exists('Images'):
            os.makedirs('Images')

        fileName = fileName.replace('  ', ' ')
        fileName = fileName.replace('?','')
        fileName = fileName.replace('/','')
        fileNameLength = len(fileName)
        fileNameInitial = 1
        fileName = fileName[fileNameInitial:fileNameLength]
        fileName = fileName.replace(' ', '-')
        fileName = fileName.lower() + '.jpg'
        print(fileName)

        urllib.request.urlretrieve(imageLink, 'Images/' + fileName)

        castDetailsObj.worksheet.write(castDetailsObj.row, castDetailsObj.column + 3, fileName)


def spider_insidePage(url, page, castDetailsObj):
    page_url = url + 'page/' + str(page) + '/'

    page_source_code = requests.get(page_url)

    page_plain_text = page_source_code.text

    soup_page = BeautifulSoup(page_plain_text, 'html.parser')

    loop = 1
    for link in soup_page.findAll('li', {'role': 'listitem'}):
        movieName = link.text
        indexOfBracket = movieName.find('(') - 1
        movieName = movieName[0:indexOfBracket]
        print(movieName)
        print('url: ' + link.a['href'])
        if loop == 20:
            break

        individual_link = link.a['href'] + '/cast/'
        
        castDetailsObj.increaseRow()
        getImageLink(individual_link, movieName, castDetailsObj)
        castDetailsObj.getCastNames(individual_link, movieName)
        print('-------------------------------------------------------')

        loop += 1


def spider(urlAlphabet, urlAlphabetPages, castDetailsObj):
    url = 'http://www.bollywoodhungama.com/directory/movies-list/alphabet/' + str(urlAlphabet) + '/'

    source_code = requests.get(url)

    plain_text = source_code.text

    soup = BeautifulSoup(plain_text, 'html.parser')

    loop = 1
    for link in soup.findAll('li', {'role': 'listitem'}):
        movieName = link.text
        indexOfBracket = movieName.find('(') - 1
        movieName = movieName[0:indexOfBracket]
        print(movieName)

        print('url: ' + link.a['href'])
        if loop == 20:
            break

        individual_link = link.a['href'] + '/cast/'
        castDetailsObj.increaseRow()
        getImageLink(individual_link, movieName, castDetailsObj)
        castDetailsObj.getCastNames(individual_link, movieName)
        print('-------------------------------------------------------')

        loop += 1

    page = 2
    while page <= ((urlAlphabetPages // 20) + 1):
        spider_insidePage(url, page, castDetailsObj)

        page = page + 1

def main():
    castDetailsObject = CastDetails()

    for i in range(len(alphabets)):
        spider(alphabets[i], alphabets_pages[i], castDetailsObject)

main()