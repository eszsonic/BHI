import bibtexparser
import pandas as pd
import webbrowser
import os
import re
import fitz




### Root Directories
year = '2023'
path_current_dir = os.path.dirname(__file__)

path_directory_Excel = r'D:\BHI-main\BHI2023'
Excel_File_Name = 'ieeebhi2023-papers.xlsx'

#bib file from EDAS - may need manual editing to correct import errors
path_directory_Bibtex = r'D:\BHI-main\BHI2023'
Bibtex_File_Name = '30715.bib'

#original files from EDAS
path_directory_Pdf = r'D:\BHI-main\BHI2023\ieeebhi2023-papers'

path_save_directory_pdf = os.path.join(path_current_dir,  r'wwwroot' + '\\' + year + r'\pdfs')
try:
    os.mkdir(path_save_directory_pdf)
except:
    print('')
path_save_directory_bib = os.path.join(path_current_dir,  r'wwwroot' + '\\' + year + r'\bib')
# os.mkdir(path_save_directory_bib)
path_save_directory_html = os.path.join(path_current_dir, r'wwwroot' + '\\' + year)
# os.mkdir(path_save_directory_html)





### Extract Bibtex from bib File
with open(path_directory_Bibtex + '\\' + Bibtex_File_Name, encoding='utf-8') as bibtex_file:
    bib_database = bibtexparser.load(bibtex_file)

    # file.write(BibTex_ID)



### Open the Excel File

df = pd.read_excel(path_directory_Excel + '\\' + Excel_File_Name)

### Inputs from the Excel File for the HTML Template
Title = df['Title']
Abstract = df['Abstract']
Authors = df['Authors']
Keywords = df['Keywords']
Registration = df['Registration']
Paper_Number = df['#']
Copyright = df['Copyright form submitted']
Conference_Title = '2023 IEEE EMBS International Conference on Biomedical and HealthInformatics (BHI) (IEEE BHI 2023)'

### Inputs from the Bibtex
# ArXiv_Link = 'https://www.ieee.org//'
pdf_link = 'https://www.ieee.org//'
# Bib_Address = 'file:///' + 'D:/pdf2Bib/new0.bib'


#List of Links to papers to ba later saved as an index file
linksToPapers=[]


## Create the HTML File
for i in range(len(df)):
    #if registration is paid and IEEE Copyright submitted
    if (pd.isna(Registration[i]) is False):
        if (pd.isna(Copyright[i]) is False):
            # try:
                for j in range(len(bib_database.entries)):
                    new_string = bib_database.entries[j]['title'].replace('\n', ' ')
                    new_string = re.sub(r"[^a-zA-Z0-9 ]", "", new_string)

                    if (new_string == re.sub(r"[^a-zA-Z0-9 ]", "", Title[i])):
                        print('\n', Title[i], 'The number in the Bibtex', j)
                        break

                try:
                    paper_author = bib_database.entries[j]['author']
                except:
                    paper_author = ''
                    print ('\nThere is no Author Information for "', Title[i], '"')

                try:
                    paper_booktitle = bib_database.entries[j]['booktitle']
                except:
                    paper_booktitle = ''
                    print('\nThere is no Booktitle Information for "', Title[i], '"')

                try:
                    paper_address = bib_database.entries[j]['address']
                except:
                    paper_address = ''
                    print('\nThere is no Booktitle Information for "', Title[i], '"')

                try:
                    paper_page = bib_database.entries[j]['pages']
                except:
                    paper_page=''
                    print('\nThere is no Page Information for "', Title[i],'"')

                try:
                    paper_days = bib_database.entries[j]['days']
                except:
                    paper_days=''
                    print('\nThere is no Page Information for "', Title[i],'"')

                try:
                    paper_month = bib_database.entries[j]['month']
                except:
                    paper_month = ''
                    print('\nThere is no Month Information for "', Title[i], '"')

                try:
                    paper_year = bib_database.entries[j]['year']
                except:
                    paper_year = ''
                    print('\nThere is no Year Information for "', Title[i], '"')

                try:
                    paper_keywords = bib_database.entries[j]['keywords']
                except:
                    paper_keywords = ''
                    print('\nThere is no Year Information for "', Title[i], '"')

                try:
                    paper_abstract = bib_database.entries[j]['abstract']
                except:
                    paper_abstract = ''
                    print('\nThere is no Year Information for "', Title[i], '"')

                # with open(str(Paper_Number[i]) + '.bib', "w", encoding="utf-8") as file:
                #     file.write(str(bib_database.entries[i]))

                BibTex_ID = str('@'+bib_database.entries[j]['ENTRYTYPE']+'{'+bib_database.entries[j]['ID']+',\n'
                                'AUTHOR    = {'+paper_author+'},\n'
                                +'TITLE    = {'+bib_database.entries[j]['title']+'},\n'
                                +'BOOKTITLE    = {'+paper_booktitle+'},\n'
                                +'ADDRESS    = {' + paper_address + '},\n'
                                +'PAGES    = {' + paper_page + '},\n'
                                +'DAYS    = {' + paper_days + '},\n'
                                +'MONTH    = {'+paper_month+'},\n'
                                +'YEAR    = {'+paper_year+'},\n'
                                +'KEYWORDS    = {'+paper_keywords+'},\n'
                                +'ABSTRACT    = {' + paper_abstract + '},\n'
                                +'}')

                file_bib = open(path_save_directory_bib+r'/'+str(int(Paper_Number[i])) + '.bib', "w", encoding="utf-8")
                file_bib.write(BibTex_ID)
                file_bib.close()

                #File links
                # Bib_Address = 'file:///' + path_directory_Bibtex + '/' + str(Paper_Number[i]) + '.bib'
                Bib_Address = 'bib/' + str(int(Paper_Number[i])) + '.bib'
                pdf_link = 'pdfs/' + str(int(Paper_Number[i])) + '.pdf'

                ### Save the PDF
                # define the position (upper-right corner)
                image_rectangle = fitz.Rect(0, 0, 610, 40)

                # retrieve the first page of the PDF
                file_handle = fitz.open(path_directory_Pdf + '\\' + str(int(Paper_Number[i])) + '.pdf')
                first_page = file_handle[0]

                img = open("stamp.png", "rb").read()  # an image file
                img_xref = 0

                first_page.insert_image(image_rectangle, stream=img, xref=img_xref)
                file_handle.save(path_save_directory_pdf + '\\' + str(int(Paper_Number[i])) + '.pdf')

                f = open(path_save_directory_html+r'/'+str(int(Paper_Number[i])) + '.html', 'w', encoding="utf-8")
                ### HTML Template
                html_template = '<!DOCTYPE html>\n' + '<html>\n' + '<head>\n' \
                + '<meta http-equiv="content-type" content="text/html; charset=UTF-8">\n' \
                + '<title>' + str(Title[i]) + '.html</title>\n' + '</head>\n' + '<body>\n' \
                + '<h1> '+ str(Title[i]) + '</h1><br>' \
                + '<i><b> ' +str(Authors[i]) + '</b></i><br>'\
                + '<h2>Abstract</h2><br>' \
                + str(Abstract[i]) + '<br><br>' \
                + '<b>Keywords: ' + str(Keywords[i]) + '<br> <br>' \
                + '<h2>Links</h2>' \
                + '[<a href="' + pdf_link + '">Full text PDF</a>]' \
                + '[<a href="' + Bib_Address + '">Bibtex</a>]' \
                + ' <br> ' + '</html>'

                # html_template = '<!DOCTYPE html>\n' + '<html>\n' + '<head>\n'\
                # + '<meta http-equiv="content-type" content="text/html; charset=UTF-8">\n'\
                # + '<title>' + str(Paper_Number[i]) + '.html</title>\n' + '</head>\n' + '<body>\n'\
                # + '<div id="papertitle" style="box-sizing: border-box; font-size: 36px; max_width: 900px; overflow-wrap: break-word; white-space: normal; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">\n'\
                # + str(Title[i]) +'<dd style="box-sizing: border-box; line-height: 1.42857; margin-left: 0px;"><br>\n' + '</dd>\n' + '</div>\n'\
                # + '<div id="authors" style="box-sizing: border-box; max_width: 900px; overflow-wrap: break-word; white-space: normal; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><br\n'\
                # + 'style="box-sizing: border-box;">\n'\
                # + '<b style="box-sizing: border-box; font-weight: bold;"><i style="box-sizing: border-box;">\n' + str(Authors[i])\
                # + '</i></b>; ' + Conference_Title + '</div>'\
                # + '<font style="box-sizing: border-box; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"\n'\
                # +'size="5"><br style="box-sizing: border-box;">' +'<b style="box-sizing: border-box; font-weight: bold;">Abstract</b></font><span\n'\
                # +'style="color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;"></span><br\n'\
                # +'style="box-sizing: border-box; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">\n'\
                # +'<br style="box-sizing: border-box; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">\n'\
                # +'<div id="abstract" style="box-sizing: border-box; max_width: 900px; overflow-wrap: break-word; white-space: normal; text-align: justify; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">\n'\
                # + str(Abstract[i]) + '</div>\n' \
                # + '<div id="keywords" style="box-sizing: border-box; max_width: 900px; overflow-wrap: break-word; white-space: normal; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"><br\n'\
                # + 'style="box-sizing: border-box;">\n'\
                # + '<b style="box-sizing: border-box; font-weight: bold;"><i style="box-sizing: border-box;">\n' + 'Keywords: ' +str(Keywords[i]) + '<br> <br>'\
                # +'<font style="box-sizing: border-box; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;"\n'\
                # +'size="5"><br style="box-sizing: border-box;">\n'\
                # +'<b style="box-sizing: border-box; font-weight: bold;">Related Material</b></font><span\n'\
                # +'style="color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;"></span><br\n'\
                # +'style="box-sizing: border-box; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">\n'\
                # +'<br style="box-sizing: border-box; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">\n'\
                # +'<dd style="box-sizing: border-box; line-height: 1.42857; margin-left: 0px; color: rgb(0, 0, 0); font-family: &quot;Open Sans&quot;, Arial, Verdana, sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; white-space: normal; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial;">[<a\n'\
                # +'href="' + pdf_link + '"\n'\
                # +'style="box-sizing: border-box; background-color: transparent; color: rgb(115, 149, 197); text-decoration: none;">Full text PDF</a>]\n'\
                # +'[<a href="' + Bib_Address +'"\n'\
                # +'style="box-sizing: border-box; background-color: transparent; color: rgb(115, 149, 197); text-decoration: none;">Bibtex</a>]\n' \
                # + ' <br> ' + '\n</html>\n'
                # +'<div class="bibtex" style="box-sizing: border-box; font-weight: normal; text-decoration: none; display: inline; margin-right: 5px;">'\
                # +'</dd>\n' +'<p></p>\n' +'</body>\n' +  '<span style="font-weight: normal">' + '@'+ bib_database.entries[j]['ENTRYTYPE']+'{'+bib_database.entries[j]['ID']+',\n'\
                # + '<br> author    = {'+paper_author+'},\n'\
                # +' <br> title    = {'+bib_database.entries[j]['title']+'},\n'\
                # +' <br> booktitle    = {'+paper_booktitle+'},\n'\
                # +' <br> address    = {'+paper_address+'},\n'\
                # +' <br> pages    = {'+paper_page+'},\n'\
                # +' <br> days    = {'+paper_days+'},\n'\
                # +' <br> month    = {'+paper_month+'},\n'\
                # +' <br> year    = {'+paper_year+'},\n'\
                # +' <br> keywords    = {'+paper_keywords+'},\n'\
                # +' <br> abstract    = {'+paper_abstract+'}\n'\
                #+' <br> }' + '\n</html>\n'

                # writing the code into the file
                f.write(html_template)
                # close the file
                f.close()

                #Add to the list of papers
                linksToPapers.append('<p> <a href = "'+year+'/'+str(int(Paper_Number[i]))+'.html"'+\
                    ' target = "_self" style="text-decoration: none">'+ str(Title[i])+ ' </a> ')
                linksToPapers.append('<br> '+str(Authors[i])+ '<br>')
                linksToPapers.append('[<a href="' +str(year)+'/'+ pdf_link + '" class="class2">pdf</a>]')
                linksToPapers.append('[<a href="' +str(year)+'/'+ Bib_Address + '" class="class2">bibtex</a>]' +'<br></p>')
            # except Exception as e: print(e)
    # else:
    #     print('There is no Registration Record for', Paper_Number[i], '\t', i)


#save the index file
papers_header=r'<!DOCTYPE html> <html> <head>'+\
    '<meta http-equiv="content-type" content="text/html; charset=UTF-8">'+\
    '<title>List of papers</title>'+\
    '<style> '+\
    'a {font-size: 20px} '+\
    'a:link {text-decoration:none;color: #000000;} '+\
	'a:visited {text-decoration:none;color: #000000;} '+\
	'a:hover {font-weight: bold;color: #000000;} '+\
	'a:active {text-decoration:underline;font-weight: bold;color: #000000;} '+\
    'a.class2 {color:blue;} '+\
    'a.class2 {font-size: 12px;} '+\
    'a.class2:link {text-decoration: none; color: blue;} '+\
    'a.class2:visited {text-decoration: none; color: blue;} '+\
    'a.class2:hover {text-decoration: underline; color: blue;} '+\
    'a.class2:active {text-decoration: none; color: blue;} '+\
    '</style> '+\
    '</head> <body> <p><br></p>'

papers_footer=r'<p><br></p></body></html>'
f = open(path_current_dir+'/wwwroot/papers.html', 'w', encoding="utf-8")
f.write(papers_header+'\n')
for i in range(len(linksToPapers)):
    f.write(str(linksToPapers[i]+'\n'))
f.write(papers_footer)
f.close()

## Reading all HTML Files in the Directory
# HTML_files=[]
# for file in os.listdir(r'wwwroot/'+year):
#     if file.endswith(".html"):
#         HTML_files.append(file)

# Open HTML Files
# for i in range(len(HTML_files)):
#     webbrowser.open("file://"+path_current_dir+"\\wwwroot\\"+year+"\\"+HTML_files[i])

print('\n\nFinish!')