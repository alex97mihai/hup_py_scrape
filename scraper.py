from urllib.request import urlopen
import unicodedata
import pandas as pd

df = pd.read_excel(io="sample_box_office.xlsx", sheet_name="Sheet1")
file_short = []
codes = []

file = [df.columns.values.tolist()] + df.values.tolist()

# Create shorter file with 3 columns: movie title, cinemagia code, generated URL
# remove signs and replace spaces with dashes
for row in file:
    file_short.append([row[3], row[6],"https://www.cinemagia.ro/filme/"+
                       row[3]
                      .replace("'","")
                      .replace(":","")
                      .replace(".","")
                      .replace("(","")
                      .replace(")","")
                      .replace(",","")
                      .replace(" ", "-")
                      .lower()+'-'+str(row[6])+"/"])
# remove header
file_short.pop(0)

# function to extract IMDB code from Cinemagia movie page
def getcode(url):
    # replace special characters with ASCII equivalent (Polish, Hungarian, Romanian etc.)
    url=unicodedata.normalize('NFKD', url).encode('ascii','ignore').decode("utf-8")
    print(url)
    # download page html code
    html = urlopen(url).read().decode("utf-8")
    # find keyword (imdb site) in html
    index=html.find("https://www.imdb.com/title/")+len("https://www.imdb.com/title/")
    # extract code
    code=html[index:index+9]
    return([url,code])

# get codes for all movies in list
for row in file_short:
    code = getcode(row[2])
    codes.append(code)
    print("Loading...")

# add codes to original file
file[0].append("IMDB")
for i in range(0,len(codes)):
    file[i+1].append(codes[i][1])

# convert list back to dataframe and output to excel
pd.DataFrame(file).to_excel('output.xlsx', header=False, index=False)
print("Success!")