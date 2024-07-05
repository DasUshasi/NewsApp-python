import requests
import win32com.client as wc

#speak
speaker=wc.Dispatch("SAPI.SpVoice")

# get news
def news(myapikey,mycategory,mycountry):
    url=f"https://newsapi.org/v2/top-headlines"
    parameters={
        "apiKey":myapikey,
        "category":mycategory,
        "country":mycountry,
    }
    response=requests.get(url,parameters)
    page=response.json()
    list=page["articles"]
    return list

# print headlines
def printheadlines(list,i,sp):
    start=5*i
    if start+5>=len(list):
        end=len(list)
    else:
        end=start+5
    index=1
    for i in range(start,end):
        s=list[i]["title"]
        print(f"{index}. {s}")
        if(sp=="yes"):
            speaker.speak(s)
        index+=1
    if end>=len(list):
        print("No more headlines on this topic for now!")

# print more
def moreinfo(list,i):
    index=int(input("Enter the serial number: "))
    index=5*i+index-1
    s=list[index]["description"]
    if s!="None":
        print(s)
    print("For more information, visit the following link: ")
    print(list[index]["url"])
    

# get country codes
with open('countrycodes.txt') as f:
    lines=f.readlines()
dict={}
for line in lines:
    list=line.split(':')
    list[0]=list[0].strip()
    list[1]=list[1][1:3]
    dict[list[0]]=list[1]

# get user inputs
myapikey="9d45a0d2ef3c4e1db30729de0a5af153"
print("Which country's news you want to know? Here are your options: ")
[print(item) for item in dict.keys()]
country=input("Enter country: ")
mycountry=dict[country]
mycategory=input("Which category you want your news from? Your options are: \nBusiness \nEntertainment \nGeneral \nHealth \nScience \nSports \nTechnology \nEnter choice: ")
list=news(myapikey,mycategory,mycountry)
sp=input("Do you want the program to speak the headlines to you? Enter \"yes\" or \"no\": ")

# print
i=0
printheadlines(list,i,sp)
while True:
    more=input("Enter \"news\" if you want more headlines, \"info\" if you want to read more about a headline, \"none\" if you want neither: ")
    if(more=="news"):
        i+=1
        printheadlines(list,i,sp)
    elif(more=="info"):
        moreinfo(list,i)
    elif(more=="none"):
        break
    else:
        continue
