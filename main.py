import requests
import time
import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")
url = ('https://newsapi.org/v2/top-headlines?country=in&apiKey=bee8cc25aa844bc8bceff7d06367d0a8')
def make_url(user,country='in'):
       category_list={'3':'business','4':'sports','5':'health','6':'entertainment'}
       if user=='2':
              country='us'
       if user == '1' or user=='2':
              url = (f'https://newsapi.org/v2/top-headlines?country={country}&apiKey=bee8cc25aa844bc8bceff7d06367d0a8')
              getarticals(url)
       else:
              category= category_list[user]
              url= (f'https://newsapi.org/v2/top-headlines/sources?category={category}&apiKey=bee8cc25aa844bc8bceff7d06367d0a8')
              getarticals(url)
def getarticals(url):
       response = requests.get(url)
       try:
              articles= response.json()['articles']
              for i,_ in enumerate(articles):
                     title= response.json()['articles'][i]['title']
                     author= response.json()['articles'][i]['author']
                     description= response.json()['articles'][i]['description']
                     url= response.json()['articles'][i]['url']
                     print(f"{title}\nAuthor: {author}\n{description}\nLink:- {url}\n\n")
                     time.sleep(.10)
       except:
              sources= response.json()['sources']
              for i,_ in enumerate(sources):
                     name= response.json()['sources'][i]['name']
                     description= response.json()['sources'][i]['description']
                     url= response.json()['sources'][i]['url']
                     print(f"{name}\n{description}\nLink:- {url}\n\n")
                     time.sleep(.10)
       speaker.Speak('Here are todays articles')

print("Welcome to our News Channel")
speaker.Speak("Welcome to our News Channel")
while True:
       print('What type of news You want to see?')
       speaker.Speak("What type of news You want to see?")
       print('1.Top Headings 2.US News 3.Business 4.Sports 5.Health 6.Entertainment ')
       speaker.Speak('1 for Top Headings 2 for US News 3 for Business 4 for Sports 5 for Health 6 for Entertainment')
       user= input("")
       
       if user.isdigit():
              if 1 <= int(user) <=6:
                     make_url(user)
                     break
              else:
                     print("That's not a vaild number! try again")
                     speaker.Speak("That's not a vaild number!")
       else:
              print("That's not a number! try again")
              speaker.Speak("That's not a number!")
       
