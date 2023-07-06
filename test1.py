import requests

url = 'http://127.0.0.1:8080/test1'
# headers = {}
r = requests.get(url)
print(r.json())