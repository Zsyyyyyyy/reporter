from fastapi import FastAPI

app = FastAPI()

@app.get('/test1')
def say_hello():
    return {'code': 200, 'message': 'hello, world!'}