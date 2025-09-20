from fastapi import FastAPI, Body, Request
import json
# import logging
from utils import printToPrinter, get_available_printer_names, get_active_printer, check_redis_cache, set_redis_cache
import os
import redis
from dotenv import load_dotenv
from loguru import logger

logger.add("logs\loguru.log", rotation="1 day", retention="1 week", format="<green>{time:YYYY-MM-DD HH:mm:ss.SSS}</green> | <level>{level}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level> | {extra}")

logger.info("STARTED")

# logging.basicConfig(filename='printserver.log', filemode='a', format='%(asctime)s %(name)s %(levelname)s %(message)s', encoding='utf-8', level=logging.DEBUG)
# log_handler = logging.FileHandler(filename='printserver.log', encoding='utf-8')
# logging.basicConfig(handlers=[log_handler], level=logging.DEBUG)
# logging.debug("start")
# logger = logging.getLogger('uvicorn.error')
# logger.setLevel(logging.DEBUG)


# Get the absolute path of the directory containing the current script (main.py)
current_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the path to the .env file in the child directory
dotenv_path = os.path.join(current_dir, '.venv', 'win7_prod.env')
print(dotenv_path)
load_dotenv(dotenv_path=dotenv_path)
# print(os.environ)
# print(f'{os.getenv("REDIS_SERVER") = }')
redis_server_url = os.getenv("REDIS_SERVER")
db = int(os.getenv('DB', "15"))
redis_client = redis.Redis(host=redis_server_url, port=6379, db=db)


app = FastAPI()


@app.get("/print")
async def root():
    return {"message": "Hello World"}

@app.post("/print")
# async def print(print_obj: str = Body(...)):
async def print(request:Request):
    # print_obj = await request.body()
    print_obj = await request.json()
    # print(f't{print_obj = }')
    obj = json.loads(print_obj)
    logger.debug(f'{obj = }')
    user = obj.get("User","")
    
    # print(f'{user = }')
    logger.info(f"printing to {get_active_printer(user)}")
    is_cached = check_redis_cache(obj)
    # dt = obj.get("DT","")
    # if not dt:
    #     return False
    # cache_key = f"PLPRINT:{dt}"
    # data = redis_client.get(cache_key)
    # print(f'{cache_key = } {data = } {type(data)=}')
    # if data:
    #     print(f"FOUND {cache_key = } {data = }")
    #     return json.loads(data)
    # print(f'{is_cached = }')
    if is_cached:
        logger.warning(f"cache is FOUND {is_cached = }")
        return is_cached
    result = printToPrinter(obj, get_active_printer(user))
    logger.debug(f'{obj = }')
    logger.info(f"pklist_ID : {obj.get('pklist_ID')}: Done")
    # print(f'{obj = } {type(obj)=}')
    set_redis_cache(obj)
    # redis_client.set(cache_key, json.dumps(obj), ex=72000)
    return obj

@app.post("/print1")
async def print1(print_obj: str = Body(...)):
    obj = json.loads(print_obj)
    logger.debug(f'{obj = }')
    user = obj.get("User","")
    # print(f'{user = }')
    logger.info(f"printing to {get_active_printer(user)}")
    dt = obj.get("DT","")
    if not dt:
        return False
    cache_key = f"PLPRINT:{dt}"
    data = redis_client.get(cache_key)
    print(f'{cache_key = } {data = } {type(data)=}')
    if data:
        print(f"FOUND {cache_key = } {data = }")
        return json.loads(data)
    result = printToPrinter(obj, get_active_printer(user))
    logger.debug(f'{obj = }')
    logger.info(f"pklist_ID : {obj.get('pklist_ID')}: Done")
    print(f'{obj = } {type(obj)=}')
    # redis_client.set(cache_key, json.dumps(obj), ex=72000)
    return obj




