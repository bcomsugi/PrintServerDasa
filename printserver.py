from fastapi import FastAPI, Body
import json
import logging
from utils import printToPrinter, get_available_printer_names

logging.basicConfig(filename='printserver.log', filemode='a', format='%(asctime)s %(name)s %(levelname)s %(message)s', encoding='utf-8', level=logging.DEBUG)
logging.debug("start")
logger = logging.getLogger('uvicorn.error')
logger.setLevel(logging.DEBUG)


app = FastAPI()


@app.get("/print")
async def root():
    return {"message": "Hello World"}


@app.post("/print")
async def print(print_obj: str = Body(...)):
    logger.info("printing")
    logging.info("printing")
    obj = json.loads(print_obj)
    printToPrinter(obj)
    logger.debug(f'{obj = }')
    logging.debug(f'{obj = }')
    return obj




