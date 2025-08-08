from fastapi import FastAPI, Body
import json
import logging
from utils import printToPrinter, get_available_printer_names, get_active_printer

# logging.basicConfig(filename='printserver.log', filemode='a', format='%(asctime)s %(name)s %(levelname)s %(message)s', encoding='utf-8', level=logging.DEBUG)
log_handler = logging.FileHandler(filename='printserver.log', encoding='utf-8')
logging.basicConfig(handlers=[log_handler], level=logging.DEBUG)
logging.debug("start")
logger = logging.getLogger('uvicorn.error')
logger.setLevel(logging.DEBUG)


app = FastAPI()


@app.get("/print")
async def root():
    return {"message": "Hello World"}


@app.post("/print")
async def print(print_obj: str = Body(...)):
    logger.info(f"printing to {get_active_printer()}")
    logging.info("printing")
    obj = json.loads(print_obj)
    logging.debug(f'{obj = }')
    printToPrinter(obj, get_active_printer())
    logger.debug(f'{obj = }')
    logging.info(f"pklist_ID {obj.get('pklist_ID')}: Done")
    return obj




