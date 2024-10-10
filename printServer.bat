cd "C:\Users\sugi\Documents\Python Projects\DBW\PrintServerDasa"
echo already cd
cmd /k "cd /d C:\Users\sugi\Documents\Python Projects\DBW\PrintServerDasa\env\Scripts\ & activate & cd /d C:\Users\sugi\Documents\Python Projects\DBW\PrintServerDasa\ & uvicorn printserver:app --host 0.0.0.0 --port 8007