cd "D:\Sources\DBW\PrintServerDasa"
echo already cd
cmd /k "cd /d D:\Sources\DBW\PrintServerDasa\env\Scripts\ & activate & cd /d D:\Sources\DBW\PrintServerDasa\ & uvicorn printserver:app --host 0.0.0.0 --port 8007