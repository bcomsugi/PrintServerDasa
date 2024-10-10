cd "C:\Sources\DBW\printserverdasa"
echo already cd
cmd /k "cd /d C:\Sources\DBW\printserverdasa\env\Scripts\ & activate & cd /d C:\Sources\DBW\printserverdasa\ & uvicorn main:app --host 0.0.0.0 --port 9999