@echo
call %userprofile%\Anaconda3\Scripts\activate.bat
cd %cd%\Desktop\BarcodeConverter\
SET ruta0=%cd%\BarcodeConverter.py
%userprofile%\Anaconda3\python.exe %ruta0%
PAUSE