# XLSX-to-KML

Demonstrates reading XLSX file in PowerShell and using the data to create a KML file.

Depending on how you download the files you may need to unblock the DLLs:
https://blogs.msdn.microsoft.com/delay/p/unblockingdownloadedfile/

You may also need to adjust the execution policy in order to be able to run Powershell scripts on your machine. As administrator:
```
set-executionpolicy RemoteSigned
```
If you don't have administrator rights, you can run the script with
```
powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile -WindowStyle Hidden -File <script_name>
```