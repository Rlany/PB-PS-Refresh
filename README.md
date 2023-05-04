# PB-PS-Refresh
Powershell script to refresh Power BI desktop via powershell. Calling windows api.

Script does following steps:
1.	Kill PB if it’s opened with new file name.
2.	Copy pbix workbook under new name.
3.	Open PB Desktop with new file and wait until it’s open.
4.	Hit the refresh button.
5.	Wait until refresh is finished.
6.	Hit the Save button. 
7.	Close PB Desktop

If you download the demo, put all files into one folder and you should be able to run ps script and it will create new Demo_New.pbix file and refresh it. Demo file is connected to website with time so you should see in new file different time then in the old one without refresh.