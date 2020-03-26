# ExcelPlugin
Borsdata ExcelPlugin is a plugin you install to your local Excel application. (link to install)
When it's installed you get access to serveral new functions to directly call the Börsdata API.
You have access to all data from Borsdata API like Stockprices, reportdata and Kpis.  

To use Borsdata ExcelPlugin you need to be a PRO borsdata member and having an API Key.
To apply for an API Key you can read more here.  

Borsdata ExcelPlugin are an extension to Borsdata API.  
In the background the ExcelPlugin calling the API to get data and formats the data to fit into excel.  
It has the same limitations, like number of requests, as the normal API, so if you have many functions it can take some time to load all data.  

Some functions in ExcelPlugin is not so easy to use and you may need to check the API documentations to understand how to call some data.
We are happy for all feedback on how to improve the API and ExcelPlugin.  

## Get Started  
	1. If you dont have an Api Key you need to apply for one as PRO member of Börsdata. (Link)  
	2. Download Zip file from github and unpack to local folder  
	3. Check if your Excel is installed as 32bit or 64bit. (Link)  
	4. Install plugin in Excel (länk till ny sida)  
	5. Enter API Key in Borsdata Excel plugin  
	6. Done

## API Key
If you dont have an API KEY you need to Apply for it on Börsdata MyPage.
Only Pro Members in Nordic Country can Apply for API Key.  

If you have an API Key and lose the Pro Membership you will be given a new API Key when you restart the Pro membership. Then Please check your accounts MyPage to see the new API Key.

## Instruments (InstId)
An Instrument at Börsdata is normaly a company or index.   
Anything that has a price or report data is an Instrument.  
All Instrument has a uniqe ID called Instrument ID (InstID).  
To get the Instrument ID (InstID) you need to call the BD_INSTRUMENTS() function and check the returning table.

## Stockprice data
All stockprice data is EndDay.  
We do not offer intraday data in Excel Addin or API.

## Info about Report data
https://github.com/Borsdata-Sweden/API/wiki/Reports

## Info about Kpi data
https://github.com/Borsdata-Sweden/API/wiki/KPI  
https://github.com/Borsdata-Sweden/API/wiki/KPI-Screener  


## Dynamic Array
Börsdata Excel Addin can return data to excel in two ways.
	1. Write data in cells around the formula (Old style).
	2. Or as Dynamic Array if you have the latest excel version that support this option.

If your Excel does not support Dynamic Array you will se a text in the support panel.
If your Excel support Dynamic Array you can select if you like to use Dynamic Array.
See movie to get a better description of Dynamic Array.
Link…  

More info  
https://support.office.com/en-us/article/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531

## Support
Make sure to read the documentation - this will avoid repeated questions.
If still having problem then send us an email or direct message on twitter.  

## Issues Tracking
We use GitHub's Issues tracker for our project. Feel free to create bug reports and features requests. Make sure to read the documentation before asking questions - this will avoid repeated questions, leaving us more time for developing the library.


## Disclaimer
Börsdata and the information contained herein is not intended to be a source of advice or credit analysis with respect to the material presented, and the information and/or documents contained in this website do not constitute investment advice.  
The user is responsible for the risk and should therefore acquire information about the terms that apply to trade with such instruments, and about the qualities of the instruments. Placements or other positions in financial instruments such as stock is done at your own risk.  
Börsdata cannot in any way be held responsible for any investment decisions that result from the financial information on the website.
Always do your own Research.  




