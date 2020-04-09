# Börsdata Excel Plugin
Börsdata ExcelPlugin is a plugin/addin you install to your Windows Excel.  
When it's installed you get access to serveral new Excel functions to directly call the Börsdata API.   
You have access to all data from Börsdata API like Stockprices, reportdata and Kpis.  

To use Börsdata ExcelPlugin you need to be a PRO Börsdata member and having an API Key.

Börsdata ExcelPlugin are an extension to Börsdata API.  
In the background the ExcelPlugin calling the API to get data and formats the data to fit into excel.  
It has the same limitations, like number of requests, as the normal API, so if you have many functions it can take some time to load all data.  

We are happy for all feedback on how to improve the API and ExcelPlugin.  

## Supported Excel versions
Only windows versions.   
Sorry - ExcelPlugin on Mac is not working.  
- Excel 2013 (v15)
- Excel 2016 (v16)
- Excel in office 365

## Installation  
[New Installation](https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Installation)   
[Upgrade to new version](https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Installation#upgrade-to-new-version-of-excel-plugin)

## Youtube (in Swedish)
[Installation Movie](https://youtu.be/Mt_mo4VXoFI)   
[Functions](https://youtu.be/0ylLc4Dl4h8)  

## ExcelPluginSamples
Dynamic Array has poor backward compatibility with older excel versions.    
If you have a new Excel version with Dynamic Array then use ExcelPluginSamples_DynamicArray.xlsx.   
If you dont have Dynamic Array then use the ExcelPluginSamples.xlsx.  

## API Key
If you dont have an API KEY you need to Apply for it on Börsdata MyPage.  
Only Pro Members in Nordic Country can Apply for API Key.    

If you have an API Key and lose the Pro Membership you will be given a new API Key when you restart the Pro membership.   
Then Please check your accounts MyPage to see the new API Key.

[Go to Api Info](https://borsdata.se/en/info/api/api_page)

## Functions
[Go to Function page](https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Functions
)

## Plugin Buttons
[Go to button page](https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Excel-Buttons)

## Instruments (InstId)
An Instrument at Börsdata is normaly a company or index.   
Anything that has a price or report data is an Instrument.  
All Instrument has a uniqe ID called Instrument ID (InstID).  
To get the Instrument ID (InstID) you need to call the BD_INSTRUMENTS() function and check the returning table.

## Stockprice data
All stockprice data is EndDay.  
We do not offer intraday data in API or Excel Plugin.

## Info about Report data
https://github.com/Borsdata-Sweden/API/wiki/Reports

## Info about Kpi data
https://github.com/Borsdata-Sweden/API/wiki/KPI  
https://github.com/Borsdata-Sweden/API/wiki/KPI-Screener  


## Dynamic Array
Börsdata Excel Plugin can return data to excel in two ways.  
	1. Write data in cells around the formula (Old style).  
	2. Or as Dynamic Array if you have the latest excel version that support this option.  

If your Excel does not support Dynamic Array you will se a text in the support panel.
If your Excel support Dynamic Array you can select if you like to use Dynamic Array.
See instruction movie to get more details about Dynamic Array.

More info  
https://support.office.com/en-us/article/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531

## Language support
You can select if you like English or Swedish translation.  
You find checkbox in [Settings panel](https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Excel-Buttons).

## Date format 
In first version the Börsdata Excel plugin returns unformated data to excel cells.  
This means that Date is returned as a numeric value and you need to select cells and format to datetime.  
(Internaly in excel all dates is a numeric value)   
We will work on a fix for this in later updates.  

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

## Version
2020-04-03 V1  


## KNOWN ISSUES

### Excel 2010 (v14)
- Refresh button is not working correct. The API is only called once.  Workaround is to save file and restart excel. Then refresh makes a new Api call.




