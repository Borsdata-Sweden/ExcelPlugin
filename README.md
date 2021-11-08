# Börsdata Excel Plugin
[Börsdata ExcelPlugin](https://borsdata.se/en/info/api/excelplugin) is a plugin/addin you install to your Windows Excel.  
When it's installed you get access to serveral new Excel functions to directly call the [Börsdata API](https://borsdata.se/en/info/api/api_page).   
You have access to all data from Börsdata API like Stockprices, reportdata and Kpis.  

To use Börsdata ExcelPlugin you need to be a PRO Börsdata member and having an API Key.   
API Key is only available to 12/24 month PRO members.   

Börsdata ExcelPlugin are an extension to Börsdata API.  
In the background the ExcelPlugin calling the API to get data and formats the data to fit into excel.  
It has the same limitations, like number of requests, as the normal API, so if you have many functions it can take some time to load all data.  

We are happy for all feedback on how to improve the API and ExcelPlugin.  

## V7 Excel upgrade fix
Microsoft made a new Excel upgrade that broke the plugin.(Excel version 16.0.13901.20366)
This is a fix that needs to be installed if you get this problem.  


## V5 Breaking Changes
- Adding more rows into Report and Summary Data may break your existing Spreadsheet.
- Report and Summary has one new datarow at top (Revenue, Total intäkter) and ReportDate at bottom.


## Supported Excel versions
Only Windows Desktop versions is Supported.   
* Sorry - ExcelPlugin on Mac is not working.    
* Sorry - ExcelPlugin on Office 365 Webb (Addins) is not working.  

Supported Windows Desktop versions
- Excel 2013 (v15)
- Excel 2016 (v16)
- Excel in office 365

The reason is that .xll files is a low-level windows plugin.    
Microsoft only support this on windows.


## Excel plugin for MAC is NOT working
Sorry - Börsdata ExcelPlugin on Mac is not working.  
Excel on Mac is not running on same core-code as Windows and Microsoft does not support .xll plugin on Mac.  
On Mac its not possible to write plugins to Excel.  
This is a Microsoft (Mac?) limitation.  


## Installation  
[New Installation](https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Installation)   
[Upgrade to new version](https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Installation#upgrade-to-new-version-of-excel-plugin)     
[Problems? Check our page for common errors.](https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Error-help)     

## Youtube (in Swedish)
[Installation Movie](https://youtu.be/Mt_mo4VXoFI)   
[Functions](https://youtu.be/tVLwrZ9tl3w)  

## ExcelPluginSamples
Dynamic Array has poor backward compatibility with older excel versions.    
If you have a new Excel version with Dynamic Array then use ExcelPluginSamples_DynamicArray.xlsx.   
If you dont have Dynamic Array then use the ExcelPluginSamples.xlsx.  
If you open ExcelPluginSamples_DynamicArray.xlsx and see all functions like {=BD_INSTRUMENTS()} then you missing Dynamic Array and need the other file.  
[Microsoft page about backward compatibility](https://support.microsoft.com/en-us/office/dynamic-array-formulas-in-non-dynamic-aware-excel-696e164e-306b-4282-ae9d-aa88f5502fa2)  
## VBA/Macro Support
Börsdata Excel plugin has poor support for VBA and Macro.  
The new Börsdata functions is not built or tested for VBA.  
Please use pure Excel formulas with Börsdata Excel plugin.  

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

## Support / Problem with Excel Plugin?
Make sure to read the documentation - this will avoid repeated questions.   
[Check our page for common errors.](https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Error-help)   
If still having problem then send us an email (info@borsdata.se) or direct message on twitter.   

## Issues Tracking
We use GitHub's Issues tracker for our project. Feel free to create bug reports and features requests. Make sure to read the documentation before asking questions - this will avoid repeated questions, leaving us more time for developing the library.

## Version

### 2020-04-10 V1  
- Release

### 2020-04-17 V2  
- Fix Stockprice columns order   
- Add BD_INSTRUMENT_SEARCH()

### 2020-04-23 v3
- "Remove Range" parameter for NO-DynamicArray users. Helper to delete old data in function.

### 2020-05-04 v4
- Faster API calls. 

### 2020-11-10 V5
- Moved Spreadsheet logic to Backend.
- Bug fixes. Datetime Format and Summary Quarter.
- Add ReportDate
- Faster API calls

### 2020-11-17 v6
- Added Refresh All sheets Button
- Added Refresh Pivottable Button

## KNOWN ISSUES

### Excel 2010 (v14)
- Refresh button is not working correct. The API is only called once.  Workaround is to save file and restart excel. Then refresh makes a new Api call.

### Poor VBA support
Börsdata Excel plugin has poor support for VBA and Macro.  
The new Börsdata functions is not built or tested for VBA.  
Please use pure Excel formulas with Börsdata Excel plugin.  

Also check our error page.
https://github.com/Borsdata-Sweden/ExcelPlugin/wiki/Error-help


## Disclaimer
Börsdata and the information contained herein is not intended to be a source of advice or credit analysis with respect to the material presented, and the information and/or documents contained in this website do not constitute investment advice.  
The user is responsible for the risk and should therefore acquire information about the terms that apply to trade with such instruments, and about the qualities of the instruments. Placements or other positions in financial instruments such as stock is done at your own risk.  
Börsdata cannot in any way be held responsible for any investment decisions that result from the financial information on the website.
Always do your own Research.  


