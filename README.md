# Data-Comparator

**Version: 1.0.3**

Changelog:

Main Window:
- Removed the 'Hide/Unhide' button.
  - Instead of manually hiding what isn't neccesary per technology, the program will do it automatically!
- Fixed some logic for getting the 'Mode' for WLAN, it should now be working as intended.
- I know some people had issues where the window on startup would go offscreen, thus making it impossible to use.
  - This potentially has been fixed (?). Test it out and see if this fix worked as it works fine on my environment.
- Noticed that the logic to get 50% RB Offset at RB Position: High is wrong. This has been fixed.

Comparator:
- Added the same automatic hide function to this. It work once you have loaded data in and pressed 'Prepare to Be Sad'.

Other Changes:
- Removed the 'Hide/Unhide' code.
- Added some stand-in code for a future addition to this program.


----


This programs function is to compare the data from the plots and excel sheets and too see if they match.

This program DOES support the following technologies:
- GSM (GPRS only)
- W-CDMA (Rel 99 only)
- LTE (QPSK only)
- FR1 5G NR (DFT-s-OFDM pi/2 BPSK only)
- Bluetooth (GFSK (BDR) only)
- WLAN (2.4 GHz and 5 GHz all modes)

This program DOES NOT support the following technologies:
- NFC
- FR2 5G NR
- Any mmWave (Power density) testing
- WLAN (6E and 7)
- UWB (lol)


----


# Table of Contents
- [HOW TO USE](#how-to-use)
  - [1) LOADING IN AN SHEET FROM EXCEL:](#1-loading-in-an-sheet-from-excel)
  - [2) LOADING IN A MICROSOFT DOCX:](#2-loading-in-a-microsoft-docx)
  - [3) BOTH EXCEL AND DOCX LOADED:](#1a-and-2a-both-excel-and-docx-loaded)
- [1) COMPARE (In Beta)](1-compare-in-beta)
- [2) HIDE/UNHIDE](2-hide-unhide)


----


## HOW TO USE

When the program is opened, you will be greated with the bottom picture:
![image](https://github.com/ul52300/Data-Comparator/assets/148300863/50ce371b-ab0e-482a-9362-aa23a308d831)

You are able to select the Excel sheet and Microsoft Docx individually if you want to just check either or both.


----


### 1) LOADING IN AN SHEET FROM EXCEL:

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/5b59a9f5-e28e-4a36-95f3-9f0ffc2bd103)

By clicking on **'Browse'** you can browse your Windows filebrowser to find the Excel with the data that you want to check.

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/20fd6c19-2cdf-44e5-9324-cd82e9702459)

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/7eb282d6-7ac3-4075-bf4b-c64dd537d68b)

**(NOTE: You MUST use an Excel that is not heavily modified from the original template.)**

Once you have selected your Excel document, you will now have to load in the sheets of the Excel. These sheets will represent what technology and band you want to check. (I.E. GSM 850, W-CDMA BV, LTE B5, etc.)
Please click on **'Load'**. This will load your sheets into the dropdown list, where you will select the technology and band.

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/60e7b4c4-cf81-4fa3-bda6-532d7479d72e)

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/e2c4c46b-c0c3-4237-92b9-8543015a1088)

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/73d6c81b-6bed-4def-a51d-a6118c167088)

**(NOTE: You MUST hit 'Load' every time you enter a new Excel. If you don't the program will crash because it thinks you are using the previous Excel's sheets when you are using one that potentially does not exist in the newly selected Excel.)**

After you have entered BOTH the Excel and loaded in the sheet you want to see, you can now display the data on the table by hitting **'Load Excel'**.

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/a3107746-721a-452e-8efb-9d1fb7351f29)

This will be the final result if everything was entered correctly:

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/1b98b619-c9b2-4ea7-812f-596671199f3d)


----


### 2) LOADING IN A MICROSOFT DOCX:

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/4de7fad7-9717-4fc7-aa83-f8c814cea0b8)

By clicking on **'Browse'** you can browse your Windows filebrowser to find the Microsoft Docx with the data that you want to check.

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/75442aba-d591-4be9-81c7-09107464943e)

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/bb9059a8-d2fe-4a2b-b1f3-cbb17b3a100e)

**(NOTE: The Microsoft Docx with the plots MUST be formatted to the standard way that the SAR department plots data.)**

Once you have selected your Microsoft Docx you can now load in the data into the table. There is no need to load the technology and band since this is all in the plot.

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/7c68bfd7-f10a-4a18-84e7-6ebe22e3e3f0)

This will be the final result if everything was entered correctly:

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/fd401924-558d-444a-b96f-9c3858588f5b)


----


### 3) BOTH EXCEL AND DOCX LOADED:

Here is what the tables look like when both Excel and Microsoft Docx are loaded:

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/a160cdf0-78ce-486f-8a84-6fc983f5bf9d)

From here you can compare the two and see if there are any discrepancies.


----


## OTHER FEATURES

### 1) COMPARE (In Beta)

Hitting **'Compare'** opens the following window:

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/77e2070c-f12b-42ce-bf32-a729140ff919)

After hitting **'Prepare to Be Sad'**, the table will now fill up with the plot numbers and whether each plot matches the 1g and 10g SAR.

![image](https://github.com/ul52300/Data-Comparator/assets/148300863/499df448-498b-41b3-b29d-b9c6f7f0da1a)

**NOTE: You'll see that there is somewhat of an error with this comparison. Notice that there is 18 tests in the Excel, but there is 17 plots, this means one of two things: (1) The person added extra data to the Excel or (2) The person forgot to plot the data. However, you'll see that it still says 17 plots and they all match. This happens because the program is comparing all the data that matches, but if the number of tests for both don't match up, well you already found an error, hence why all the comparisons say 'Yes'. In this case just remove the outlier and just compare the tests that exists in both documents.**

**Please keep in mind that you cannot interact with the main window when this window is open!**


----


### 2) HIDE/UNHIDE

Hitting **'Hide/Unhide'** will remove any unnessesary columns for whatever technology that you are viewing currently. This must be done AFTER you have loaded your data in.

Here is the unfiltered tables
![image](https://github.com/ul52300/Data-Comparator/assets/148300863/ba7bf2e8-6ab2-4199-b7e2-bf6f8b0e2696)

and here is the filtered tables
![image](https://github.com/ul52300/Data-Comparator/assets/148300863/63f0229f-e04d-42dd-91ab-7ff42bddffa1)

Notice the removal of 'RB Allocation', 'RB Offset', and 'Max Area (W/kg)'. These are NOT needed for GSM.

The following is a list of the technologies that we currently test and what **'Hide/Unhide'** removes:
- GSM removes:
  - RB Allocation
  - RB Offset
  - Max Area (W/kg)
- W-CDMA removes:
  - RB Allocation
  - RB Offset
  - Max Area (W/kg)
- LTE removes:
  - Max Area (W/kg)
- FR1 (DFT-s-OFDM pi/2 BPSK)
  - Max Area (W/kg)
- WLAN 2.4/5 GHz removes:
  - RB Allocation
  - RB Offset
- Bluetooth removes:
  - RB Allocation
  - RB Offset
  - Max Scan (W/kg)
- NFC (Not supported by this program :c)
- WLAN 6E (Not supported by this program :c)
- FR2 (Not supported by this program :c)
- UWB (Not supported by this program :c)
