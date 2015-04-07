# MakeE164
Format Mobile and Telephone numbers to E.164 format numbering plan.

An excel macro script to normalize (validate length, remove extra spaces/chars, append country code, convert old number format to new etc) a selection of numbers in excel sheet to the format of a specific country.

Validation of mobile numbers for the below countries are covered:
1.	Kuwait (This is set as the default country to takes care of numbers with no country code.)
2.	Saudi Arabia
3.	UAE
4.	Qatar
5.	Bahrain
6.	Oman
7.	Yemen
8.	Egypt
9.	Lebanon
10.	India
11.	Pakistan

Usage in Office 2007:
Open the excel sheet with the data. Goto Macro under View, give a temporary name and click create. In the editor replace the code with the attached code, save and quit the editor.
Now again goto Macro under View and you will see all the macros listed. Select the Run button to run for normalization for the desired country.
