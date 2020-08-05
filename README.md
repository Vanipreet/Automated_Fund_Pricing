# Automated_Fund_Pricing

### [Introduction](#introduction-1)
### [Assumptions](#assumptions-1)
### [Automated_Daily_Fund_Pricing_processflow diagram](#Automated_Daily_Fund_Pricing_processflow_diagram-diagram-1)
### [Steps to run the macro](#steps-to-run-the-macro-1)
### [Forward](#forward-1)

# Introduction
The purpose of [Daily_Fund_Pricing.xlsm](https://github.com/Vanipreet/Automated_Fund_Pricing/blob/master/Fund%20Pricing/Daily_Fund_Pricing.xlsm) is to create an automated process for the daily fund pricing. The purpose of fund pricing is to faciliate updated NAV calculation for the fund(s). In real world the prices are obtained from third party prices provider such as Bloomberg, ICE, however to generalise the process it is assumed that the price input files is available as an external input source. There are multiple assumptions and pre-requisites that has been kept in place for the smooth execution of this macro.


# Assumptions
This macro is built with few assumptions in mind as listed below

1. For the macro to run, we need two external data inputs, Daily Pricing Sheets and Daily Fund Instrument list
2. The macro is prepared for the purpose of daily fund pricing to faciliate NAV calculation
3. Macro and the data inputs are available under "H:\Fund Pricing" directory, however as per the drive parition, this can be appropriately mapped
4. Macro expects static input data. However with appropriate mapping, database can be also accurately utilized
5. In real world the prices are obtained from third party prices provider such as Bloomberg, ICE, however to generalise the process it is assumed that the price input files is available as an external input source
Please note: This macro is a pre-requisite for carring out NAV calculations, hence in real world pricing of fund should be incorporated before NAV calculation is initiated

# Automated_Daily_Fund_Pricing_processflow_diagram
![alt text](https://github.com/Vanipreet/Automated_Fund_Pricing/blob/master/ProcessFlow%20Diagram%20Fund_Pricing.png)


# Steps to run the macro

1. Download the zip file [Fund Pricing](https://github.com/Vanipreet/Automated_Fund_Pricing/tree/master/Fund%20Pricing) onto your preferred location and unzip the file there
2. Open "Daily_Fund_Pricing.xlsm" workbook
3. Right click on macro button "Run Today's Pricing"
4. From the drop down options click on "Assign Macro"
5. Select "Daily_Fund_pricing" and click on "Edit" button
6. Under module 1, map the Pricing files, Fund Instrument file and Updated_Pricing (output) file directory as per requirement
Please Note: All the mapping is required only on the Module 1
7. For test run, change the file name "EODPrices_ETFs_05082020.xlsx" to current Date, Month and Year in DDMMYYYY format. Save the file and close
8. Redo the Step 7 for "EODPrices_Equities_05082020.xlsx" and "Fund_instrument_05082020.xlsx" as well. 
9. Click on "Run Today's Pricing" macro button under Control sheet of "Daily_Fund_Pricing.xlsm"
10. When the macro has completed processing, new file will be created under Updated_Pricing folder


# Forward

This macro is a starting step for Daily Fund Pricing, utilizing some tweaks this can accurately fulfil Daily fund pricing requirement.

However the bigger picture is that instead of making this macro to work on a single fund, it can be accurately "Looped" for multiple funds.

This macro can be mapped and can directly call NAV calculation macro so the pricing and NAV calculation can be completed by a single click on the macro button

If this macro is to run on windows system, "Task Scheduler" can be utilized to automatically update the input files at a precribed time from the source drive and also to automatically run this macro so fund pricing can be calculated without any human intervention. (Apple systems might also have something similar to task scheduler that can be utilized)
