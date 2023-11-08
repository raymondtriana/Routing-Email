# Routing-Email
An Email automation tool to notify hundreds of recipients each day based upon pdf files.

# Requirements
Python 3.10.11 or greater
PyPDF2
pywin32

# Order of execution
    -loop through each pdf in the ./pdf folder
    -parse the pdf looking for the driver name, driver ID, and route info
    -check if this information is found within ./cache
    -if the information was not cached then cross reference driver ID with driverEmailList.csv
    -send email with this pdf as attachment to the driver email
    -add the information to ./cache/[pdf].txt
