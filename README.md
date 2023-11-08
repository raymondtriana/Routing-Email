# Routing-Email
An Email automation tool to notify hundreds of recipients each day based upon pdf files.

# Requirements
<ul>
    <li>Python 3.10.11 or greater</li>
    <li>
    PyPDF2    
    </li>
    <li>
    pywin32    
    </li>
</ul>

# Order of execution
<ul>
    <li>
    loop through each pdf in the ./pdf folder    
    </li>
    <li>
    parse the pdf looking for the driver name, driver ID, and route info
    </li>
    check if this information is found within ./cache
    <li>
    if the information was not cached then cross reference driver ID with driverEmailList.csv
    </li>
    <li>
    send email with this pdf as attachment to the driver email
    </li>
    <li>
    add the information to ./cache/[pdf].txt
    </li>
</ul>
