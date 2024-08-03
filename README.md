# PowerPoint Report Automation
Here you can find an automated reporting tool for HR functions that complies with Mandatory Disclosures S1-7 and S1-16 of the European Sustainability Reporting Standards. 
Using the mock SQL Employees Sample Database (available at https://dev.mysql.com/doc/employee/en/) on the employees of a software company, the Python script fetches the 
relevant data, analyses it and passes the results into a PowerPoint presentation. The script can be set to run periodically (i.e. every two months) via a scheduling tool 
like Windows Task Scheduler, and can be adjusted to respond to a different database structure or to report alternative metrics relevant to the company. 

In this instance, it recovers some basic aspects like current number of employees, average and total salary mass, distribution of employees and average salary across 
departments and titles, and evolution in the number of employees in total and across departments. It also incorporates some metrics relevant to the gender gap in the 
workforce, examining the differential evolution in the number of male and female employees as well as the gender differences in average starting salary and in salary 
progression across time (via a polynomial regression).

In the context of ESG reporting, monitoring the gender gap is essential to make sure that the company plans oriented to correct it are on tracks to meet targets. The 
topic lies at the intersection of SDGs 5 (Gender equality) and 8 (Fair Payment and Living Wage), which in the European Sustainability Reporting Standards are referenced 
by components S1 and S3.
