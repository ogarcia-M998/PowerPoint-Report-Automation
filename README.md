# HR Reporting Tool
## Overview
The HR Reporting Tool is designed to automate the creation of a periodic PowerPoint report based on data from a mock SQL relational database named 'Employees' (available under https://dev.mysql.com/doc/employee/en/). The script extracts data, performs analysis, and generates visualizations, which are then compiled into a PowerPoint presentation. The script can be set to run periodically (i.e. every two months) via a scheduling tool like Windows Task Scheduler, and can be adjusted to respond to a different database structure or to report alternative metrics relevant to the company. The tool fulfills Mandatory Disclosures S1-7 (Worforce Composition by Gender) and S1-16 (Gender Pay Gap) of the European Sustainability Reporting Standards.

In this instance, it recovers some basic aspects like current number of employees, average and total salary mass, distribution of employees and average salary across departments and titles, and evolution in the number of employees in total and across departments. It also incorporates some metrics relevant to the gender gap in the workforce, examining the differential evolution in the number of male and female employees as well as the gender differences in average starting salary and in salary progression across time (via a polynomial regression).

In the context of ESG reporting, monitoring the gender gap is essential to make sure that the company plans oriented to correct it are on tracks to meet targets. The topic lies at the intersection of SDGs 5 (Gender equality) and 8 (Fair Payment and Living Wage), which in the European Sustainability Reporting Standards are referenced by components S1 and S3.

## Author
Oscar García

## Purpose
This script automates the generation of HR reports by:

Extracting data from a MySQL database.
1. Extracting data from a MySQL database.
2. Analyzing and visualizing the data using Python libraries.
3. Compiling the visualizations and analyses into a PowerPoint presentation.

## Prerequisites

Before running the script, ensure you have the following:

1. **Python 3.12.4** installed on your system.
2. Necessary Python libraries:
    - pandas
    - numpy
    - matplotlib
    - seaborn
    - mysql-connector-python
    - statsmodels
    - scikit-learn
    - python-pptx

You can install these libraries using pip:

```sh
pip install pandas numpy matplotlib seaborn mysql-connector-python statsmodels scikit-learn python-pptx
```

3. Access to the MySQL 'Employees' database. You can download and set up the database from the following links:

## Employee Database Documentation
MySQL Documentation

## Configuration
Database Connection
Update the db_config dictionary in the script with your database connection details:

```sh
db_config = {
    'user': 'your_username',
    'password': 'your_password',
    'host': 'your_host',
    'database': 'employees',
    'charset': 'utf8mb4',
    'collation': 'utf8mb4_general_ci'
}
```

## Output Directory
Specify the directory where plots and slides will be saved:

```sh
output_dir = r'path_to_your_output_directory'
```

## Usage
To run the script, execute the following command in your terminal:

```sh
python HR_Reporting_Tool.py
```

## License
MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

## Contact
For any inquiries or support, please contact Oscar García at [oscaringc@hotmail.es].
