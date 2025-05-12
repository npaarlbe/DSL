# Welcome to the MorningStar Wiki!

This project automates the extraction of key metrics from Excel financial data using a custom domain-specific language (DSL). Below are screenshots and visualizations from parsing, regression, and error handling.

# Project Overview

MorningStar DSL enables financial analysts to correlate and validate projections across multiple company workbooks, especially when metrics vary in terminology. The system uses a DSL input file to define associations, specify independent/dependent Excel sources, and generate analysis through regression, filtering, sorting, and charting.

#  Features

* DSL grammar to describe associations, data files, filters, and visual outputs

* Automatically pulls metrics from Excel using keywords

* Produces combined or individual graphs (bar, line, pie, scatter)

* Supports regression (linear, exponential, logistic, etc.)

* Summary statistics (optional)

* Date filtering via ranges or lists

* Sorts data by alpha/numeric rules

* Graceful error handling and logging

* Graphs stored as interactive HTML and embedded PNG

#  Architecture

* ANTLR Grammar: Parses DSL input.

* Test.txt: Input file for ANTLR Grammer.

* Associations.txt: Input file that Test.txt uses for Associations. 

* ENTIREFILE (Java):Java Listener, that Captures DSL output and passes variables to the extractor.

* ExcelKeywordExtractor (Java): Parses Excel using Apache POI, matches keywords and years, and outputs a formatted file.

* POM (XML): Depencies files that uses maven to pull in all the dependcies for the project

* convertHTMLtoPNG (JavaScript): Uses Puppeteer to take screenshots of the html files to add them to the output excel file as PNG images.

* ECharts + Puppeteer: Generates interactive graphs and captures snapshots.

<img width="634" alt="Screenshot 2025-05-06 at 1 51 29 PM" src="https://github.com/user-attachments/assets/744e4f91-ea0e-4b33-a809-e237d819b486" />

1. User writes DSL input and association files
2. ANTLR parses DSL grammar and stores commands via listener
3. Java listener sends file paths, filters, and keywords to ExcelKeywordExtractor
4. Apache POI extracts matching rows from Excel files
5. Graphs generated via ECharts and converted to PNG with Puppeteer
6. Outputs saved into new Excel file with summary table and graphs


---
## Grammar of DSL

| Rule              | Description                                                                                           |
|-------------------|-------------------------------------------------------------------------------------------------------|
| `prog`            | Main entry point: expects an association block, independent/dependent file definitions, and stats.    |
| `ass`             | Declares the file path to the keyword associations.                                                   |
| `ind` / `dep`     | Defines independent and dependent Excel files (can include multiple `TICK` symbols with file paths).  |
| `date`            | Specifies either a list of years (with commas) or a year range using a dash.                          |
| `stat`            | Block that configures all analysis settings — includes date, regression, summary, listing, sort, etc. |
| `reg`             | Specifies regression type (linear, exponential, logistic, etc.).                                      |
| `sum`             | Toggles summary stats (`true/false`).                                                                 |
| `listall`         | Shows all data vs. just filtered (`true/false`).                                                      |
| `sort` / `sorttype` | Sorting logic — by year, alpha, high to low, etc.                                                   |
| `graphtype`       | Chart type to generate: bar, line, pie, scatter.                                                      |

```
grammar Expr;

prog        : ass ind dep stat;

ass         : ASS filePath;

filePath    : absolutePath | relativePath;

absolutePath: (WINDOWS_DRIVE | UNIX_ROOT) pathSegments?;

relativePath: pathSegments;

pathSegments: pathSegment ((UNIX_ROOT pathSegment)* | (WINDOWS pathSegment)*) |;
pathSegment : PATHSTRINGS | DOT | DOTDOT;

ind         : IND ind
            | TICK filePath ind
            |
            ;

dep         :
            | DEP dep
            | TICK filePath dep
            |
            ;

date        : DATE EQUAL num (comma num)*
            | DATE EQUAL num dash num;

dash        : DASH;
comma       : COMMA;

DASH        : ' ' '-' ' ';
COMMA       : ',';

num         : NUM;
DATE        : 'date';

stat        : date reg sum listall sort sorttype graphtype;

graphtype   : GRAPH EQUAL graphtypes;
GRAPH       : 'graph' | 'Graph';

graphtypes  : bar
            | line
            | pie
            | scatter;

bar         : BAR;
line        : LINE;
pie         : PIE;
scatter     : SCATTER;

BAR         : 'bar' | 'Bar';
LINE        : 'line' | 'Line';
PIE         : 'pie' | 'Pie';
SCATTER     : 'scatter' | 'Scatter';

reg         : REGRESSION EQUAL regtypes
            |
            ;

regtypes    : linear
            | exponential
            | log
            | multiple
            | logistic
            | ridge;

linear      : LINEAR;
exponential : EXPONENTIAL;
log         : LOG;
multiple    : MULTIPLE;
logistic    : LOGISTIC;
ridge       : RIDGE;
regression  : REGRESSION;

sum         : SUMMARY EQUAL TRUE
            | SUMMARY EQUAL FALSE
            |
            ;

listall     : LISTALL EQUAL TRUE
            | LISTALL EQUAL FALSE
            |
            ;

sort        : SORT EQUAL YEAR
            | SORT EQUAL
            |
            ;

sorttype    : SORTTYPE EQUAL HIGHTOLOW
            | SORTTYPE EQUAL ALPHA
            | SORTTYPE EQUAL LOWTOHIGH
            | SORTTYPE EQUAL BACKALPHA
            |
            ;

LINEAR        : 'linear' | 'Linear';
EXPONENTIAL   : 'exponential' | 'Exponential';
LOG           : 'log' | 'Log';
MULTIPLE      : 'multiple' | 'Multiple';
LOGISTIC      : 'logistic' | 'Logistic';
RIDGE         : 'ridge' | 'Ridge';

TRUE          : 'true' | 'True';
FALSE         : 'false' | 'False';

WINDOWS_DRIVE : [a-zA-Z] ':';
UNIX_ROOT     : '/';
WINDOWS       : '\\';
DOT           : '.';
DOTDOT        : '..';
SLASH         : '/' | '\\';

HIGHTOLOW     : 'HighToLow';
ALPHA         : 'Alpha';
LOWTOHIGH     : 'LowToHigh';
BACKALPHA     : 'BackAlpha';
YEAR          : 'year';

REGRESSION    : 'regression';
SORTTYPE      : 'sorttype';
SORT          : 'sort';
LISTALL       : 'listall';
SUMMARY       : 'summary';
EQUAL         : '=';
COLON         : ':';
ASS           : 'Association:';
IND           : 'Independent:';
DEP           : 'Dependent:';
TICK          : [A-Z]+;
NUM           : [1-9][0-9][0-9][0-9];
PATHSTRINGS   : [a-zA-Z0-9_.-]+;

WSPACE        : ' ' -> skip;
WS            : [ \t\r\n]+ -> skip;

```

<img width="703" alt="ParseTreeCenter" src="https://github.com/user-attachments/assets/6db5bb99-60b7-4fdc-a9f7-58268d7690c6" />
<img width="656" alt="ParseTreeLeft" src="https://github.com/user-attachments/assets/d2fb0b2f-78ec-4614-9d47-25e7822d3720" />
<img width="1176" alt="ParseTreeRight" src="https://github.com/user-attachments/assets/6901df81-712b-45ea-939d-72a92a2194bb" />
<img width="647" alt="FullParseTree" src="https://github.com/user-attachments/assets/2e522222-84dd-441e-8514-ed49b8d8c692" />


## Input File Example 

| File           | Role                                                                 |
|---------------------|----------------------------------------------------------------------|
| **DSL Input File**  | Directs the parser to analyze specific files with specific commands. |
| **Association File**| Maps search keywords to graph IDs for grouped visual output. 

### DSL Input File Example
```
Association:
    Association/Associations.txt

Independent:
    AAPL
    /ada/file.txt

Dependent:
    NEWTICK
    Excel/NAS_AMZN_CashFlowModel_20250207.xlsm
    AAPL
    Excel/NYS_VZ_CashFlowModel_20250219COPY.xlsm

date= 2022 - 2028
regression = Linear
summary = true

listall = False

sort = year

sorttype = BackAlpha

graph = scatter

```

### Association File Examples
If the current graph label (e.g., graph1) is the same as the previous one, the system adds the matching data to a combined graph. This behavior is optional and allows grouping multiple metrics into a single visualization when desired.

```
graph1 = est. iPhone main model units
graph2 = est. iPhone main model units growth y/y %
graph3 =  est. iPhone Pro revenue
graph4 = est. iPhone Pro sales growth y/y %
graph1 = est. iPhone Pro units
graph2 = est. iPhone Pro units growth y/y %
graph5 =  est. iPhone Pro ASP
graph4 = est. iPhone Pro ASP growth y/y %
graph3 =  est. iPhone main model revenue
graph4 = est. iPhone main model sales growth y/y %
graph1 = est. iPhone main model units
graph2 = est. iPhone main model units growth y/y %
graph5 =  est. iPhone main model ASP
graph6 = est. iPhone main model ASP growth y/y %
graph3 =  est. iPhone last year model revenue
graph4 = est. iPhone last year model sales growth y/y %
graph1 = est. iPhone last year model units
graph2 = est. iPhone last year model units growth y/y %
graph5 = est. iPhone last year model ASP
graph6 = est. iPhone last year model ASP growth y/y %
graph3 = est. iPhone SE & Older revenue
graph4 = est. iPhone SE/Older sales growth y/y %
graph1 = est. iPhone SE/Older units
graph2 = est. iPhone SE/Older units growth y/y %
graph5 = est. iPhone SE/Older ASP
graph6 = est. iPhone SE/Older ASP growth y/y %
```
Another Example of Associations

```
iPhone reported revenue
iPhone sales growth y/y %
iPhone sales growth q/q %
```


### Errors & Graceful Handling
![graceful2](https://github.com/user-attachments/assets/3c94b05c-ce3b-44dd-a7a7-a4138d4aea67)
![GracefuleError1](https://github.com/user-attachments/assets/f58f13c8-9722-4270-bde1-db97fb42b489)
<img width="1278" alt="fail" src="https://github.com/user-attachments/assets/b5353566-7f61-467b-9ecd-039f015c35d7" />
<img width="1278" alt="fail1" src="https://github.com/user-attachments/assets/96cb7345-e8c4-4200-8562-3a6e5eff0d50" />
<img width="1278" alt="fail2" src="https://github.com/user-attachments/assets/2f33a9b9-ee48-4092-aeff-dbb9ba719049" />

---
## meaningful and non-trivial application example

* The MorningStar DSL was tested on a non-trivial use case involving multi-year financial projections for Apple’s iPhone product lines, sourced from separate Excel files with varying metric names. The DSL allowed the user to:

* Parse multiple Excel models for tickers like AAPL, 
* Extract metrics with inconsistent labels (e.g., "est. iPhone Pro units" vs "iPhone sales growth y/y %")
* Group multiple metrics under the same chart using shared graphX labels
* Generate visual regression trends (e.g., linear and exponential) to forecast performance
* Create output files with both summary statistics and PNG chart exports
* This example replicates what an equity research analyst or investment firm would need to do when aggregating and comparing projections from different models or analysts.

---
## Validate the output of your DSL for the application example
* DSL input file that was parsed by ANTLR
```
Association:
    Association/Associations.txt


Dependent:
    AAPL
    Excel/NYS_VZ_CashFlowModel_20250219COPY.xlsm

date= 2022 - 2028
regression = Linear
summary = true

listall = False

sort = year

sorttype = BackAlpha

graph = scatter

```
* Association File Used
<img width="567" alt="ComplexAss1" src="https://github.com/user-attachments/assets/daa0d952-16cd-4d3d-a1c4-e2eeef5a789e" />

* Output excel file with the data from each association
<img width="1468" alt="AllExcelComplex1" src="https://github.com/user-attachments/assets/ab84e6f9-c0b5-4987-aca6-4a71bea70c92" />

* The combined graphs
<img width="620" alt="Complexg1" src="https://github.com/user-attachments/assets/16dd38c3-b910-4257-8bdf-79e01564e204" />
<img width="614" alt="Complexg2" src="https://github.com/user-attachments/assets/4b5af411-74dc-40fc-95ba-ecce4a40a4a5" />
<img width="615" alt="Complexg3" src="https://github.com/user-attachments/assets/2f14d0e2-c467-4e68-bc37-6ed3715cf767" />
<img width="612" alt="Complexg5" src="https://github.com/user-attachments/assets/877f4916-21d1-4330-9773-8e34dced49c3" />
<img width="612" alt="Complexg6" src="https://github.com/user-attachments/assets/7089189e-c36a-478f-9626-74bdaa85c99d" />

* The indidual graph for each association with linear regression as trendiness
<img width="620" alt="Complexg11" src="https://github.com/user-attachments/assets/2f59cebd-330f-407c-b413-038319b6fd1c" />

* Summary that was sorted by year

<img width="1450" alt="Screenshot 2025-05-08 at 11 47 23 AM" src="https://github.com/user-attachments/assets/2b6ca5ea-5996-4ab2-94c5-9c25ab25fbd7" />
<img width="1450" alt="Screenshot 2025-05-08 at 11 47 30 AM" src="https://github.com/user-attachments/assets/5f60c7d0-3066-4559-9faf-21c9ce7d1a6d" />
<img width="1450" alt="Screenshot 2025-05-08 at 11 47 35 AM" src="https://github.com/user-attachments/assets/e9b32a52-3370-4b29-8ef9-8204f526d4a7" />


---
## Discuss how the DSL compares to generating the same output with a general-purpose language or by hand

Using this DSL is way easier and faster than doing the same thing with a general-purpose language like Java or Python, or by hand in Excel. With the DSL, a financial user like an analyst or banker can just write a short input file that tells the system what data to pull, what years to include, and how to sort or graph it. That’s it. If you were doing this in Java, you’d have to write a full program that handles file input, filtering, sorting, matching keywords, and generating graphs — which would take a lot more time and technical knowledge. And doing it manually in Excel would mean opening every file, searching for the right keywords, checking column headers for the right years, copying things out, and building graphs one by one. When you're dealing with many big Excel models, that could take hours just to find and organize the right data. The DSL makes all of that automatic and scalable, so the user only has to focus on what they want to see, and not how to get it.

---
## Clearly indicate which group member is responsible for each portion of the project.
| Person           | Role                                                                 |
|---------------------|----------------------------------------------------------------------|
| **Nate Paarlberg**  | For this project, I mostly worked on the ExcelKeywordExtractor file. I focused on using Apache POI to go through a database of Excel files and pull out the specific associations and years we needed. I encorporate Echarts github, to create charts, and used pupiteer to add the charts to the output excel file. I focussed on using the variables and data from the input text file that was parsed, and using those to do everything we needed to do to find the correct information.  |
| **Caleb Hodel**| |
| **Together**| |
---
## Contributions
Created by Caleb Hodel and Nate Paarlberg | COS 382 Final Project 
Help from Jon Denning for overseeing this project and help along the way 
Built in Java using Apache POI and ANTLR  
Charts via ECharts → Headless Puppeteer snapshot

---
