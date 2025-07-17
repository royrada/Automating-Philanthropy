# Automating-Philanthropy
We build tools to semi-automate a certain type of philanthropy.  The journey is at its first steps, and this repository first contains Excel VBA code to score IRS Form 990s.  The scoring ranks the forms based on the extent to which the 501c3 that submitted the 990 has an endowment, awards scholarships, and focuses on science education or research. 
## Getting Started
To use this system:
Take the Excel file code.xlms and put it in a directory x (call it whatever you want).  
Download into x the txt files labeled nodenames, stopwords, punctuation, and rule.  
Download a zip file of IRS Form 990s from https://www.irs.gov/charities-non-profits/form-990-series-downloads and unzip it into a directory x/testforms.     
Create subdirectories in testforms called 990 and errant.
Now you are ready to start running the code.
### Prerequisites 
You need Excel or something that runs Excel VBA.  The plan is to create a Python version which uses a relational database backend and supports a web-based user interface.  For prototyping purposes for such a data intensive, simple, decision support tool, Excel can be convenient. 
### Installing
The system is meant to be modular and readily modified by users to suit their purposes.  
The nodenames file identifies the nodes from the IRS Form 990 files that the user wants to analyze.    The file is a semi-colon delimited line-by-line list of nodenames prefixed by datatype (string, date, or integer, or absint), number of characters in that data type, and the path to the node in the xml file.  Absint means absolute integer and takes a negative integer and converts it to a positive one.   Here are 4 illustrative lines from the nodename file currently available in this repository (which again the user is welcome to change):
Date;10;Return/ReturnHeader/TaxPeriodBeginDt
Integer;4;Return/ReturnHeader/TaxYr
AbsInt;15;Return/ReturnData/IRS990/CYInvestmentIncomeAmt
String;600;Return/ReturnData/IRS990/ActivityOrMissionDesc
The stopwords and punctuation files can be changed, and then the filtering of text descriptions from the Form 990s changes.  Most importantly, the rules which score the Forms are clearly defined and modifiable.  
The Parsed990Data worksheet will be populated with a header of nodenames in row 1 and form unique ids down column 1.  Inside the worksheet will be placed the values of the nodenames for each form.
The Scored990Data worksheet will have the same first column of form 990 ids but the column headers will be the rulename of each rule.  The cells within this worksheet will be populated with a 1 or 0 depending on whether the particular form satisfies the given rule.
There are 4 types of rules.  Each rule is a string of parameters delimited by semi-colons. The first parameter specifies the type of rule.  The number of parameters in a rule and the meaning of each parameter depends on the type.  In the templates the first parameter is always the type and the second the name of the rule to be used in the column of the worksheet that displays the score for each rule.  
Here are the 4 templates, each followed by a paragraph description.
Substring;Rulename;Nodename;Present; Tokens 
The third parameter specifies the Nodename from the Parsed990 worksheet that holds the data to be searched for substrings.   The fourth parameter is a T/F value which means whether the rule wants the nodename to contain or not contain a certain substring.  The final parameter is a comma-separated list of substrings to be tested for their presence in the nodename.
Trend;Rulename;comma-delimited list of Nodenames 
The third parameter specifies a list of Nodenames from the Parsed990 worksheet that hold the data to be trended.  The given order of the Nodenames is crucial.  The rule will count how many times the values go up or down and award a score of 1 to a trend that has more ups than downs.
Percentile;Rulename; Nodename; Cutoff 
The third parameter specifies the Nodename from the Parsed990 worksheet that holds the data.    Cutoff specifies in what percentile of values for all forms on this nodename must be for the rule to get assigned a 1.
Eval;Rulename; Nodename; NumOrTxt;Expression
Expression will be evaluated with value from Nodename and the data will be treated as text or numeric depending on the value of NumOrTxt.
Sample rules from rule.txt follow:
Eval;Age;IRS990_FormationYr;Num;Year(Now()) - IRS990_FormationYr > 15
Substring;Web;IRS990_WebsiteAddressTxt;T;academy,eduGive the example
Percentile;EndYrBal;CYEndwmtFundGrp_EndYearBalanceAmt;0.50
Trend;YrNet;IRS990_NetAssetsOrFundBalancesBOYAmt,IRS990_NetAssetsOrFundBalancesEOYAmt
## Running the tests
Go to the module VBA 'move990' and run the sub Move990Files which will move the 990 files to your 990 subdir and leave unmoved the other files, such as 990EZ.
Go to the module Parse and run the Subroutine ParseXML990Files.  This will parse the files in the 990 subdirectory and populate the Parsed990Data worksheet.
Next go to the module Strip and run the subroutine Master.   This sub will clean and strip that Desc file and moves the words unique, sorted woreds into a new column called DescFiltered.  It will then cleanse the web address in WebsiteAddressTxt.
Next go to the Score module and run the Sub Score.
## Authors and Contributing
Roy Rada conceived, designed, coded, and tested this system.  Microsoft Copilot provided invaluable support in the coding part.
## License
This project is licensed under GNU General Public License v3.0
