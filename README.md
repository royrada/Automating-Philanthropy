# Automated Philanthropy Scoring Framework
This project develops tools to semi-automate a specific style of philanthropic evaluation. The current system uses Excel VBA to score IRS Form 990s submitted by 501(c)(3) organizations. The scoring highlights entities that have endowments, award scholarships, and emphasize science education or research.
This framework is at an early stage of development, and future versions will be implemented in Python with relational database support and a web-based interface.
## 📦 Getting Started
To get the system running:
1. Download `Code.xlsm` and place it in your working directory (e.g., `x/`).
2. Download the `.txt` files:
   - `nodenames.txt`
   - `stopwords.txt`
   - `punctuation.txt`
   - `rule.txt`
   Place them in the same directory as the Excel file.
3. Download IRS Form 990 XML files from [IRS Form 990 Series Downloads](https://www.irs.gov/charities-non-profits/form-990-series-downloads).
4. Unzip those forms into `x/testforms/`, and create subdirectories:
   - `x/testforms/990` for standard 990 files
   - `x/testforms/errant` for nonstandard or filtered-out files

## 🛠 Prerequisites

- Microsoft Excel with VBA enabled (or compatible environment)
- Basic familiarity with running macros and editing file paths

> ⚠️ This Excel-based prototype will be superseded by a Python version with improved scalability and accessibility.
## 🧩 Installation Notes
- The system is modular and can be customized to suit different evaluation criteria.
- You may edit any of the supporting `.txt` files to change what data is parsed or scored.
### 🗂 nodenames.txt
Each line defines:
- Data type (`String`, `Date`, `Integer`, `AbsInt`)
- Field length
- XML path to the node
Example:
```
Date;10;Return/ReturnHeader/TaxPeriodBeginDt  
Integer;4;Return/ReturnHeader/TaxYr  
AbsInt;15;Return/ReturnData/IRS990/CYInvestmentIncomeAmt  
String;600;Return/ReturnData/IRS990/ActivityOrMissionDesc  
```
### 🧹 stopwords.txt and punctuation.txt
Used to clean and tokenize text fields—feel free to modify.
### 📜 rule.txt
Defines scoring logic for each rule. Users can modify or add rules.
## 📊 Parsed & Scored Worksheets
- `Parsed990Data` contains extracted data:
  - Headers: nodenames
  - Rows: form unique IDs and their values
- `Scored990Data` contains rule evaluations:
  - Headers: rule names
  - Rows: binary scores (1 or 0)
## 🧠 Rule Types
There are four rule types. Each uses a semicolon-delimited format:
### 1. `Substring`
```
Substring;RuleName;Nodename;Present;token1,token2,...
```
- Checks if tokens are present (or absent) in the specified text node.
### 2. `Trend`
```
Trend;RuleName;Nodename1,Nodename2,...
```
- Compares values across nodes for an upward/downward trend.

### 3. `Percentile`
```
Percentile;RuleName;Nodename;Cutoff
```
- Scores 1 if a value is above the given percentile cutoff.

### 4. `Eval`
```
Eval;RuleName;Nodename;NumOrTxt;Expression
```
- Evaluates logical expressions involving the node's value.

### 📌 Sample Rules from `rule.txt`:
```
Eval;Age;IRS990_FormationYr;Num;Year(Now()) - IRS990_FormationYr > 15  
Substring;Web;IRS990_WebsiteAddressTxt;T;academy,edu  
Percentile;EndYrBal;CYEndwmtFundGrp_EndYearBalanceAmt;0.50  
Trend;YrNet;IRS990_NetAssetsOrFundBalancesBOYAmt,IRS990_NetAssetsOrFundBalancesEOYAmt  
```
## ▶️ Running the System
1. **Move Files**  
   In VBA module `move990`, run `Move990Files`  
   > Moves Form 990 files to `/990/`, skips Form 990EZ and others

2. **Parse XMLs**  
   In module `Parse`, run `ParseXML990Files`  
   > Extracts nodename data into `Parsed990Data`

3. **Clean Text**  
   In module `Strip`, run `Master`  
   > Cleans descriptions and web addresses; populates `DescFiltered`

4. **Score Data**  
   In module `Score`, run `Score`  
   > Evaluates rules and outputs to `Scored990Data`

## 👤 Authors & Contributions

- **Roy Rada**: Project lead, architecture, Excel VBA development, testing  
- **Microsoft Copilot**: Collaborative assistance in coding and system design

## 📄 License

This project is licensed under the **GNU General Public License v3.0**.
