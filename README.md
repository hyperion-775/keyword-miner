# keyword-miner
A program using Pandas which has the capability to use given data, filter through thousands of complaints and sort them into Sequences of Events (SOEs), using keywords, and output this filtration into an excel spreadsheet or output of choice.
# Prerequisites
You will need either python or an IDE for usage.
You will also need to install pandas with pip.
# Preparation
A few things need to be done to the data before the program is run. First, the SOE file must be created. In order to create the SOE file, a new excel spreadsheet must be made, in which there will be two columns, the SOE ID, and the SOE keywords. They will formatted as such: 
| BLANK     | BLANK                    |
| --------- | ------------------------ |
| SOE ID 1  | comma separated keywords |
| SOE ID 2  | comma separated keywords |
| SOE ID 3  | comma separated keywords |
| SOE ID 4  | comma separated keywords |
| SOE ID 5  | comma separated keywords |
| SOE ID 6  | comma separated keywords | and so forth.

## Keywords Formatting
- The keyword list MUST be comma seperated, and certain phrases must be phrased
properly. The three are “and”,”does not contain”, and “(in both ticket notes along with
conclusion summary columns)”.
-  Besides those phrases above, there should never be any quotation marks, or parenthesis.
# Editing for Use
Step 1 of this process is inserting the complaint and SOE files, insert the file in a path format inside the quotes, on lines 8, 10, and 12.
Step 2 of this process is inserting the columns letter in which the data resides on lines 14 through 29.
Step 3 is the insertion of the final output sheet in path format (IMPORTANT: This must be an empty, newly created excel spreadsheet) on line 32.
# Final Notes 
## Customization of Data Format
In order to change the format of data that the program uses, the python library Pandas can be used. It has documentation on many similar data types, such as CSV (If using CSV, use CSV UTF-8 format). Pandas documentation: https://pandas.pydata.org/docs/reference/io.html
