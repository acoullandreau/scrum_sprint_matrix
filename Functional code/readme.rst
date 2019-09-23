==========================
Scrum sprint matrix script
==========================


---------------------
Purpose of the script
---------------------

This script intends to create or update the base of a sprint matrix.
The sprint matrix is divided in four sections:
* the issues already DONE
* the issues that are expected to be done by the end of the sprint
* the issues that have a low probability of being done by the end of the sprint
* the issues that were not sorted yet

The first time the matrix is created, it is possible to use as a source file the template files. Once the file has been updated with the issues sorted based on the prediction of the completion, this new file can be used as a source file, and will be updated by the script.


-------------------------------------------
Structure of the script and the input files
-------------------------------------------

The script is written in Python 3. Three libraries are used to parse the configuration, input and output files: csv, json and openpyxl (v. 2.6.0).

The configuration file (a json file) is used to "map" a username in JIRA with the name of the person we want to display in the output file.

There are two input files:
* the 'JIRA.csv' file, exported from JIRA (all the issues of the current sprint)
* the 'Sprint_matrix.xslx', that can either be the template (first creation) or a version already updated of the matrix

No arguments are used to launch the script.


-----------------------------
Structure of the output files
-----------------------------

There is one output file: 'Sprint_matrix_update.xslx' that contains a copy of the input file and the update from the csv file.

The sprint matrix contains only "standard issues", i.e no subtasks. The script maps the subtasks to their associated standard issue, and reports all the information of all issues and subtasks for the standard issue.

The script loops through the input 'Sprint_matrix.xslx' file, and if a key is found the line is updated. The position of the line in the file remains as-is.

If there are some keys in the 'JIRA.csv' file not present in the input matrix, new lines are added at the end of the file to be sorted.
The script is not able to "predict" the completion of a story, and will therefore always add the information about an unknown issue (i.e an issue that is not yet in the source file) in the end of the file (section To be sorted).
If an unknown issue (issue not in the source file) has a status DONE, it is automatically added to the Already done section.

The total sums for each section is computed by the script.

In case some issues were removed from the sprint in between two updates of the file, the script highlights the line of the removed story (the removal is not done automatically to allow tracking)


------------
Step by step
------------

1. Verify that the configuration file contains the list of possible assignees included in the sprint matrix
2. Download the sprint report file from JIRA and add it to the same folder from where the script is run (the file must be names 'JIRA.csv')
3. Add either a template (to be renamed 'Sprint_matrix.xslx') or add an existing file with the name 'Sprint_matrix.xslx'
4. Execute the script
5. Update the 'Sprint_matrix_update.xslx' file by sorting the additional lines, and reorganising the Done issues


----------------------
Initial files provided
----------------------

- readme.rst
- run.bat
- main.py
- conf.json
- Sprint_matrix_template.xlsx
