from classfile import JiraToolBox
import json

with open('conf.json', encoding='utf-8') as config_file:
    conf_data = json.load(config_file)

sprintmatrix = JiraToolBox(conf_data, 'matrix')
sprintmatrix.parseJiraIssues('JIRA.csv')
sprintmatrix.processJiraIssues()
sprintmatrix.parseExistingFile('Sprint_matrix.xlsx')
sprintmatrix.processIssueOwnership()
new_issues = sprintmatrix.processNewIssues()
sprintmatrix.addNewIssues(new_issues)
sprintmatrix.updateExistingIssues()
sprintmatrix.updateTotalsPerPredictionSection()
sprintmatrix.highlightRemovedIssues()
sprintmatrix.writeUpdatedFile('Sprint_matrix_updated.xlsx')