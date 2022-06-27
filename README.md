# Run At-Risk Assessment

## Project Description
An automation to run at-risk assessments for individual associates and entire cohorts from Trello checklist completion and QC metrics.

## Technologies Used
- UiPath Studio
- UiPath Orchestrator
- UiPath Assistant
- Microsoft Excel
- Microsoft Word
- Trello API

## Features
- Accepts data input through Excel workbooks.
- Calculates average QC score and checklist completion of Trello list for Project 1.
- Generates a PDF report for an associate containing average QC score and checklist completion if average QC score is at or below 2 (Yellow) or if checklist completion is below 75%.
- Generates a PDF report for a cohort containing the percentage of the cohort that is at-risk if more than 50% of the cohort is at-risk.
- Allows custom input and output directories to be set in config file.
- Sends emails containing reports to trainers and managers, or optionally to one or more override email addresses.

### Future Features
- Include option to allow reports to be grouped together, reducing number of emails sent.

## Getting Started

### Setting Up Orchestrator and Accounts
1. In Orchestrator, create a queue to contain the transaction items for the automation.
2. Create a Trello account that will have access to the workspace containing the Project 1 boards and generate an API key and token.
3. Add the API key and token as text assets in Orchestrator.
4. Add the login details of the email to send reports from as a credential asset in Orchestrator.

### Setting Up Automations
1. Clone project repository using `https://github.com/Revature-Capstone-1362/run-at-risk-assessment.git`.
2. Open `project.json` in both AtRiskAssessmentDispatcher and AtRiskAssessmentPerformer folders and publish the projects to Orchestrator.
3. Assign processes to robots.
4. Navigate to package directory on robot machines, usually located in `%userprofile%\.nuget\packages`.
5. Open `config.xlsx` in `1.x.x\lib\net45\Data` directory of dispatcher and set the queue name.
    - Optionally set the path to the directory for input data workbooks. The default directory for input data is the Input folder in this directory.
6. Open `config.xlsx` in `1.x.x\lib\net45\Data` directory of performer and set the queue name, the name of the Trello workspace containing the boards, information for the email account including the credential asset, server, and port, and information for the Trello assets in the Asset sheet.
    - Optionally set the paths to the directories for output report PDFs. The default directory for output reports is the Output folder in this directory.
    - Optionally set override email(s) that reports will be sent to. By default, automation will send emails to the trainer and manager.
    - Optionally modify document templates in the Input folder in this directory. Text enclosed in brackets must be present for automation to succeed.

### Constraints
- Input data workbooks must contain the following information with the specified column names: `Associate_Name, Trainer_Name, Manager_Name, GitHub_Organization, Core_Language, QC_Metrics`.
- The `QC_Metrics` must be formatted as a string of numbers corresponding to QC scores (1 - red, 2 - yellow, 3 - green, 4 - blue). (ex. 233 means the associate received a yellow, green, and green)

## Contributors
- Team Baymax
    - Warren Lee (Team Lead)
    - Cynnetha Bellinger
    - Frankline Elong
    - Jaye Ayala
- Team Iron Giant
    - Mike Tuskey (Team Lead)
    - Anthony McJunkins (Scrum Master)
    - Diahandra Christian
