# SaveLife-Foundation-Internship

## About the Project
This project was developed during my internship at SaveLIFE Foundation, where the primary goal was to support the creation of automated PowerPoint presentations and action plans. These deliverables were intended to help recommend targeted interventions to agencies like NHAI, PWD, and Urban/Local Bodies for reducing road crashes.

Since the organization was working with large sets of spreadsheets and presentations, they needed a solution to automate the linking and updating of data across these files.

Due to the nature of the data and presentations, the actual files cannot be shared publicly. However, the core logic and implementation of the automation scripts are included here.

## What the Scripts Do
I was able to automate two major sections of the workflow:

Table Creation from Data (table_with_template.gs)  
This script fetches real-time data from the spreadsheet and automatically inserts it into a pre-defined table format within the presentation. It was used for generating data-driven slides directly from the source files.

Key Recommendations (Key_recommendation.gs)  
This script generates spreadsheets based on crash types and counts on specific roads under specific agencies. It also links these spreadsheets to the updated numbers in the presentation and action plans.

Multiple Presentation Copies (Apps Script + Python)  
To support large-scale creation of customized presentation decks, I wrote scripts using both Google Apps Script and Python (via Google Drive API in Google Colab). These scripts create multiple copies of a base slide deck and save them in a designated Google Drive folder.

This was useful when generating personalized versions of presentations for different regions, agencies, or use cases.

## Work in Progress

I had also started working on automating the Summary Recommendation section (Summary_Recommendation.gs), but this part remained incomplete as my internship concluded before it could be finalized.

## Note
These scripts are specifically tailored to the structure and requirements of the SaveLIFE Foundation's internal presentations, spreadsheets, and action plans. They may not be directly reusable without modifications, but they can serve as a reference for similar use cases involving Google Apps Script automation.
