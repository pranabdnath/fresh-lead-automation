Fresh Lead Automation (Google Apps Script)
Problem Statement

Manually filtering CSV reports everyday leads to:

time wastage

duplicate buyers

wrong date-based results

missing fresh leads

difficulty tracking leads created after 6 PM

This script automates the entire process.



Solution Overview

This Google Apps Script automatically:
Fetches the latest email report
Extracts & cleans the CSV
Filters only valid entries (Condition Match 1)
Deduplicates buyers
Flags leads created after 6 PM
Updates a clean sheet called FreshLead



Logic Used

Get latest email → extract attachment

Clean dirty CSV values

Apply condition match (city + status + date range)

Remove duplicates using Set()

Flag “Today after 6PM” leads

Write fresh output to Google Sheet
