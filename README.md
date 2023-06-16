# Scraping-9-000-dynamic-Webpages
Project Requirement:
To extract 9,000 data records (Company records) from a website

Project Overview
The project required extracting 9,000 data records from a dynamic webpage. The main challenge was working with JavaScript-rendered pages, including the main page and the 9,000 additional webpages from which the data was needed to be scraped.

Project Solution:
- Utilized Selenium for scraping data from 9,000 dynamic webpages
- Divided the project into two phases:
- First: to extract link for all 9,000 webpages
- Second: to access each page and scrape desired data.
- Developed a solution that allowed the scraping process to be resumed from the exact point where it was interrupted. This ensured that even in the case of internet issues or code errors, the program could pick up where it left off and continue extracting data without starting from scratch.
- Exported the final scraped data to Excel with optimal formatting.
