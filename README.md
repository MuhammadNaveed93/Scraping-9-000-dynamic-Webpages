# Scraping-9-000-dynamic-Webpages
Project Requirement:
To extract 9,000 data records (Company records) from a website

Project Overview
The project required extracting data/records of 9,000+ companies, seperate webpage for each company. The main challenge was working with JavaScript-rendered pages, including the main page and the 9,000 additional webpages from which the data was needed to be scraped.

Project Solution:
- Utilized Selenium for scraping data from 9,000 dynamic webpages
- Divided the project into two phases:
- First: Extract links for all 9,000 webpages and stored the links in the excel file (refer to file "Scraping all links.py"). 
- Second: Using pandas, read the links (stored in excel file) and accessed each webpage to scrape desired information of that specific company - (Refer to "Scraping data from each links.py").
- Developed a solution that allowed the scraping process to be resumed from the exact point where it was interrupted. This ensured that even in the case of internet issues or code errors, the program could pick up where it left off and continue extracting data without starting from scratch.

A separate function file, where all necessary required functions(codes) related to Project were written. This included prpgrams like, creating excel file, updating existing excel file, excel file formatting (merging of cells) etc. 

Finally, Exported the scraped data to Excel file with optimal formatting.
