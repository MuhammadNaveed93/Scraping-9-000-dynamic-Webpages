# Scraping-9-000-dynamic-Webpages
Project Requirement:
To extract 9,000 data records (Company records) from a website

Project Overview
The project required extracting data/records of 9,000+ companies, seperate webpage for each company. The main challenge was working with JavaScript-rendered pages, including the main page and the 9,000 additional webpages from which the data was needed to be scraped.

Project Solution:
- Utilized Selenium for scraping data from 9,000 dynamic webpages
- Divided the project into two phases:
- First: Extract links for all 9,000 webpages - "Scraping all links.py" (I only ran program for 02 districts where there were appx 2000+ comapnies)
- Second: to access each page and scrape desired data - 2. Scraping data from each links.
- Developed a solution that allowed the scraping process to be resumed from the exact point where it was interrupted. This ensured that even in the case of internet issues or code errors, the program could pick up where it left off and continue extracting data without starting from scratch.
- Exported the final scraped data to Excel with optimal formatting.
- A separate function file was created for the Project where all required functions, including program for formatting of excel file, were developed. 
