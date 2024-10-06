# CUI Online Data Scraper

This project automates the extraction of student information from the CUI Online portal. It uses **Selenium** for interacting with the web interface, bypasses CAPTCHA with **captcha_bypass**, and scrapes data from HTML pages with **BeautifulSoup**. Finally, it exports the data into structured Excel files using **Pandas**. The script handles profile data, course attendance, and results, offering a clean, organized Excel sheet as output.

## Key Features
- **CAPTCHA Bypass**: Automatically solves Google reCAPTCHA using the `captcha_bypass` module.
- **HTML Parsing**: Scrapes data like profile details, course attendance, and exam results using `BeautifulSoup`.
- **Excel Export**: Converts scraped data into structured Excel sheets for easy access, including custom formatting for better readability.
- **Image Capture**: Captures and stores profile images directly from the user profile page.

## Project Workflow
1. **Login Automation**: Uses Selenium to log in, solving CAPTCHA challenges.
2. **HTML Scraping**: Downloads profile, course, and result pages.
3. **Data Parsing**: Extracts structured information from HTML using regex and BeautifulSoup.
4. **Excel Generation**: Compiles the parsed data into multiple Excel sheets with clean formatting and image embedding.

## Challenges & Solutions
- **CAPTCHA Handling**: Google reCAPTCHA blocks scraping, but we overcome this using a third-party CAPTCHA solver.
- **Dynamic Page Elements**: The website uses iframes and dynamic elements, so the script includes explicit waits to ensure proper element detection.
- **Data Parsing Complexity**: Profile, courses, and results are structured differently, requiring custom parsing logic for each section.

## Requirements
- Python 3.x
- Selenium
- BeautifulSoup4
- Pandas
- captcha_bypass

## How to Use
1. Set up ChromeDriver and ensure it is available at `driver/chromedriver.exe`.
2. Install dependencies using `pip install -r requirements.txt`.
3. Provide your login credentials in the script and run `python scraper.py`.

### Disclaimer
This script is intended for educational purposes. Ensure you have permission to scrape data from the portal before use.
