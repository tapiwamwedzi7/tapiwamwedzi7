import requests
from bs4 import BeautifulSoup
import pandas as pd

def scrape_jobs():
    url = "https://vacancymail.co.zw/jobs/"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"}
    
    # Send the HTTP request to get the page content
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Find all job listings
    job_listings = soup.find_all('a', class_='job-listing')
    
    # Data lists
    job_titles = []
    companies = []
    descriptions = []
    locations = []
    expiry_dates = []
    
    for job in job_listings:
        try:
            job_title = job.find('h3', class_='job-listing-title').get_text(strip=True)
        except AttributeError:
            job_title = 'No job title found'

        try:
            company = job.find('h4', class_='job-listing-company').get_text(strip=True)
        except AttributeError:
            company = 'No company found'

        try:
            description = job.find('p', class_='job-listing-text').get_text(strip=True)
        except AttributeError:
            description = 'No description found'
        
        try:
            location = job.find('i', class_='icon-material-outline-location-on').find_next('li').get_text(strip=True)
        except AttributeError:
            location = 'No location found'
        
        try:
            expiry_date = job.find('i', class_='icon-material-outline-access-time').find_next('li').get_text(strip=True)
        except AttributeError:
            expiry_date = 'No expiry date found'

        # Append to lists
        job_titles.append(job_title)
        companies.append(company)
        descriptions.append(description)
        locations.append(location)
        expiry_dates.append(expiry_date)
    
    # Create a DataFrame from the lists
    df = pd.DataFrame({
        'Job Title': job_titles,
        'Company': companies,
        'Description': descriptions,
        'Location': locations,
        'Expiry Date': expiry_dates
    })
    
    # Save to CSV
    df.to_csv('scraped_jobs.csv', index=False)
    print("Job data scraped successfully and saved to 'scraped_jobs.csv'")

# Call the scrape_jobs function to execute
scrape_jobs()