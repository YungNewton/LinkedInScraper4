import os
import re
import time
import random
import pickle
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.hyperlink import Hyperlink
import undetected_chromedriver as uc
import chromedriver_autoinstaller

class LinkedInProfileScraper:
    def __init__(self, output_file, include_columns, connection_range=(0, 10), excel_file_path=None):
        self.driver = self.init_driver()
        self.output_file = output_file
        self.urls = []
        self.cookies_file = "cookies.pkl"
        self.connection_range = connection_range
        self.include_columns = include_columns  # Save INCLUDE_COLUMNS as an instance attribute
        self.excel_file_path = excel_file_path


    def init_driver(self):
        chromedriver_autoinstaller.install()  # Ensures ChromeDriver matches your Chrome version
        options = uc.ChromeOptions()
        options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")

        driver = uc.Chrome(options=options, version_main=132, use_subprocess=True)
        return driver

    def save_html_content(self, company_name):
        try:
            # Create folder if it doesn't exist
            folder_path = os.path.join(os.getcwd(), "html_files")
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)

            # Save HTML content to file
            html_content = self.driver.page_source
            file_path = os.path.join(folder_path, f"{company_name}.html")
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(html_content)
            print(f"HTML content saved for {company_name} at {file_path}")
        except Exception as e:
            print(f"Error saving HTML content: {e}")

    def load_urls_from_excel(self):
        try:
            if not self.excel_file_path:
                return []  # Return an empty list if no file path is provided
            
            df = pd.read_excel(self.excel_file_path)
            if 'URL' not in df.columns:
                raise ValueError("Input Excel file must contain a 'URL' column.")
            return df['URL'].dropna().tolist()  # Return a list of non-empty URLs
        except Exception as e:
            print(f"Error loading URLs from Excel: {e}")
            return []

    def save_cookies(self):
        try:
            cookies = self.driver.get_cookies()
            with open(self.cookies_file, "wb") as f:
                pickle.dump(cookies, f)
            print("Cookies saved.")
        except Exception as e:
            print(f"Error saving cookies: {e}")

    def load_cookies(self):
        try:
            if os.path.exists(self.cookies_file):
                with open(self.cookies_file, "rb") as f:
                    cookies = pickle.load(f)
                    for cookie in cookies:
                        self.driver.add_cookie(cookie)
                print("Cookies loaded.")
            else:
                print("Cookies file not found.")
        except Exception as e:
            print(f"Error loading cookies: {e}")

    def manual_login(self):
        self.driver.get("https://www.linkedin.com/login")
        input("Log in manually and press Enter when done.")
        self.save_cookies()

    def login(self):
        self.driver.get("https://www.linkedin.com")
        self.load_cookies()
        time.sleep(2)
        self.driver.refresh()
        if not self.is_session_valid():
            print("Session invalid. Please log in manually.")
            self.manual_login()

    def is_session_valid(self):
        try:
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'img.global-nav__me-photo'))
            )
            return True
        except:
            return False

    def random_pause(self):
        pause_duration = random.uniform(1, 2)
        time.sleep(pause_duration)
    
    def human_scroll(self):
        """Perform a small human-like scroll to trigger HTML and avoid bot detection."""
        try:
            scroll_pause = random.uniform(1, 1.5)  # Small pause to simulate human-like behavior
            scroll_times = random.randint(2, 4)  # Scroll only 2-4 times to trigger content loading
            
            for _ in range(scroll_times):
                scroll_step = random.randint(200, 500)  # Small scroll step
                self.driver.execute_script(f"window.scrollBy(0, {scroll_step});")
                time.sleep(scroll_pause)  # Pause after each scroll step to avoid detection
            
        except Exception as e:
            print(f"Error during human scroll: {e}")

    def scroll_to_end(self):
        """Scroll to the bottom of the page until content stops loading."""
        try:
            scroll_pause = random.uniform(1, 1.5)  # Pause between scrolls to simulate human behavior
            last_height = self.driver.execute_script("return document.body.scrollHeight")  # Initial page height
            
            while True:
                # Scroll down by a small step
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(scroll_pause)
                
                # Wait for new content to load
                new_height = self.driver.execute_script("return document.body.scrollHeight")
                
                # Check if the scroll height has not increased
                if new_height == last_height:
                    print("No more content to load.")
                    break
                
                last_height = new_height  # Update last height
                
        except Exception as e:
            print(f"Error during scroll to end: {e}")
    
    def normalize_url(self, url):
        """Remove the trailing slash from a URL if present."""
        return url.rstrip("/")

    def get_excel_connection_urls(self):
        # Load URLs from the Excel file
        excel_urls = set(map(self.normalize_url, self.load_urls_from_excel()))  # Use a set for faster lookups
        if not excel_urls:
            print("No URLs found in the Excel file.")
            return []

        # Navigate to the LinkedIn sent invitations page
        self.driver.get("https://www.linkedin.com/mynetwork/invitation-manager/sent/")
        time.sleep(2)  # Allow the page to load

        pending_profiles = set()  # Store unique URLs globally
        pending_connections = []  # Store scraped details
        last_profile_url = None  # Track the last profile's URL
        retry_count = 0  # Retry counter

        while len(pending_connections) < len(excel_urls):
            self.human_scroll()
            time.sleep(1)

            # Retrieve all currently loaded pending connection items
            current_profiles = self.driver.find_elements(By.CSS_SELECTOR, 'li.invitation-card')

            # Add URLs from current profiles to the cumulative set
            for profile in current_profiles:
                try:
                    url = profile.find_element(By.CSS_SELECTOR, 'a[href*="linkedin.com/in/"]').get_attribute('href')
                    pending_profiles.add(url)
                except Exception:
                    continue

            # Scrape details for profiles in the Excel file
            for url in pending_profiles:
                # Skip URLs not in Excel or already processed
                if url not in excel_urls or any(conn["profile_url"] == url for conn in pending_connections):
                    continue

                try:
                    # Extract profile details
                    message = "N/A"
                    sent_time = "N/A"

                    try:
                        see_more_button = profile.find_element(By.CSS_SELECTOR, 'a.lt-line-clamp__more')
                        self.driver.execute_script("arguments[0].click();", see_more_button)
                        time.sleep(0.5)  # Allow time for the content to expand
                    except Exception:
                        pass  # If no button is found, proceed to scrape the visible text

                    message_element = profile.find_elements(By.CSS_SELECTOR, '.invitation-card__custom-message span.lt-line-clamp__line')
                    if message_element:
                        message = message_element[0].text.strip() if message_element[0].text.strip() else "N/A"

                    sent_time_element = profile.find_elements(By.CSS_SELECTOR, '.time-badge.t-12.t-black--light.t-normal')
                    if sent_time_element:
                        sent_time = sent_time_element[0].text.strip() if sent_time_element[0].text.strip() else "N/A"

                    pending_connections.append({"profile_url": url, "message": message, "sent_time": sent_time})

                except Exception as e:
                    print(f"Error processing profile {url}: {e}")
                    continue

            # Check if all Excel URLs are processed
            if all(url in {conn["profile_url"] for conn in pending_connections} for url in excel_urls):
                print("All Excel URLs processed. Exiting loop.")
                break

            # Check for pagination if no new profiles are loaded
            try:
                current_last_url = current_profiles[-1].find_element(By.CSS_SELECTOR, 'a[href*="linkedin.com/in/"]').get_attribute('href')
            except Exception:
                current_last_url = None

            if current_last_url == last_profile_url:
                retry_count += 1
                print(f"No new profiles loaded. Retry {retry_count}/2.")
                time.sleep(3)
                if retry_count >= 2:
                    print("Attempting to navigate to the next page.")
                    try:
                        next_button = self.driver.find_element(By.CSS_SELECTOR, 'button.artdeco-pagination__button--next')
                        if not next_button.get_attribute("disabled"):  # Check if the "Next" button is enabled
                            self.driver.execute_script("arguments[0].click();", next_button)
                            time.sleep(2)  # Allow time for the next page to load
                            retry_count = 0  # Reset retry count
                            continue  # Retry loading profiles on the new page
                        else:
                            print("No more pages to navigate. Exiting.")
                            break
                    except Exception as e:
                        print(f"Failed to navigate to the next page: {e}")
                        break
            else:
                retry_count = 0  # Reset retry count if new content is found

            # Update the last profile URL
            last_profile_url = current_last_url

        return pending_connections

    def get_unanswered_connection_urls(self, connection_range):
        self.driver.get("https://www.linkedin.com/mynetwork/invitation-manager/sent/")

        # Scrape the number of sent invitations for People
        try:
            people_invites = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button[aria-label*='sent People invitation'] span.artdeco-pill__text"))
            ).text
            match = re.search(r"\((\d{1,3}(?:,\d{3})*)\)", people_invites)
            if match:
                total_pending_invites = int(match.group(1).replace(",", ""))
                print(f"Total Pending Invites: {total_pending_invites}")
            else:
                raise ValueError("No match found in 'people_invites'")

            # Adjust the target count based on range and available invites
            start_count = connection_range[0] - 1
            end_count = min(connection_range[1], total_pending_invites)
        except Exception as e:
            print(f"Failed to retrieve pending invites count: {e}")
            start_count = connection_range[0] - 1
            end_count = connection_range[1]

        pending_profiles = []  # Store URLs in a list to preserve order
        seen_urls = set()  # Track unique URLs to avoid duplicates
        pending_connections = []  # Store scraped details
        last_profile_url = None  # Track the last profile's URL
        retry_count = 0  # Retry counter

        while len(pending_profiles) < end_count:
            self.human_scroll()
            time.sleep(1)

            # Retrieve all currently loaded pending connection items
            current_profiles = self.driver.find_elements(By.CSS_SELECTOR, 'li.invitation-card')

            # Add URLs from current profiles to the list, maintaining order
            for profile in current_profiles:
                try:
                    url = profile.find_element(By.CSS_SELECTOR, 'a[href*="linkedin.com/in/"]').get_attribute('href')
                    if url not in seen_urls:
                        seen_urls.add(url)
                        pending_profiles.append(url)
                except Exception:
                    continue

            # Scrape details for profiles within the specified range
            for url in pending_profiles[start_count:end_count]:
                # Skip URLs already processed
                if any(conn["profile_url"] == url for conn in pending_connections):
                    continue

                try:
                    # Extract profile details
                    message = "N/A"
                    sent_time = "N/A"

                    try:
                        see_more_button = profile.find_element(By.CSS_SELECTOR, 'a.lt-line-clamp__more')
                        self.driver.execute_script("arguments[0].click();", see_more_button)
                        time.sleep(0.5)  # Allow time for the content to expand
                    except Exception:
                        pass  # If no button is found, proceed to scrape the visible text

                    message_element = profile.find_elements(By.CSS_SELECTOR, '.invitation-card__custom-message span.lt-line-clamp__line')
                    if message_element:
                        message = message_element[0].text.strip() if message_element[0].text.strip() else "N/A"

                    sent_time_element = profile.find_elements(By.CSS_SELECTOR, '.time-badge.t-12.t-black--light.t-normal')
                    if sent_time_element:
                        sent_time = sent_time_element[0].text.strip() if sent_time_element[0].text.strip() else "N/A"

                    pending_connections.append({"profile_url": url, "message": message, "sent_time": sent_time})

                except Exception as e:
                    print(f"Error processing profile {url}: {e}")
                    continue

            # Check for pagination if no new profiles are loaded
            try:
                current_last_url = current_profiles[-1].find_element(By.CSS_SELECTOR, 'a[href*="linkedin.com/in/"]').get_attribute('href')
            except Exception:
                current_last_url = None

            if current_last_url == last_profile_url:
                retry_count += 1
                print(f"No new profiles loaded. Retry {retry_count}/2.")
                time.sleep(3)
                if retry_count >= 2:
                    print("Attempting to navigate to the next page.")
                    try:
                        next_button = self.driver.find_element(By.CSS_SELECTOR, 'button.artdeco-pagination__button--next')
                        if not next_button.get_attribute("disabled"):  # Check if the "Next" button is enabled
                            self.driver.execute_script("arguments[0].click();", next_button)
                            time.sleep(2)  # Allow time for the next page to load
                            retry_count = 0  # Reset retry count
                            continue  # Retry loading profiles on the new page
                        else:
                            print("No more pages to navigate. Exiting.")
                            break
                    except Exception as e:
                        print(f"Failed to navigate to the next page: {e}")
                        break
            else:
                retry_count = 0  # Reset retry count if new content is found

            # Update the last profile URL
            last_profile_url = current_last_url

        return pending_connections


    def click_see_more_button(self):
        """Use JavaScript to click the 'See more' button in the 'About' section if available."""
        try:
            # Locate the button with CSS selector
            see_more_button = self.driver.find_element(By.CSS_SELECTOR, 'button.inline-show-more-text__button')

            # Execute JavaScript to click the button
            self.driver.execute_script("arguments[0].click();", see_more_button)
            
            # Pause to allow the content to expand
            self.random_pause()  
            print("'See more' button clicked via JavaScript.")
        except Exception as e:
            print(f"'See more' button not found or couldn't be clicked: {e}")
        
    def scrape_experience(self, profile_url):
        # Navigate to the profile's experience details page
        experience_url = f"{profile_url}/details/experience/"
        self.driver.get(experience_url)
        self.random_pause()
        self.human_scroll()
        self.scroll_to_end()

        current_positions = {"Position Title": [], "Position Description": [], "Company Name": []}
        more_positions = []
        more_descriptions = []
        more_skills = []
        processed_anchors = set()  # Track processed anchors to avoid duplication
        experiences = []  # Store experiences for calculating total experience
        current_firm_experiences = []  # NEW: Stores (start_date, end_date) pairs for current roles


        try:
            # Wait for all experience list items to load
            experience_items = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'li.pvs-list__paged-list-item'))
            )

            for experience in experience_items:
                try:
                    # Check for company with multiple roles
                    company_name_elements = experience.find_elements(By.CSS_SELECTOR, '.t-bold span[aria-hidden="true"]')
                    if len(company_name_elements) > 1:
                        # Find all company name elements (this includes both the company name and roles)
                        anchors = experience.find_elements(By.CSS_SELECTOR, 'a.optional-action-target-wrapper.display-flex.flex-column.full-width')

                        # Company name
                        company_name = company_name_elements[0].text.strip()
                        role_dates = []  # Collect all start/end dates for this experience item
                        has_current_role = False  # Track if any role is ongoing (no end date)

                        # Process multiple roles under the same company
                        if len(company_name_elements) > 1:
                            # Loop through each anchor and extract role, start date, and end date
                            for idx, anchor in enumerate(anchors[1:], start=1):
                                # Skip processing if this anchor has already been handled
                                if anchor in processed_anchors:
                                    continue

                                # Add anchor to processed set
                                processed_anchors.add(anchor)

                                try:
                                    role_element = anchor.find_element(By.CSS_SELECTOR, '.mr1.hoverable-link-text.t-bold span[aria-hidden="true"]')
                                    job_title = role_element.text.strip()
                                except:
                                    job_title = "N/A"

                                try:
                                    date_range_element = anchor.find_element(By.CSS_SELECTOR, '.pvs-entity__caption-wrapper[aria-hidden="true"]')
                                    date_range = date_range_element.text.strip()
                                    start_date, end_date, _ = self.extract_dates_and_duration(date_range)
                                except:
                                    start_date, end_date = "N/A", "N/A"

                                role_dates.append((start_date, end_date))

                                # After extracting start_date and end_date
                                experiences.append({
                                    "start_date": start_date,  # Use your existing logic for start_date
                                    "end_date": end_date       # Use your existing logic for end_date
                                })

                                # Fetch position description if available
                                try:
                                    description_spans = anchor.find_elements(By.CSS_SELECTOR, '.t-14.t-normal.t-black span[aria-hidden="true"]')
                                    position_description = " ".join([
                                        span.text.strip() for span in description_spans
                                        if span.text.strip() and not span.find_elements(By.TAG_NAME, "strong") and not span.find_elements(By.TAG_NAME, "svg")
                                    ])
                                    position_description = f"{position_description}" if position_description else "N/A"
                                except:
                                    position_description = "N/A"

                                # Extract skills if available using a simplified selector
                                try:
                                    skills_elements = experience.find_elements(By.CSS_SELECTOR, 'span[aria-hidden="true"]')
                                    skills_text = "N/A"
                                    for skill in skills_elements:
                                        if "Skills:" in skill.text:
                                            skills_text = skill.text.replace("Skills:", "").strip()
                                            break
                                    skills_text = f"{skills_text}" if skills_text != "N/A" else "N/A"
                                except Exception as e:
                                    skills_text = "N/A"

                                # Format position entry
                                position_entry = (
                                    f"Position: {job_title} - Company: {company_name} - StartDate: {start_date} - EndDate: {end_date}"
                                )

                                more_descriptions.append(position_description)
                                more_skills.append(skills_text)

                                # Add positions to current or more positions list based on end date
                                if end_date == " ":
                                    current_positions["Position Title"].append(job_title)
                                    current_positions["Position Description"].append(position_description)
                                    current_positions["Company Name"].append(company_name)
                                    more_positions.insert(0, position_entry)
                                    has_current_role = True  

                                else:
                                    more_positions.append(position_entry)

                            if has_current_role:
                                current_firm_experiences.append(role_dates)

                    else:
                        # Fetch job title
                        job_title_element = experience.find_element(By.CLASS_NAME, 'mr1')
                        job_title = job_title_element.find_element(By.CSS_SELECTOR, 'span[aria-hidden="true"]').text.strip()

                        # Skip if the job title is N/A
                        if job_title == "N/A":
                            continue

                        # Fetch company name, avoiding visually hidden elements
                        company_name = None
                        try:
                            # Select the company name that is not inside a visually hidden span
                            company_name_element = experience.find_element(By.CSS_SELECTOR, '.t-14.t-normal span[aria-hidden="true"]')
                            company_name = company_name_element.text.strip().split('·')[0].strip()
                        except:
                            company_name = "N/A"

                        # Check if this anchor has already been processed
                        anchor_element = experience.find_element(By.CSS_SELECTOR, 'a.optional-action-target-wrapper')
                        if anchor_element in processed_anchors:
                            continue

                        # Mark this anchor as processed
                        processed_anchors.add(anchor_element)

                        # Fetch date range and extract start and end dates
                        date_range = experience.find_element(By.CLASS_NAME, 'pvs-entity__caption-wrapper').text.strip()
                        start_date, end_date, _ = self.extract_dates_and_duration(date_range)

                        # After extracting start_date and end_date
                        experiences.append({
                            "start_date": start_date,  # Use your existing logic for start_date
                            "end_date": end_date       # Use your existing logic for end_date
                        })

                        # Fetch the full job description only from span elements without <strong> or <img> tags
                        try:
                            description_spans = experience.find_elements(By.CSS_SELECTOR, '.t-14.t-normal.t-black span[aria-hidden="true"]')
                            position_description = " ".join([
                                span.text.strip() for span in description_spans
                                if span.text.strip() and not span.find_elements(By.TAG_NAME, "strong") and not span.find_elements(By.TAG_NAME, "svg")
                            ])
                            position_description = f"{position_description}" if position_description else "N/A"
                        except:
                            position_description = "N/A"

                        ## Extract skills if available using a simplified selector
                        try:
                            skills_elements = experience.find_elements(By.CSS_SELECTOR, 'span[aria-hidden="true"]')
                            skills_text = "N/A"
                            for skill in skills_elements:
                                if "Skills:" in skill.text:
                                    skills_text = skill.text.replace("Skills:", "").strip()
                                    break
                            skills_text = f"{skills_text}" if skills_text != "N/A" else "N/A"
                        except Exception as e:
                            skills_text = "N/A"

                        # Format position entry
                        position_entry = (
                            f"Position: {job_title} - Company: {company_name} - StartDate: {start_date} - EndDate: {end_date}"
                        )

                        more_descriptions.append(position_description)
                        more_skills.append(skills_text)

                        # Add positions to current or more positions list based on end date
                        if end_date == " ":
                            current_positions["Position Title"].append(job_title)
                            current_positions["Position Description"].append(position_description)
                            current_positions["Company Name"].append(company_name)
                            more_positions.insert(0, position_entry)
                            current_firm_experiences.append((start_date, end_date))

                        else:
                            more_positions.append(position_entry)

                except Exception as e:
                    print(f"Error scraping experience item: {e}")
                    continue

            # Convert current positions into comma-separated strings for each column
            current_positions = {k: ", ".join(v) for k, v in current_positions.items()}
            more_positions_string = ',\n'.join(more_positions)
            more_descriptions_string = ',\n'.join(more_descriptions)
            more_skills_string = ',\n'.join(more_skills)

            return current_positions, more_positions_string, more_descriptions_string, more_skills_string, experiences, current_firm_experiences

        except Exception as e:
            return {"Position Title": "N/A", "Company Name": "N/A"}, "N/A", "N/A", "N/A"

    def extract_dates_and_duration(self, date_range):
        """
        Extract start and end dates from the date range string.
        Sets the end date as a space if 'Present' is indicated.
        """
        try:
            if " · " in date_range:
                date_part, _ = date_range.split(" · ")
            else:
                date_part = date_range

            if " - " in date_part:
                start_date_str, end_date_str = date_part.split(" - ")

                # Parse start date
                start_date = self.parse_date(start_date_str.strip())

                # Use a single space if "Present" in end_date_str, otherwise parse date
                end_date = " " if "Present" in end_date_str else self.parse_date(end_date_str.strip())

                # Format dates in mm/yyyy format if they are valid datetime objects
                start_date_formatted = start_date.strftime("%m/%Y") if start_date else "N/A"
                end_date_formatted = " " if end_date == " " else (end_date.strftime("%m/%Y") if end_date else "N/A")
                return start_date_formatted, end_date_formatted, ""
            else:
                single_date = self.parse_date(date_part.strip())
                single_date_formatted = single_date.strftime("%m/%Y") if single_date else "N/A"
                return single_date_formatted, " ", ""
        except Exception as e:
            print(f"Error extracting dates and duration: {e}")
            return "N/A", " ", ""
    
    def parse_date(self, date_str):
        try:
            return datetime.strptime(date_str, "%b %Y")  # Parses dates like "Jan 2014"
        except ValueError:
            try:
                return datetime.strptime(date_str, "%Y")  # Parses dates like "2014"
            except ValueError:
                print(f"Unrecognized date format for: {date_str}")
                return None
    
    def calculate_total_experience(self, experiences):
        """
        Calculate the total experience based on a list of experiences.
        Keeps 'Present' as a string in end dates, using the current date when needed.
        """
        try:
            date_pairs = []

            for experience in experiences:
                start_date_str = experience.get("start_date", "N/A")
                end_date_str = experience.get("end_date", " ")
                
                # Convert start_date_str to datetime object
                if start_date_str != "N/A":
                    start_date = datetime.strptime(start_date_str, "%m/%Y")
                else:
                    continue  # Skip if no valid start date is found

                # Convert end_date_str to datetime object or use current date if "Present"
                end_date = datetime.now() if end_date_str == " " else datetime.strptime(end_date_str, "%m/%Y")
                date_pairs.append((start_date, end_date))

            # If there are no valid date pairs, return "0.0"
            if not date_pairs:
                return "0.0"

            # Find the minimum start date and the maximum end date
            earliest_start_date = min([start for start, _ in date_pairs])
            latest_end_date = max([end for _, end in date_pairs])

            # Calculate the difference in years and months
            total_months = (latest_end_date.year - earliest_start_date.year) * 12 + (latest_end_date.month - earliest_start_date.month)
            years = total_months // 12
            months = total_months % 12

            # Format the result as 'years.months' with zero-padded months if less than 10 (e.g., 0.10 for 10 months)
            if years == 0:
                total_experience = f"0.{months}"  # Only show months as decimal for less than 1 year
            else:
                total_experience = f"{years}.{str(months).zfill(2)}"

            return total_experience
        except Exception as e:
            print(f"Error calculating total experience: {e}")
            return "0.0"
        
    def scrape_education(self, profile_url):
        # Navigate to the profile's education details page
        education_url = f"{profile_url}/details/education/"
        self.driver.get(education_url)
        self.random_pause()
        self.human_scroll()
        self.scroll_to_end()

        education_degree = "N/A"
        school_name = "N/A"
        more_educations = []

        try:
            # Wait for all education list items to load
            education_items = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'li.pvs-list__paged-list-item'))
            )

            for idx, education in enumerate(education_items):
                try:
                    # Extract school name
                    school_element = education.find_element(By.CSS_SELECTOR, '.t-bold span[aria-hidden="true"]')
                    school_name_text = school_element.text.strip()

                    # Try to extract the degree name
                    try:
                        degree_element = education.find_element(By.CSS_SELECTOR, '.t-14.t-normal span[aria-hidden="true"]')
                        degree_text = degree_element.text.strip()
                        # Check if the degree text contains only numbers (or date-like text)
                        if degree_text.replace(" ", "").isdigit():
                            degree_text = "N/A"
                    except:
                        degree_text = "N/A"

                    # Extract date range for education (e.g., "2008 - 2012")
                    try:
                        date_element = education.find_element(By.CSS_SELECTOR, '.pvs-entity__caption-wrapper[aria-hidden="true"]')
                        date_text = date_element.text.strip()
                        start_date, end_date, _ = self.extract_dates_and_duration(date_text)
                    except:
                        start_date, end_date = "N/A", "N/A"

                    education_entry = (
                        f"Degree: {degree_text} - School Name: {school_name_text} - StartDate: {start_date} - EndDate: {end_date}"
                    )

                    # If it's the first education item, set it for the main columns
                    if idx == 0:
                        education_degree = degree_text
                        school_name = school_name_text
                        # Insert current education at the top of more_educations
                        more_educations.insert(0, education_entry)
                    else:
                        # Format other education details into the "More Educations" column
                        more_educations.append(
                            f"Degree: {degree_text} - School Name: {school_name_text} - StartDate: {start_date} - EndDate: {end_date}"
                        )
                except Exception as e:
                    print(f"Error scraping education item: {e}")
                    continue

            # Join all additional education entries into a single string
            more_educations_string = ',\n'.join(more_educations)
            return education_degree, school_name, more_educations_string

        except Exception as e:
            print(f"Error locating education items.")
            return "N/A", "N/A", "N/A"
        
    def scrape_contact_info(self, profile_url):
        contact_info_url = f"{profile_url}/overlay/contact-info/"
        self.driver.get(contact_info_url)
        self.random_pause()
        
        contact_info = {
            "PhoneNumber": "N/A",
            "Email Address": "N/A"
        }
        birthday = "N/A"
        connected_on = "N/A"

        try:
            # Wait for the contact info modal to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.artdeco-modal__content'))
            )

            # Wait for the loader to disappear, if it exists
            WebDriverWait(self.driver, 10).until_not(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.artdeco-loader'))
            )

            # Use JavaScript to extract the "Connected On" and "Birthday" information
            contact_data = self.driver.execute_script('''
                let contactModal = document.querySelector('div.artdeco-modal__content');
                let data = { "email": "N/A", "phone": "N/A", "birthday": "N/A", "connectedOn": "N/A" };
                
                contactModal.querySelectorAll('section.pv-contact-info__contact-type').forEach(section => {
                    let header = section.querySelector('h3') ? section.querySelector('h3').innerText : '';
                    if (header.includes('Email')) {
                        let emailElem = section.querySelector("a[href^='mailto:']");
                        if (emailElem) data["email"] = emailElem.href.replace('mailto:', '');
                    }
                    else if (header.includes('Phone')) {
                        let phoneElem = section.querySelector('span.t-14.t-black.t-normal');
                        if (phoneElem) data["phone"] = phoneElem.innerText.trim();
                    }
                    else if (header.includes('Birthday')) {
                        let birthdayElem = section.querySelector('span.t-14.t-black.t-normal');
                        if (birthdayElem) data["birthday"] = birthdayElem.innerText.trim();
                    }
                    else if (header.includes('Connected')) {
                        let connectedOnElem = section.querySelector('span.t-14.t-black.t-normal');
                        if (connectedOnElem) data["connectedOn"] = connectedOnElem.innerText.trim();
                    }
                });
                return data;
            ''')

            contact_info["Email Address"] = contact_data.get("email", "N/A")
            contact_info["PhoneNumber"] = contact_data.get("phone", "N/A")
            birthday_raw = contact_data.get("birthday", "N/A")
            connected_on_raw = contact_data.get("connectedOn", "N/A")
            birthday = birthday_raw

            # # Convert "Birthday" to MM-DD format (Month and Day only)
            if birthday_raw != "N/A":
                try:
                    birthday_date = datetime.strptime(birthday_raw, "%B %d")
                    birthday = birthday_date.strftime("%d-%b")  # Format as "15-Nov"
                except ValueError:
                    birthday = birthday_raw  # Retain original if format doesn't match

            # Convert "Connected On" date to YYYY-MM-DD format
            if connected_on_raw != "N/A":
                try:
                    connected_on_date = datetime.strptime(connected_on_raw, "%b %d, %Y")
                    connected_on = connected_on_date.strftime("%d/%m/%Y")
                except ValueError:
                    connected_on = connected_on_raw  # Retain original if format doesn't match

            # Format and return
            formatted_contact_info = ", ".join([f"{key}: {value}" for key, value in contact_info.items() if value != "N/A"])
            return formatted_contact_info if formatted_contact_info else "N/A", birthday, connected_on

        except Exception as e:
            print(f"Failed to scrape contact info: {e}")
            return "N/A", "N/A", "N/A"

    def scrape_interests(self, profile_url):
        relevant_interests = ['Groups', 'Newsletters', 'Companies', 'Top Voices', 'Schools']
        scraped_interests = {interest: [] for interest in relevant_interests}

        try:
            # Load the profile URL and navigate to the interests section (index 0)
            interest_url = f"{profile_url}/details/interests/?detailScreenTabIndex=0"
            self.driver.get(interest_url)
            time.sleep(2)  # Allow page to load
            self.human_scroll()

            # Extract the buttons to determine their index based on the interest names
            interest_buttons = self.driver.find_elements(By.CSS_SELECTOR, 'div.artdeco-tablist button.artdeco-tab')

            # Map the interest names to their indices
            interest_map = {}
            for index, button in enumerate(interest_buttons):
                try:
                    tab_name = button.find_element(By.CSS_SELECTOR, 'span[aria-hidden="true"]').text.strip()
                    if tab_name in relevant_interests:
                        interest_map[tab_name] = index
                except Exception as e:
                    continue

            print(f"Detected Interests: {interest_map}")

            # Now navigate to each relevant interest tab using its index and scrape items
            for interest_name, tab_index in interest_map.items():
                interest_url = f"{profile_url}/details/interests/?detailScreenTabIndex={tab_index}"
                self.driver.get(interest_url)
                time.sleep(2)
                self.human_scroll()

                # Extract all the interest items (name and URL)
                interest_items = self.driver.find_elements(By.CSS_SELECTOR, 'li.pvs-list__paged-list-item')
                for item in interest_items:
                    try:
                        # Scrape the interest name
                        interest_name_element = item.find_elements(By.CSS_SELECTOR, 'div.hoverable-link-text.t-bold span[aria-hidden="true"]')
                        if interest_name_element:
                            interest_name_text = interest_name_element[0].text.strip()
                        else:
                            interest_name_element = item.find_elements(By.CSS_SELECTOR, 'span.visually-hidden')
                            interest_name_text = interest_name_element[0].text.strip() if interest_name_element else ""

                        # Only proceed if the name is not empty
                        if interest_name_text:
                            # Scrape the interest URL from the second anchor if available
                            interest_url = "N/A"
                            interest_url_elements = item.find_elements(By.CSS_SELECTOR, 'a.optional-action-target-wrapper')
                            if len(interest_url_elements) > 1:
                                interest_url = interest_url_elements[1].get_attribute('href')
                            elif interest_url_elements:
                                interest_url = interest_url_elements[0].get_attribute('href')

                            # Format and store the interest
                            formatted_interest = f"{interest_name}: {interest_name_text} - URL: {interest_url}"
                            scraped_interests[interest_name].append(formatted_interest)

                    except Exception as e:
                        print(f"Error scraping interest item: {e}")
                        continue

            # Format each interest list into a string with each item on a new line
            formatted_interests = {key: '\n'.join(value) for key, value in scraped_interests.items()}
            return formatted_interests

        except Exception as e:
            print(f"Error navigating to interests: {e}")
            return {interest: [] for interest in relevant_interests}

    def scrape_profiles_for_you(self, profile_url):
        # Navigate to the "Profiles for You" section of the profile
        profiles_for_you_url = f"{profile_url}/overlay/browsemap-recommendations/"
        self.driver.get(profiles_for_you_url)
        self.random_pause()
        self.human_scroll()

        profiles_data = []

        try:
            # Wait for the list of profile links to load
            profile_items = WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'li.artdeco-list__item'))
            )

            for item in profile_items:
                try:
                    # Extract the name element
                    name_element = item.find_element(By.CSS_SELECTOR, 'div.hoverable-link-text.t-bold span[aria-hidden="true"]')
                    name = name_element.text.strip() if name_element else None

                    # Extract the profile URL element
                    url_element = item.find_element(By.CSS_SELECTOR, 'a.optional-action-target-wrapper')
                    profile_link = url_element.get_attribute('href').split('?')[0] if url_element else None

                    # Extract the description element
                    try:
                        description_element = item.find_element(By.CSS_SELECTOR, 'div.t-14.t-normal.display-flex.align-items-center span[aria-hidden="true"]')
                        description = description_element.text.strip() if description_element else "N/A"
                    except:
                        description = "N/A"

                    # Only proceed if both name and URL are available
                    if name and profile_link:
                        # Append the formatted string to the list
                        profiles_data.append(f"Name: {name}, URL: {profile_link}, Description: {description}")

                except Exception as e:
                    continue

            # Join all profiles into a single string with each on a new line
            return "\n".join(profiles_data)

        except Exception as e:
            print(f"Error scraping 'Profiles for You': {e}")
            return "N/A"

    def scrape_profile(self, url):
        self.driver.get(url)
        self.random_pause()
        self.human_scroll()

        result = {
                    "flagshipProfileUrl": url
                }

        # Define column-method mappings
        METHOD_COLUMN_MAP = {
            "scrape_name": {"fullName"},
            "scrape_summary": {"summary"},
            "scrape_headline": {"headline"},
            "scrape_location": {"location"},
            "scrape_connections": {"numOfConnections"},
            "scrape_degree": {"Degree"},
            "scrape_contact_info": {"ContactInfo", "Birthday", "ConnectedOn"},
            "scrape_experience": {
                "Position Title", "Position Description", "Company Name",
                "More Positions", "Descriptions", "Skills"
            },
            "scrape_education": {"Education Degree", "SchoolName", "More Educations"},
            "scrape_total_experience": {"Total Years of Exp(in Yrs)"},
            "scrape_current_firm_experience": {"Exp in Current Firm(In Yrs.Months)"},
            "scrape_interests": {
                "Interest: Groups", "Interest: Newsletters",
                "Interest: Companies", "Interest: Top Voices",
                "Interest: Schools"
            },
            "scrape_profiles_for_you": {"Profiles for You"},
            "scrape_connection_status": {"Connection Status"},  # New Column

        }

        if METHOD_COLUMN_MAP["scrape_name"].intersection(self.include_columns):
            try:
                # Scrape full name from the h1 tag
                full_name = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.TAG_NAME, 'h1'))
                ).text
            except Exception as e:
                full_name = "N/A"
                print(f"Failed to scrape name: {e}")
            result["fullName"] = full_name
            

        if METHOD_COLUMN_MAP["scrape_summary"].intersection(self.include_columns):
            try:
                # Locate all sections with the potential "About" heading
                sections = self.driver.find_elements(By.CSS_SELECTOR, 'section.artdeco-card')

                summary = "N/A"  # Default value

                for section in sections:
                    try:
                        # Check if the section has an "About" heading
                        heading_element = section.find_element(By.CSS_SELECTOR, 'h2.pvs-header__title span[aria-hidden="true"]')
                        heading_text = heading_element.text.strip()

                        if heading_text == "About":
                            # Locate the summary within the correct "About" section
                            summary_element = section.find_element(By.CSS_SELECTOR, 'div.display-flex.ph5.pv3 span[aria-hidden="true"]')
                            summary = summary_element.text.strip()

                            # Check for unwanted phrases
                            if "You've previously worked with" in summary or "You've previously worked together" in summary:
                                summary = "N/A"  # Set to N/A if it contains unwanted phrases
                            
                            # Break after finding the first valid "About" section
                            break
                    except Exception:
                        # Continue to the next section if the current one doesn't match or fails
                        continue

            except Exception as e:
                summary = "N/A"
                print(f"Failed to scrape summary: {e}")

            # Add the scraped or default summary to the result
            result["summary"] = summary

        if METHOD_COLUMN_MAP["scrape_headline"].intersection(self.include_columns):
            try:
                # Scrape headline from the div with class text-body-medium
                headline = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div.text-body-medium'))
                ).text
            except Exception as e:
                headline = "N/A"
                print(f"Failed to scrape headline: {e}")
            result["headline"] = headline

        # Check for "Connection Status"
        if METHOD_COLUMN_MAP["scrape_connection_status"].intersection(self.include_columns):
            try:
                # Locate the svg icon first, then find its parent button
                clock_svg = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'svg[data-test-icon="clock-small"]'))
                )
                pending_button = clock_svg.find_element(By.XPATH, './ancestor::button')
                
                # Verify the button contains the "Pending" text
                connection_status = pending_button.find_element(By.CSS_SELECTOR, 'span.artdeco-button__text').text.strip()
                if "Pending" not in connection_status:
                    connection_status = "-"
            except Exception as e:
                connection_status = "-"
                print(f"Error retrieving connection status: {e}")

            result["Connection Status"] = connection_status
        
        if METHOD_COLUMN_MAP["scrape_location"].intersection(self.include_columns):
            try:
                # Scraping the location
                location = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'span.text-body-small.inline.t-black--light.break-words'))
                ).text
            except Exception as e:
                location = "N/A"
                print(f"Failed to scrape location: {e}")
            result["location"] = location
            
        
        if METHOD_COLUMN_MAP["scrape_connections"].intersection(self.include_columns):
            try:
                # Scrape number of followers or connections
                num_of_connections = self.driver.find_element(By.CSS_SELECTOR, 'p.text-body-small').text

                # Use regex to extract only the first numeric value and remove commas
                match = re.search(r'\d{1,3}(?:,\d{3})*', num_of_connections)  # Extracts the first occurrence like 1,600
                if match:
                    # Convert the number to an integer and remove commas
                    num_of_connections = int(match.group(0).replace(',', ''))
                else:
                    num_of_connections = "N/A"

            except Exception as e:
                num_of_connections = "N/A"
                print(f"Failed to scrape followers: {e}")
            result["numOfConnections"] = num_of_connections

        if METHOD_COLUMN_MAP["scrape_degree"].intersection(self.include_columns):
            try:
                # Scrape degree information
                degree_element = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'span.dist-value'))
                )
                degree = degree_element.text.strip()  # Get the degree text (e.g., '3rd')
            except Exception as e:
                degree = "N/A"
                print(f"Failed to scrape degree: {e}")
            result["Degree"] = degree

        if METHOD_COLUMN_MAP["scrape_contact_info"].intersection(self.include_columns):        
            try:
                contact_info, birthday, connected_on = self.scrape_contact_info(url)
            except:
                contact_info = "N/A"
                birthday = "N/A"
                connected_on = "N/A"
            result["ContactInfo"] = contact_info
            result["Birthday"] = birthday
            result["ConnectedOn"] = connected_on

        if METHOD_COLUMN_MAP["scrape_experience"].intersection(self.include_columns):
            # Call scrape_experience with profile URL
            current_positions, more_positions, more_descriptions, more_skills, experiences, current_firm_experiences = self.scrape_experience(url)

            try:
                total_years_of_exp = self.calculate_total_experience(experiences)
                result["Total Years of Exp(in Yrs)"] = total_years_of_exp
            except:
                total_years_of_exp = "N/A"
                result["Total Years of Exp(in Yrs)"] = "N/A"

            try:
                exp_in_current_firm = self.calculate_current_firm_experience(current_firm_experiences)
                result["Exp in Current Firm(In Yrs.Months)"] = exp_in_current_firm
            except:
                exp_in_current_firm = "N/A"
                result["Exp in Current Firm(In Yrs.Months)"] = "N/A"

            result.update({
                "Position Title": current_positions.get("Position Title", "N/A"),
                "Position Description": current_positions.get("Position Description", "N/A"),
                "Company Name": current_positions.get("Company Name", "N/A"),
                "More Positions": more_positions,
                "Descriptions": more_descriptions,
                "Skills": more_skills,
            })
            

        if METHOD_COLUMN_MAP["scrape_education"].intersection(self.include_columns):
            # Call scrape_education with profile URL
            education_degree, school_name, more_educations = self.scrape_education(url)
            result.update({
                "Education Degree": education_degree,
                "SchoolName": school_name,
                "More Educations": more_educations,
            })

        if METHOD_COLUMN_MAP["scrape_interests"].intersection(self.include_columns):
            # Scrape interests and map to the correct columns
            interest_data = self.scrape_interests(url)
            result.update({
                "Interest: Groups": interest_data.get("Groups", "N/A"),
                "Interest: Newsletters": interest_data.get("Newsletters", "N/A"),
                "Interest: Companies": interest_data.get("Companies", "N/A"),
                "Interest: Top Voices": interest_data.get("Top Voices", "N/A"),
                "Interest: Schools": interest_data.get("Schools", "N/A"),
            })

        if METHOD_COLUMN_MAP["scrape_profiles_for_you"].intersection(self.include_columns):
            try:
                # Call scrape_profiles_for_you with the current profile URL
                profiles_for_you_data = self.scrape_profiles_for_you(url)
                result["Profiles for You"] = profiles_for_you_data
            except Exception as e:
                result["Profiles for You"] = "N/A"
                print(f"Failed to scrape 'Profiles for You' section: {e}")

        return {k: v for k, v in result.items() if k in self.include_columns}

    def calculate_current_firm_experience(self, current_firm_experiences):
        """
        Calculate the total experience for current firm experiences.
        Uses the current date for ongoing roles where end_date is empty.
        """
        try:
            date_pairs = []

            # Process each list of (start_date, end_date) pairs in current_firm_experiences
            for experience_group in current_firm_experiences:
                for start_date_str, end_date_str in experience_group:
                    
                    # Convert start_date_str to datetime object
                    if start_date_str != "N/A":
                        start_date = datetime.strptime(start_date_str, "%m/%Y")
                    else:
                        continue  # Skip if no valid start date is found

                    # Convert end_date_str to datetime object or use current date if "Present"
                    end_date = datetime.now() if end_date_str == " " else datetime.strptime(end_date_str, "%m/%Y")
                    date_pairs.append((start_date, end_date))

            # If there are no valid date pairs, return "0.0"
            if not date_pairs:
                return "0.0"

            # Find the minimum start date and the maximum end date
            earliest_start_date = min([start for start, _ in date_pairs])
            latest_end_date = max([end for _, end in date_pairs])

            # Calculate the difference in years and months
            total_months = (latest_end_date.year - earliest_start_date.year) * 12 + (latest_end_date.month - earliest_start_date.month)
            years = total_months // 12
            months = total_months % 12

            # Format the result as 'years.months' with zero-padded months if less than 10 (e.g., 2.06 for 2 years, 6 months)
            if years == 0:
                total_experience = f"0.{months}"  # Only show months as decimal for less than 1 year
            else:
                total_experience = f"{years}.{str(months).zfill(2)}"

            return total_experience
        except Exception as e:
            print(f"Error calculating current firm experience: {e}")
            return "0.0"
    
    def generate_custom_title(self, full_name, summary):
        """
        Generate a custom title using the full name and summary.
        Ensures the title does not exceed 50 characters.
        """
        try:
            if not full_name and not summary:
                return "No Title Available"

            # Combine name and summary
            combined = f"{full_name} - {summary}" if full_name and summary else (full_name or summary)

            # Truncate to 50 characters and add "..." if too long
            return combined[:47] + "..." if len(combined) > 50 else combined
        except Exception as e:
            print(f"Error generating title: {e}")
            return "Error Generating Title"

    def save_to_excel(self, data):
        # Generate a unique timestamped filename
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
        output_file = f"{os.path.splitext(self.output_file)[0]}_{timestamp}.xlsx"

        try:
            # Create a new workbook
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "LinkedIn Data"

            # Write the header row
            ordered_columns = [col for col in self.include_columns if col in data[0].keys()]
            sheet.append(ordered_columns)

            # Write data rows
            for row in data:
                row_data = []
                for col in ordered_columns:
                    if col == "flagshipProfileUrl":
                        url = row.get("flagshipProfileUrl", "")
                        title = self.generate_custom_title(row.get("fullName", ""), row.get("summary", ""))
                        if title != "No Title Available":
                            # Add hyperlink with the generated title
                            # cell = f'=HYPERLINK("{url}", "{title}")'
                            cell = f'=HYPERLINK("{url}", "{url}")'
                        else:
                            # Use the URL as-is
                            cell = f'=HYPERLINK("{url}", "{url}")'
                        row_data.append(cell)
                    else:
                        # Add normal data for other columns
                        row_data.append(row.get(col, "N/A"))
                sheet.append(row_data)

            # Save the workbook
            workbook.save(output_file)
            print(f"Data saved to {output_file}")

            # If a previous file exists, delete it
            if hasattr(self, 'previous_output_file') and os.path.exists(self.previous_output_file):
                os.remove(self.previous_output_file)
                print(f"Old file {self.previous_output_file} deleted.")

            # Update the reference to the latest file
            self.previous_output_file = output_file

        except Exception as e:
            print(f"Error saving to {output_file}: {e}")

    def run(self):
        self.login()
        profiles_data = []

        # Load URLs from Excel if the file path is provided
        if self.excel_file_path:
            print("Retrieving data for URLs from Excel file...")
            pending_connections = self.get_excel_connection_urls()
        else:
            print("Retrieving URLs via scraping...")
            pending_connections = self.get_unanswered_connection_urls(self.connection_range)

        processed_urls = set()  # Track URLs to ensure no duplicates

        for connection in pending_connections:
            if not isinstance(connection, dict):
                print(f"Skipping malformed connection entry: {connection}")
                continue  # Skip if not a valid dictionary
            
            url = connection.get("profile_url", "").strip()

            # Ensure `url` is extracted correctly
            if isinstance(url, dict):  
                url = url.get("profile_url", "").strip()

            if not isinstance(url, str) or not url:
                print(f"Skipping invalid URL: {url}")
                continue  # Skip invalid or empty URLs
            
            # 🚀 **Fix duplicate URL issue**
            if url in processed_urls:
                print(f"Skipping duplicate URL: {url}")
                continue
            processed_urls.add(url)  # Add to processed URLs set

            message = connection.get("message", "N/A")
            sent_time = connection.get("sent_time", "N/A")

            print(f"Scraping profile: {url}...")  # ✅ Ensure each URL is different

            try:
                # Ensure **each profile is actually scraped fresh**
                self.driver.get(url)  
                time.sleep(2)  # Give time for page to load before scraping

                profile_data = self.scrape_profile(url)

                # ✅ **Ensure data is tied to the specific profile**
                profile_data["profile_url"] = url
                profile_data["message"] = message
                profile_data["sent time"] = sent_time

                profiles_data.append(profile_data)

                # Save progress (avoids losing data if script crashes)
                self.save_to_excel(profiles_data)
            except Exception as e:
                print(f"Error scraping {url}: {e}")
                continue

        # Final cleanup
        self.driver.quit()

if __name__ == "__main__":
    # Columns to include in the output
    INCLUDE_COLUMNS = [
        "fullName",
        "Search Query",
        "summary",
        "headline",
        # "location",
        # "flagshipProfileUrl",
        # "numOfConnections",
        # "Degree",
        # "Position Title",
        # "Position Description",
        # "Company Name",
        # "More Positions",
        # "Descriptions",
        # "Skills",
        # "Education Degree",
        # "SchoolName",
        # "More Educations",
        # "Total Years of Exp(in Yrs)",
        # "Exp in Current Firm(In Yrs.Months)",
        # "ContactInfo",
        # "Interest: Groups",
        # "Interest: Newsletters",
        # "Interest: Companies",
        # "Interest: Top Voices",
        # "Interest: Schools",
        # "Birthday",
        # "ConnectedOn",
        # "Profiles for You",
        # "Connection Status",
        "message",
        "sent time",
    ]
    output_file = "linkedin_output.xlsx"  # Output file
    connection_range = (92, 94)  # Specify the range of connections to scrape
    excel_file_path = "linkedin_profiles.xlsx"  # Replace with actual Excel file path or set to None
    scraper = LinkedInProfileScraper(
        output_file,
        include_columns=INCLUDE_COLUMNS,
        connection_range=connection_range,
        # excel_file_path=excel_file_path  # Pass the Excel file path here
    )

    scraper.run()