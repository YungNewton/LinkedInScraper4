    def get_follower_urls(self, follower_range):
        self.driver.get("https://www.linkedin.com/mynetwork/network-manager/people-follow/followers/")

        try:
            # Extract the number of followers
            followers_header = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "p.t-black--light.t-14.mn-network-manager__subtitle"))
            ).text
            total_followers = int(re.search(r"(\d+(?:,\d+)*)", followers_header).group(1).replace(",", ""))
            print(f"Total Followers: {total_followers}")

            # Adjust the target count based on available followers
            target_count = min(follower_range[1], total_followers)
        except Exception as e:
            print(f"Failed to retrieve follower count: {e}")
            target_count = follower_range[1]

        # Scroll to load followers and retrieve their profile URLs
        total_visible_profiles = 0
        while total_visible_profiles/2 < target_count:          
            self.human_scroll()
            time.sleep(1)
            followers = self.driver.find_elements(By.CSS_SELECTOR, 'div.linked-area.flex-1.cursor-pointer a[href*="/in/"]')
            total_visible_profiles = len(followers)
            if total_visible_profiles/2 >= target_count:
                break

        # Collect URLs with a step of 2 to skip duplicates
        follower_urls = [follower.get_attribute('href') for i, follower in enumerate(followers) if i % 2 == 0]

        return follower_urls[follower_range[0] - 1:target_count]