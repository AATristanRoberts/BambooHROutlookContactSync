import os
import time
import json
import asyncio
import requests
import tempfile
import win32com.client as win32
from pyppeteer import launch


CHROME_PATH = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
DIRECTORY_URL = "https://automationanywhere.bamboohr.com/employees/directory.php"

FILTER_CARDS = "AA AU"  # A filter to apply to the contacts

async def main():
    tmp_img_path = os.path.join(tempfile.gettempdir(), "AAContactSyncTemp.jpg")
    if not os.path.exists(CHROME_PATH):
        print("Could not find Chrome at", CHROME_PATH)
        exit()

    # Open Chrome
    browser = await launch(
        headless=False,  # Make it visible to user
        executablePath=CHROME_PATH
    )

    # Navigate to page and wait for user login
    page = await browser.newPage()
    await page.goto(DIRECTORY_URL, {
        "waitUntil": "networkidle0"
    })

    if page.url != DIRECTORY_URL:
        print("Waiting for login...")
        await asyncio.wait([page.waitForNavigation()])

    # Wait for the elements to load (very slow)
    await asyncio.wait([page.waitForSelector(".EmployeeCardContainer")])

    print("Page loaded")

    # Filter the cards
    if FILTER_CARDS and len(FILTER_CARDS) > 0:
        code = """
            var __valid = false;
            document.querySelectorAll("li.FilterListItem a").forEach(function(el) {
                if (el.innerText.indexOf("FILTER_CARDS") > -1) {
                    el.click();
                    __valid = true;
                }
            })
        """.replace("FILTER_CARDS", FILTER_CARDS)
        result = await page.evaluate(code)

    print("Cards filtered")

    time.sleep(2)  # Make sure filter has applied

    # Scroll through the page
    time_required = await page.evaluate("""
        function scroll() {
            if (document.body.scrollHeight > window.scrollY + window.innerHeight) {
                window.scrollTo(0, Math.min(
                    document.body.scrollHeight,
                    window.scrollY + window.innerHeight
                ));

                setTimeout(scroll, 500);

                return (document.body.scrollHeight - window.scrollY) / window.innerHeight;
            }

            return 0;
        }

        scroll();
    """, force_expr=True)

    print("Time Required To Scroll:", time_required)

    if time_required > 0:
        time.sleep(time_required + 0.5)

    print("Page scrolled")

    # Export Cards
    cards = await page.evaluate("""
        (function() {
            var results = []
            var elements = document.querySelectorAll(".EmployeeCardContainer")
            for (var i = 0;i < elements.length;i++) {
                data = {
                    image: elements[i].querySelector("img[alt=profile]").src,
                    name: elements[i].querySelector(".JobInfo__name").innerText,
                    role: elements[i].querySelectorAll(".JobInfo__text")[0].innerText.split(" in ")[0],
                    team: elements[i].querySelectorAll(".JobInfo__text")[0].innerText.split(" in ")[1],
                    office: elements[i].querySelectorAll(".JobInfo__text")[1].innerText,
                    org: elements[i].querySelectorAll(".JobInfo__text")[2].innerText,
                    email: elements[i].querySelectorAll(".ContactInfo__dataContainer")[0].innerText,
                    phone: elements[i].querySelectorAll(".ContactInfo__dataContainer")[2].innerText,
                }

                results.push(data);
            }

            return results;
        })();
    """, force_expr=True)

    print("Cards exported")

    # Close Browser
    await page.close()
    await browser.close()

    print("Browser closed")

    # Open Outlook
    outlook = win32.gencache.EnsureDispatch("Outlook.Application")

    print("Outlook opened")

    # Retrieve existing contacts
    contacts = outlook.Session.GetDefaultFolder(10)

    existing_contacts = {}

    for i in range(contacts.Items.Count):
        card = contacts.Items[i + 1]
        existing_contacts[card.Email1Address] = card

    print("Sorted existing contacts")

    for new_contact in cards:
        if new_contact["email"] in existing_contacts:
            card = existing_contacts[new_contact["email"]]
        else:
            card = outlook.CreateItem(2)

        card.FullName = new_contact["name"]
        card.JobTitle = new_contact["role"]
        card.OfficeLocation = new_contact["office"] + " (" + new_contact["org"] + ")"
        card.CompanyName = new_contact["team"]
        card.Email1Address = new_contact["email"]
        card.MobileTelephoneNumber = new_contact["phone"]

        if not card.HasPicture:
            print("Fetching contact image:", new_contact["image"])
            r = requests.get(new_contact["image"], stream=True)
            if r.status_code == 200:
                with open(tmp_img_path, "wb") as tmp_img:
                    for chunk in r.iter_content(1024):
                        tmp_img.write(chunk)

                card.AddPicture(tmp_img_path)

        card.Save()

    print("Contacts synced")

    if os.path.exists(tmp_img_path):
        os.remove(tmp_img_path)

    outlook.Quit()

    print("Outlook Closed")
    print("Done")


asyncio.get_event_loop().run_until_complete(main())
