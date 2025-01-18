

# **Automated Assignment Writer**

This Python script automates the process of logging into a Microsoft account, creating a Word document, and writing your assignment based on a description you provide. It integrates the **Google Gemini AI** model to generate content and **Selenium WebDriver** to interact with the Microsoft login and Word online application.

## **Features**
- Generates content for your assignment using Google's Gemini AI model.
- Logs into your Microsoft account and creates a new Word document.
- Automatically pastes the generated content into the Word document.

## **Requirements**
Before running the script, you need to install the following Python packages:

- `google-generativeai` (for interacting with the Google Gemini AI model)
- `selenium` (for automating browser actions)
- `pyautogui` (for controlling the keyboard and mouse)
- `pyperclip` (for clipboard management)
- `python-dotenv` (for managing environment variables)

### To install these dependencies, run:

```bash
pip install google-generativeai selenium pyautogui pyperclip python-dotenv
```

## **Environment Variables (.env file)**

The script uses a `.env` file to securely store your API keys, Microsoft login credentials, and other necessary data.

Create a `.env` file in your project directory with the following structure:

```
APIKEY=<Your_Google_GenerativeAI_API_Key>
EMAIL=<Your_Microsoft_Email>
PASSWORD=<Your_Microsoft_Password>
```

### **Explanation of .env file:**

- **APIKEY:** Your API key for Google’s Gemini AI (get it from the [Google Cloud Console](https://console.cloud.google.com/)).
- **EMAIL:** Your Microsoft login email.
- **PASSWORD:** Your Microsoft login password.

Ensure that you don’t share or expose this `.env` file to others.

## **How to Run the Script**

1. Clone the repository to your local machine:
   ```bash
   git clone https://github.com/Abdulraheem232/Word-assignment-bot.git
   cd Word-assignment-bot
   ```

2. Install the required dependencies by running:
   ```bash
   pip install -r requirements.txt
   ```

3. Create a `.env` file as described above.

4. Run the script:
   ```bash
   python your_script_name.py
   ```

5. The script will ask you to describe your assignment, and it will generate and paste the content into a Word document in your Microsoft account.

## **Important Notes**
- **Do not use your keyboard or mouse while the script is running.** The script controls your mouse and keyboard, and any interruption may cause issues or unexpected behavior.
- **Occasionally, the script may not work properly.** Due to varying web page layouts or temporary issues with the services (Google Gemini, Microsoft login), the automation might fail. If the script fails, try again or check the script for errors related to web elements or API responses.
  
## **Known Issues**
- The script sometimes faces issues with web element loading, and may need adjustment in waiting times (`time.sleep()`).
- Web elements might change on Microsoft’s login page, which could break the script.
- **It may take a few seconds for the page to load and for the document to open. Please be patient.**

## **Important:**
- **Do not interact with the system while the program is running.** Touching the keyboard or moving the mouse will interfere with the automation and cause failure in the process. Simply sit back and let the script do the job.
- You can safely exit the program by typing `y` when prompted.

## **Disclaimer**
This is a personal project created for educational and automation purposes. It may not work flawlessly at all times and may require adjustments based on changes to the web pages, API, or environment.

---

### Example `.env` file (Create in the same directory as your script):

```
APIKEY=your_google_generativeai_api_key_here
EMAIL=your_microsoft_email_here
PASSWORD=your_microsoft_password_here
```

## **Author Information:**
- **Created by:** https://github.com/Abdulraheem232
- **Contact:**abdulraheemabdullah859@gmail.com

---
