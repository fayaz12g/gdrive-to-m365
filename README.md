# gdrive-to-m365
 A converter for Google Drive gdocs, gslides, and gsheets files to Microsoft Office file types docx, pptx, and xlsx


This is an awesome program that converts the Google Drive synced files to things you can open in any Microsoft app by downloading them all at once, even if you aren't the owner!

Before starting the python program, install the dependencies. To do this all at once, run: pip install customtkinter google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client


Before the program will work, you need to add the Google Drive API. To do this, create a free project in Google Cloud Console, then follow these steps.

Step One:
[./instructions/stepone.png]

Click "APIs and Services" on the main screen

Step Two:
[./instructions/steptwo.png]

Search for "Google Drive API" and click on it.

Step Three:
[./instructions/stepthree.png]

Click "Turn On". In the screenshot above, it says "Manage" as I already have it active.

Step Four:
[./instructions/stepfour.png]

Follow the three steps in the image, first clicking back, then selecting the "Credentials" tab.
Next, click "Create Credentials".
Then, select "OAuth client ID"

Step Five:
[./instructions/stepfive.png]

Now, select "Desktop app" under "Application type", and name it whatever you want. Here I named it "Converter"

Step Six:
[./instructions/stepsix.png]

Then click "DOWNLOAD JSON" to download the credentials you just created.

Step Seven:
[./instructions/stepseven.png]

Lastly, rename the downloaded .json file to "credentials.json" and place it next to the python script.

And then you're done! It will work. 
