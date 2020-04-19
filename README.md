## Required before you run
1. Install dependencies
  ```
  pip install -r requirements.txt
  ```
2. Follow all the steps under "Using Signed Credentials" in [this page](https://gspread.readthedocs.io/en/latest/oauth2.html#using-signed-credentials)
3. rename the json file you downloaded to 'credentials.json' and it must be in the same directory with merge-cli.py
4. Share the spreadsheets that will be merged (hopefully) with client email found in the json file with edit access

## TODO
* Change authentication method
* Change input method to a json file
