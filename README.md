# python-spreadsheet-practice

This project have 2 program:
1.  Read xlsm file in local and write some data
2. Read Google sheet and write some data


# ReadWriteGoogleSheets.py Docs

# Why I do this?

Just have fun want to play with google sheet api by python to see what I can do :)

# Program Purpose
Create new Col Values

By run ReadWriteGoogleSheets.py

This program will do 3 things:

1. Read Inventory Google Sheet and calculate each product total inventory values.

2. Copy Col Title Format For New Col

3.  Write Google Sheet

## before
![](https://i.imgur.com/MXwLj3G.png)

## after
![](https://i.imgur.com/F1lAGvY.png)

# How to use

1. Download this project

2. Create secret dir

3. Go to Google Developer Console
https://console.developers.google.com
4. Create new project

5. Enable Google Sheet API
![](https://i.imgur.com/8kBlo1u.png)
6. Create service account creds
![](https://i.imgur.com/t92Pgq8.png)
7. Download creds json file and put into secret dir

8. Rename json file to google-api-secret.json

9. Create copy file in your google drive

![](https://i.imgur.com/LjFpT90.png)

ps: If your file is xlsx extension you have to cover to google sheet format (see below image)

![](https://i.imgur.com/deFLQC3.png)


10. open your creds json file and copy your email

11. share google sheet with your service account by paste

![](https://i.imgur.com/JTtqF36.png)

12. Run ReadWriteGoogleSheets.py

13. Enter your spreadsheet id (see in google sheet url)

```
https://docs.google.com/spreadsheets/d/<your spreadsheetId>/edit#gid=<sheet id>
```

Done!

You should see your google sheet now have E1:E75 values data.
