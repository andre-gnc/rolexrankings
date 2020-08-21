# rolexrankings
Scrap some data of players in rolexrankings.com.

This is my project of web scraping. I try to do a job offer posted on Upwork. The detail was screenshoted as image file named screencapture-upwork-jobs-01578ee58f5ddae5d1-2020-08-17-11_32_20.png.

The employer need an output in csv/ xlsx as shown on file named Experience data.docx.

The project was succesfully done. I got approximately 1300 records of players. But i realize that the code is awful and hard to understand. Please give me your advices.

First, I used Selenium to load the whole document. There was a button that should be clicked to load it. Then I grabbed the data by Requests and BeautifulSoup. Last, I made an xlsx file to save them by xlsxwriter.

If you just test the code, i suggest to set the LOOP = 1 in line 13 and "break" in line 110. It will generate small xlsx file of a player's data. Otherwise, if the code will be run fully then set that LOOP more than 28 (LOOP = 40 for example) and set the "break" as comment (# break).

Once again, please give me your advice. Thank you. 