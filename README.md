# NEAT_Invoice_Exporter
WebDriver program for automatically exporting multiple customer invoices from Neat.

Since, for reasons beyond me, [Neat](https://www.neat.com/) does not have an export feature I created this webdriver to export a large number of invoices automatically. The program goes through the designated number of invoices and pulls the invoice amount and customer name. It then stores the invoice values for each customer in an Excel sheet.

## Requirements
1. Will require Microsoft Edge WebDriver to be installed. Download can be found [here](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/).

2. Make sure to create the required blank Excel sheet in the program folder. It should be named "YearlyInvoices.xlsx".

3. Only tested on Windows. Will likely not work on other operating systems due to use of MS Edge WebDriver. To make it run on other operating systems you would need to replace the MS Edge driver with another browser.

## Usage Notes
- Make sure to first read the code comments and make the necessary changes indicated by a "NOTE:..." comment.

- You will need to make sure and replace the counter with the invoice numbers you wish to export. By default it exports the first 322 invoices on your account.

## Developer Notes
This is a lot more stable and reliable than my previous WebDriver programs but still occasionally crashes for unknown reasons. Luckily progress is saved on the Excel sheet so if it crashes it just needs to be restarted and it should pick up where it left off.

I had a tight timeline to complete this by so there's a couple areas I'm not totally satisfied with but it doesn't need to be pretty it just needs to work :sweat_smile:. Also made use of the *pandas* package for the first time and I can see now why it's so popular. It works very well and can definitely see a lot of possibilies and uses for it in the future. Especially when it comes to data analysis. 
