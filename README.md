# Notes:

- To run, add a python file called filepaths.py that houses these variables:

    ```python
    data = {
        'stored_path': [
            #excel sheet(s) that are logged to.
        ],
        'share_path': [
    	    #path(s) of a file (csv) with the names of shares that are to be tracked.
        ],
        'base_url': [
            #url(s) that is to be scraped.
        ],
        'ext_url': [
            #the variable extension(s) for the said url(s).
        ]
        'parameters' : [
            #the parameters for the scraping for the different sites.
        ]
    }
    ```

- The variables in this program are all lists to allow for multiple excel sheets that are created by scraping parameterized data from different urls.
- The program initialise_db is to be run once to initialise the excel sheet with the names of the stocks.
- Following that the store_data program is to be run once a week (or more) - you could schedule this on a UNIX device - and it retrieves all the data available for the week. Data is available for Monday through Friday and can be only accessed the next day. Running the program on a Saturday or Sunday would retrieve data for the entire week.

# Outline:

- A simplistic scraper that cleans up data regarding share prices for a list of pre-defined shares or all shares based on daily NSE reports. It can find data for multiple parameters based on the users preference and stores the data neatly in an excel sheet, presenting date and the data for the parameters. It is soon going to be updated to include rudimentary analysis and visualisation as well.

# Purpose:

- This project helped me solidify my understanding of pandas and working with csv files as well as MultiIndexed dataframes. While the build seems relatively easy at first glance - accounting for the various parameters that the users could choose, a preset share list versus all the shares (that change everyday) as well as having to deal with the fact that the reports are not available on certain days made it so that to write clean code efficiently, I had to plan the program thoroughly and document it well.
- With reusability in mind, I therefore implemented the filepaths program that would allow me to build multiple excel sheets on the same principle simultaneously.

# Description:

- The project has 2 key programs: initialise_db and store_data. Let us look at each of them individually.
- First, initialise_db:
    - This program is responsible for initialising the excel sheet with a list of shares, either through a list provided or by scraping the data for the names recorded at that instant.
    - The cells are formatted with the alignment and borders all taken care of.
- Next, store_data:
    - The purpose of store data is simple - it is meant to scrape a url for data for shares, see which shares, in the excel sheet already stored, match and fill in the corresponding cells with data for the parameters as decided by the user giving a result like this:

        ![alt text](https://github.com/akashvshroff/Stock_Market_Database/blob/master/Sample_Output/Output.png)

    - It determines the corresponding cells by querying the heading row and matching the dates stored to the date that it has just scraped and thereby determining the position of the column.
    - If a share does not have data present, it simply logs a '-' and if there are new cells present, it appends them to the bottom of the column and then adds the data for the share.
