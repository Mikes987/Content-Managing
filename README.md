# Content-Managing with VBA
### Prepare DataSheets given by specific Database and reform it to be imported in the correct appearance

This project was a short term project in order to automate certain processes concerning:
- database export
- database import
- service for supplier

The project is completely written in Excel VBA and all userforms and modules in this repository are meant to be saved within one file.
The macros handle data information that is given by the database iPIM of Novomind. Data can be exported as excel files that can be used to create supplier data sheets.
These supplier sheets can then be handled for automatic import within the database.

In General, the code was written within a german environment, so all content was in German. I translated comments and certain strings (that are mostly used for match findings) into english.

But this code is right now for presentation only, since all information of content is in German, it is not possible to demonstrate the function of code as it is presented here.

The goal was to write macros in a way, people are able to use them without seeing the code itself or going into the developing environment. Four Buttons are placed on the first sheet which will then open up userforms. Within these userforms, Data files can be loaded and certain processes can be applied.

The buttons will activate the following processes:
- Create supplier product data sheets as a new excel file
- Transform default values into their associated IDs of database within this excel after receiving from supplier
- Change shape of data file to prepare for import into database
- export current data content into supplier product data sheet for supplier to check if supplier is still confident with content
