# HSCBoardMarksScraper

This implements Selenium browser automation to scrape data from website.

Problem : To get results of a whole class from Maharashtra state results website for a class teacher to submit the same to supervisors. As it is a tedious process, automation is required.

Implementation : For automation, Selenium using chrome driver was used in Python. XPath and Element name of required data fields were hardcoded. For I/O, XLRD and XLUTILS is used here.

Input: Excel sheet(.xls) containing Roll Numbers and Mother's name required in order to check the result.

Output : Same format as the input excel sheet along with the marks for the respective subjects added.
