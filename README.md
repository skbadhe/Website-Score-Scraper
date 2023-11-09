# HSCBoardMarksScraper

This implements Selenium browser automation to scrape data from website.

Problem : To get exam results of a set of candidates from official state website by using a excel sheet full of information and get results in another excel sheet.

Implementation : For automation, Selenium using chrome driver was used in Python. XPath and Element name of required data fields were hardcoded. For I/O, XLRD and XLUTILS is used here.

Input: Excel sheet(.xls) containing details (Roll Numbers and Mother's name).

Output : Same format as the input excel sheet along with the scores and calculated grades.
