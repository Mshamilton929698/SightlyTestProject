# SightlyTestProject

This is a C# Selenium project. The test file will navigate a stagning website to download a report and then compare the data to a master file.
The project was created using an NUnit Template.

This project just contains a single test case. If this project was intended to run a suite of tests, a Base Class and helper files would need to be added so shared methods 
and classes would not need to be added to every test file. 

### Target Framework

- netcoreapp3.1

### Packages Included

- ClosedXML -- Using to read and extract the data from the Excel files
- NUnit
- NUnit3TestAdapter
- Selenium.Support
- Selenium.WebDriver
- Selenium.WebDriver.ChromeDriver

### Run Test Steps
- To run the test, assumes that Visual Studio 2019 is installed

1. Clone the repo. The test is designed to use a master xlsx file to compare the downloaded report against. It is included in the project folder "ComparisonFiles" 
2. Build the Solution.
3. Select Test > Test Explorer 
4. Click on the Run icon to run the test
5. Test should pass by validating the data and then cleaning up the downloaded file. 
6. If any data within the main table of the master file is changed, the test should fail. If the test fails, the downloaded file is not deleted so it can be inspected.
