Feature: Parsing del file
Scenario: Open a existent file
     Given document called "sample.xls"
       And sheet 0
      Then The sheet exists     
       And the row 1 and the column 1 contains 'type'
       And the row 73 and the column 5 contains 1234567890.1234567890
       And the row 2 and the column 3 is empty