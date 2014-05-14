XLSToCSVConeVitor
=================

### Bulk Convert XLS file to CSV or Text with any delimiter. 


How to Run - 
----
````
1. Download the src folder

2. Rename src\lib\poi-3.10-FINAL-20140208.jar.txt to src\lib\poi-3.10-FINAL-20140208.jar

3. Open dir.properties file in notepad.

   + Mentioned the directory in which all the XLS file are present against the property 'DIR'
      Use \\ instead of \ in the directory path. 
      e.g. DIR=D:\\Workspace\\sheet
   + Mention the output file type  against the property 'FILE_TYPE'
      e.g. FILE_TYPE=csv 
   + Mention the desired delimiter (separater) against the property 'DELIMITER'
      e.g. DELIMITER=,

4. Run the Conevitor.bat file to generate CSV or TXT files in the same folder as the XLS files
````

API used for XLS convertion 
---------------------------
[Apache POI](http://poi.apache.org/download.html)
