# hgarc_xwalk
## Cross-walk for marcxml to excel and back to marcxml for HGARC bib records

These python files are designed to convert marcxml records to an excel spreadsheet for non-catalogers to be able to edit the content of bibliographic records and to convert the updated spreadsheet back to marcxml for load into a library system.

### MarcXML2Excel.py

This script looks for one or more marcxml files to process. It reads the marcxml file and creates a dataframe with the following columns. A row in the dataframe is created for each field/subfield in the marcxml record

CollectionNum
oclcnum
field_type
marc_tag
ind1
ind2
field_valuesub_code
sub_value

The DataFrame is written to an Excel spreadsheet

### Excel2MarcXML.py

This script loads an existing Excel spreadsheet to create a DataFrame. Iterating through the rows of the DataFrame, a new marcxml record is created and added to a collection of records that are written out in files of 100 records.

### hgarc_oclc_recs.xlsx

Example spreadsheet format read and created by the two python scripts

### hgarc_100.xml

Example MarcXML file downloaded from OCLC and the source of records converted to the spreadsheet

### hgarc-updated-records200.xml

Example MarcXML file created by the Excel2MarcXML.py script




