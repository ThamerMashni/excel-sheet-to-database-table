# excel-sheet-to-database-table
using c# for web applications, you can insert data from excel sheet to a table in your database by uploading the excel file into a webpage.

# discription 
The main goal of this component is to insert an excel sheet into a table already exists in the database, by providing the information of the table, and uploading the excel sheet.

# Classes
  1- general: contains the commonly needed functions among other classes. 
    Main functions: 
      a-	public Database makeConnection()
      b-	public string getCurrentUser()
      c-	public string convertListToSQLReady(List<string> list): 
          1.	Input: list of type string. 
          2.	Output: string with the list values separated by comma with the non-integer values wrapped by “ ‘ “  .
          3.	Example: input = list<string> = <55, Ahmed , 1234>, output= “55,’Ahmed’,1234”

  2- DatabaseTable: regular class extends general
    Variables: 
        a-	int columnCount: number of columns in the database.
        b-	string TableName: name of the table in the database, (must match), e.g T_IIR_HR_TABLE
        c-	string columnsNames: names of the columns in the table in the database (must match), e.g columnsNames ="ID,NAME,BADGE" 
        d-	string columnsRule: columns data types in csv format e.g string columnRule = "int,string,int", NOTE: ORDER IS IMPORTANT, same as cloumnsNames
        e-	string idSequence: the sequance name generated in SQL Developer i.g idSequence ="SQ_IIR_HR_DASH_CONTRACTING"
        f-	string[] validDataTypes: allowed data types for initial columns validation, by default valid types are: string, int,decimal, DateTime, use function  modifyValidDataTypes(string[] newValidTypes) for one time change or edit the class.
  Main functions:
  a-	public DatabaseTable(string TableName, string columnsNames, string columnsRule, string idSequence): a constructor that takes four paramerters, TableName:"T_IIR_Table_X", CulomnsNames in CSV format :"ID,NAME,BADGE", ColumnsRules in CSV format Also:"int,string,string”, And sequance name generated in SQL Developer idSequence ="SQ_IIR_HR_DASH_CONTRACTING".
  b-	public validataColumns():validate the CoulmnsRules against the validDataTypes[].
  c-	public void addRow(string csvRow): input: csv string that has the values of one row. idSequence is not included in the csvRow parameter, by default it will be in the query.
  d-	public void addRowNoID(string csvRow): same as addRow(…) but idSequence is not included by default.
  e-	public bool isValidRow(string csvRow, int row): input: the row values, row number IN EXCEL to track where is the issue in excel. Output: bool value true or false. Mainly, used in excelWorksheet class to validate the row before inserting it.

3- excelWorksheet: regular class that extends general
  Variables:     
      a.	public HttpPostedFileBase excelFile : HttpPostedFileBase that should be the uploaded excel file. The component does not have a validation if the uploaded is excel file or not.
  Main functions:
        a.	public void excelToDb(DatabaseTable table): 
        Input: an object of the class DatabaseTable (documented above).
        Output: void.
        Functionality:
          1.	Check how many columns in the excel sheet, and validate it with the number of columns in the input table.
          2.	Validate all the rows in the uploaded excel sheet.
          3.	If a row is valid the method convertListToSQLReady() will be applied to the row to make it ready for the insertion phase, then it will be saved in the array “csvRows[]”.
          4.	If all the rows are valid, then all the rows in csvRows[] will be inserted to the input database table.


•	NOTE: THE FIRST ROW IN THE EXCEL SHEET IS ALWAYS SKIPPED, ASSUMING IT IS A HEADER ROW.


#  External Packages Used
	- OfficeOpenXml

# How to use it: 

IMPORTANT NOTE: MAKE SURE THAT EXCEL FILE CULOMNS DOES NOT CONTAIN COMMAS, BECAUSE IT IS THE MECHANISM TO SEPRATE COLUMN VALUES.


1. Create two objects one of each class: 
    DatabaseTable table = null;
    excelWorksheet uploadedSheet = null;

2. Create the HTML input ‘tileDataFile’ where the file will be uploaded:
<input type="file" class="form-control" id="tileDataFile" name="tileDataFile" accept=".xlsx,.xlsm,.xlsb, .xltx, .xltm, .xls, .xlt, .xml,">

3. Create the query String to get the uploaded file on POST: 

    if (IsPost){
HttpPostedFileBase tileDataFile = Request.Files["tileDataFile"];
}

4. Initialize the variable ‘table’ with the table name, columns, columnsRules and the sql sequence:
table = new DatabaseTable("T_III_TABLE1","NAME,BADGE,BIRTH_DATE","string,int,DateTime", "SQ_III_TABLE1_Sequence");

5. Now, pass the variable created in step #3 in the constructor of the  excelWorksheet variable:
      uploadedSheet = new excelWorksheet(tileDataFile);

6. Last step is to call the method excelToDb() passing into it the table. 
 
      uploadedSheet.excelToDb(table);

Now, the rows in the excel sheet should have been inserted in the DatabaseTable ‘table’.


For more information please see the ducomentation .doc or contact me.
 

