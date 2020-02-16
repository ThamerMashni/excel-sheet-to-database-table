using System;
using System.Collections.Generic;
using System.Web;
using WebMatrix.Data;
using System.DirectoryServices.AccountManagement;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Linq;
using System.Text.RegularExpressions;

namespace myExcel
{
    /// <summary>
    /// Summary description for myExcel
    ///   - This component allows you to insert rows to a table in the database from an excel file, 
    ///   - the component will validate all the rows first, if all the rows are valid then the insertion to the datebase starts.
    ///  
    /// NOTE: date format in Excel MUST BE like: 11/21/2018 12:00:00 AM, MM/dd/yyyy hh:mm:ss tt
    /// 
    /// 
    /// ~~~~~~~~~~~~~~~~~~IMPORTANT: MAKE SURE THAT EXCEL FILE CULOMNS DOES NOT CONTAIN COMMAS, BECAUSE IT IS THE MECHANISM TO SEPRATE COLUMN VALUES:~~~~~~~~~~~~~~~~~~
    /// 
    ///  Classes: 
    ///  1- general: abstract class contains the commen functions among other classes.
    ///  
    ///  2- DatabaseTable: reguler class, its main purpose is to handle the conections with the database, mapping to a real table in the database, and inserting to the database.
    /// 
    ///  3- excelWorksheet: reguler class, that takes input of type HttpPostedFileBase(it expict an excel file, no validation so far), then using the function  excelToDb() passing
    ///  as a parameter an object of type DatabaseTable, there where the component will extract the excel rows and validate them against the databaseTabel passed in the parameter, 
    ///  then call the function DatabaseTable.addRow().
    ///  
    /// 
    /// </summary>
    /// 


    public abstract class general
    {
        public Database makeConnection()
        {
            return Database.Open("pms");
        }
        public string getCurrentUser()
        {
            return HttpContext.Current.User.Identity.Name.Split('\\')[1].ToUpper();
        }
        public string getCurrentUserFullname()
        {
            string fullName = "";
            using (var context = new PrincipalContext(ContextType.Domain))
            {
                var principal = UserPrincipal.FindByIdentity(context, HttpContext.Current.User.Identity.Name);
                fullName = principal.DisplayName;
            }
            return fullName;
        }
        public string errorMessge { get; set; }
        public bool isError { get; set; }
        public string convertListToSQLReady(List<string> list)
        {
            string[] a = list.ToArray();
            int number;
            DateTime dt;
            for (int i =0; i < a.Length; i++) // Iterate throgh the row values and wrap the non-integers with "'"
            {
                if (DateTime.TryParse(a[i], out dt))
                {
                    string aVal = a[i];
                    a[i] = "To_Date( '" + dt.ToString("yyyy-MM-dd hh:mm:ss tt") + "', 'YYYY-MM-DD HH:MI:SS AM')";////To_Date('2019-11-14 07:51:10 AM', 'YYYY-MM-DD HH:MI:SS AM');
                }
                else if (!int.TryParse(a[i],out number))
                {
                    string aVal = a[i];
                    a[i] = "'" + aVal + "'";
                }
            }
            string SQLReady =  String.Join(String.Empty, string.Join(",", a)) ;
            return SQLReady;
        }
    }
    public class DatabaseTable : general
    {
        //// number of rows (no need for this)
        //public int rowCount { get; set; }

        // number of column 
        public int columnCount { get; set; }

        //Name of the table in the database, e.g T_IIR_HR_TABLE
        public string TableName { get; set; }

        // column names in csv format e.g columnsNames ="ID,NAME,BADGE"
        public string columnsNames { get; set; }

        //columns data types that in csv format e.g string columnRule = "int,int,string"; 
        // ORDER IS IMPORTANT
        public string columnsRule { get; set; }

        // the sequance name generated in SQL Developer i.g idSequence ="SQ_IIR_HR_DASH_CONTRACTING" 
        public string idSequence { get; set; }

        // allowed data types for initial columns validation, by default valid types are: string, int,decimal, DateTime
        public string[] validDataTypes
        {
            get { return new string[] { "string", "int", "decimal", "DateTime"}; }
            set { }
        }

        // constructor that takes four paramerters, TableName:"T_IIR_Table_X", CulomnsNames in CSV format :"ID,NAME,BADGE", ColumnsRules in CSV format Also:"int,string,string", and sequance name generated in SQL Developer idSequence ="SQ_IIR_HR_DASH_CONTRACTING" 
        public DatabaseTable(string TableName, string columnsNames, string columnsRule, string idSequence)
        {
            this.TableName = TableName;
            // this.rowCount = rowCount;
            this.columnsNames = columnsNames;
            this.columnsRule = columnsRule;

            this.idSequence = idSequence;

            this.columnCount = columnsNames.Split(',').Length; // count the columnsNames

            this.validataColumns();
        }

        public string errorMessage { get; set; }
        //public bool isError { get; set; }

        public void modifyValidDataTypes(string[] newValidTypes)
        {
            this.validDataTypes = newValidTypes;
        }

        //validate the CoulmnsRules against the validDataTypes[]  
        public void validataColumns()
        {
            string[] columnsType = columnsRule.Split(',');
            if (columnCount != columnsType.Length)
            {
                isError = true;
                errorMessge = "Columns rules length does not equal number of columns!";
            }
            bool isValidType = false;
            foreach (string type in columnsType) // check if all 
            {
                foreach (string validType in validDataTypes)
                {
                    if (validType.Equals(type))
                    {
                        isValidType = true;
                    }
                }
                if (isValidType == false)
                {
                    isError = true;
                    errorMessge = "one or more datatype in columns rule is not valid (not in validDataTypes)";
                    return;
                }
            }


        }

        public void addRow(string csvRow)
        {
            string id = idSequence + ".nextVal,";
            Database db = this.makeConnection();
            string query = "INSERT INTO " + TableName + " (" +"ID,"+ columnsNames + ") VALUES ("+ id + csvRow + ")";
            db.Execute(query);
            db.Close();

        }
        public void addRowNoID(string csvRow)
        {
            // string[] values = csvRow.Split(',');
            Database db = this.makeConnection();
            string query = "INSERT INTO " + TableName + " ("  + columnsNames + ") VALUES (" + csvRow + ")";
            db.Execute(query);
            db.Close();

        }

        public IEnumerable<dynamic> getAllRows()
        {
            Database db = this.makeConnection();
            string query = "SELECT " + this.columnsNames + " FROM " + TableName;
            IEnumerable<dynamic> rows = db.Query(query);
            return rows;
        }

        // validate the row before inserting it.
        public bool isValidRow(string csvRow, int row)
        {
            string[] rules = this.columnsRule.Split(',');
            string[] values = csvRow.Split(',');

            if (rules.Length != values.Length)
            {
                isError = true;
                errorMessage = errorMessage + " \r\n Row: " + row + ", the number of rules does not equal to the number of columns in row number: ";
                return false;
            }
            int number;
            decimal num;
            DateTime dt;
            for (int i = 0; i < rules.Length; i++) 
            {
                if (rules[i] == "int" &&(!string.IsNullOrWhiteSpace(values[i]) && !int.TryParse(values[i], out number))) // allow null values with int Rule
                {
                    isError = true;
                    errorMessage = errorMessage + " \r\n Row: " + row + ", one or more columns type contradict with the column rules! (int: "+values[i]+")";
                    return false;
                }
                if (rules[i] == "decimal" && !decimal.TryParse(values[i], out num))
                {
                    isError = true;
                    errorMessage = errorMessage + " \r\n Row: " + row + ", one or more columns type contradict with the column rules! (decimal)";
                    return false;
                }
                if (rules[i] == "string")
                {
                    Regex tagRegex = new Regex(@"<\s*([^ >]+)[^>]*>.*?<\s*/\s*\1\s*>");
                    bool hasTags = tagRegex.IsMatch(values[i]);
                    if (hasTags)
                    {
                        isError = true;
                        errorMessage = errorMessage + " \r\n Row: " + row + ", a column may contains a script!";
                        return false;
                    }
                }
                //dateFormat yyyy-MM - dd
                //Regex dateFormat = new Regex(@"^([0-9]{4}[-/]?((0[13-9]|1[012])[-/]?(0[1-9]|[12][0-9]|30)|(0[13578]|1[02])[-/]?31|02[-/]?(0[1-9]|1[0-9]|2[0-8]))|([0-9]{2}(([2468][048]|[02468][48])|[13579][26])|([13579][26]|[02468][048]|0[0-9]|1[0-6])00)[-/]?02[-/]?29)$");
                //// '2018-11-21 10:34:09 PM'
                //Regex dateFormat = new Regex(@"[0-9]{4}-[0-9]{2}-[0-9]{2}");=
                //Regex dateFormat = new Regex(@"[0-9]{4}/[0-9]{2}/[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2} [APM]{2}");
                //bool isDate = dateFormat.IsMatch(values[i]);

                if (rules[i] == "DateTime" && ( !string.IsNullOrWhiteSpace(values[i]) && !DateTime.TryParse(values[i], out dt))) // this is with ALLOW NULL DATE
                {
                    isError = true;
                    errorMessage = errorMessage + " \r\n Row: " + row + ", one or more columns type contradict with the column rules! (DateTime)";
                    return false;
                }

            }
            return true;
        }
    }
    public class excelWorksheet : general
    {
        public HttpPostedFileBase excelFile { get; set; }

        public excelWorksheet(HttpPostedFileBase actionFile)
        {
            excelFile = actionFile;
        }
        //this function takes as a parameter an object of type DatabaseTable (custom class created above)
        public void excelToDb(DatabaseTable table)
        {
            var package = new ExcelPackage(excelFile.InputStream);

            ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault();
            var start = workSheet.Dimension.Start;
            var end = workSheet.Dimension.End;


            int columnRange = workSheet.Dimension.Columns;//get how many columns
            if (columnRange != table.columnCount)
            {
                isError = true;
                errorMessge = "excel sheet and the table: " + table.TableName + " is not having the same number of columns!";
                return;
            }
            ///
            // a List to save all valid rows in it, if all row are valid then insert.
            List<string> csvRows = new List<string>();
            ///
            /// create a list of stirng to save the row values temporerly, only one row at a loop iteration will be saved on it. 
            List<string> rowValues = new List<string>();

            ///
            ///the loop will iterate and save each row then use
            ///convertListToSQLReady() function to convert the list to a string ready to be part of a query i.g. "1,THAMER,1234"
            ///finally insert it into the database table
            for (int row = start.Row + 1; row <= end.Row; row++) // (+1 to skip the header of the template)
            { // Row by row...  
                rowValues.Clear(); // empty the previous row

                // this loop will fill the list with a row value
                for (int col = start.Column; col <= end.Column; col++)
                { // ... Cell by cell...  
                    rowValues.Add(workSheet.Cells[row, col].Text);
                }
                if (rowValues.Count > 0)
                {
                    string[] arr = rowValues.ToArray();
                    string csvValues = string.Join(",", arr);
                    // add validation here 
                    if (!table.isValidRow(csvValues, row)) // if the row is not valid stop and return 
                    {
                        return;
                    }
                    else
                    {
                        string csvRow = convertListToSQLReady(rowValues); // convert list to csv format
                        csvRows.Add(csvRow); // add valid row to the list to insert later if all rows are valid
                    }
                }
            }//out of excel sheet


            //inserting to DB
            if (csvRows != null)
            {
                foreach (string row in csvRows)
                {
                     table.addRow(row);// finally add row to the database passed as parameter.
                    //table.addRowNoID(row);
                }
            }

        }
    }
    
}