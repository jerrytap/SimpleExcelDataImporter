using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace GenericUtils.Data {
	//I typically don't care for comments all that much but I'm trying out Git.
    public class DataImporter {
        private static object lockObject = new object();
        private static string importLocation;

        public static string ImportLocation {
            get { return DataImporter.importLocation; }
            set { DataImporter.importLocation = value; }
        }


        protected static string GetOleToExcelConnectionString(string pFilePath) {
            string strConnectionString = string.Empty;
            string strExcelExt = System.IO.Path.GetExtension(pFilePath);

            if (strExcelExt == ".xls") {
                strConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties= ""Excel 8.0;HDR=YES""";
            } else if (strExcelExt == ".xlsx" || strExcelExt == ".xlsm") {
                strConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES""";
            } else {
                throw new ApplicationException("Excel file extenstion is not known.");
            }
            return string.Format(strConnectionString, pFilePath);

        }

        public static DataSet ReadFile(string filename) {
            if (filename.ToLower().Contains(".xls")) {
                DataSet excelFileContents = new DataSet();
                string connectionString = GetOleToExcelConnectionString(filename);
                List<string> sheetNames = GetSheetNames(connectionString);
                foreach (string sheetName in sheetNames) {
                    excelFileContents.Tables.Add(GetSheetAsDataTable(connectionString, sheetName));
                }
                return excelFileContents;
            }
         
            throw new ApplicationException("Excel file extenstion is not known.");

        }

        private static void CopyColumnNamesToFirstRecord(DataTable dt) {
            foreach (DataColumn column in dt.Columns) {
                string newColumnName = dt.Rows[0][column].ToString();
                if (string.IsNullOrEmpty(newColumnName)) {
                    //column.ColumnName = ".";
                } else {
                    column.ColumnName = newColumnName;
                }
            }
            dt.Rows.RemoveAt(0);
        }



        protected static string BuildSelect(string sheetName) {
            return string.Format("SELECT * FROM [{0}$]", sheetName);
        }

        protected static DataTable GetSheetAsDataTable(string connectionString, string sheetName) {
            DataTable dt = new DataTable(sheetName);
            using (OleDbConnection conn = new OleDbConnection(connectionString)) {
                conn.Open();
                using (OleDbCommand cmd = conn.CreateCommand()) {
                    cmd.CommandText = BuildSelect(sheetName);
                    using (OleDbDataReader reader = cmd.ExecuteReader()) {
                        if (reader.Read()) {
                            DataTable schemaTable = reader.GetSchemaTable();
                            foreach (DataRow row in schemaTable.Rows) {
                                dt.Columns.Add(row[0].ToString());
                            }

                            object[] rowValues = new object[reader.FieldCount];
                            for (int currRow = 0; currRow < reader.FieldCount; currRow++) {
                                rowValues[currRow] = reader[currRow];
                            }
                            dt.Rows.Add(rowValues);
                        }
                        while (reader.Read()) {
                            object[] rowValues = new object[reader.FieldCount];
                            for (int currRow = 0; currRow < reader.FieldCount; currRow++) {
                                rowValues[currRow] = reader[currRow];
                            }
                            dt.Rows.Add(rowValues);
                        }
                    }
                }
            }
            return dt;
        }

        protected static List<string> GetSheetNames(string connectionString) {
            List<string> sheets = new List<string>();
            using (OleDbConnection conn = new OleDbConnection(connectionString)) {
                conn.Open();
                DataTable tbl = conn.GetSchema("Tables");
                foreach (DataRow row in tbl.Rows) {
                    string sheetName = (string)row["TABLE_NAME"];
                    if (sheetName.EndsWith("$")) {
                        sheetName = sheetName.Substring(0, (sheetName.Length - 1));
                    }
                    if (!sheets.Contains(sheetName) && !sheetName.Contains('$') ){
                        sheets.Add(sheetName);
                    }
                }
            }
            return sheets;
        }

    }
}
