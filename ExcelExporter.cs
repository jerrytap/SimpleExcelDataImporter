using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace GenericUtils.Data {
    public class ExcelExporter:DataImporter {
        private static object lockObject = new object();
        private static string exportLocation;
        private string tableName;

        public string TableName {
            get { return tableName; }
            set { tableName = value; }
        }
        private string emptyExcelFile;

        public string EmptyExcelFile {
            get { return emptyExcelFile; }
            set { emptyExcelFile = value; }
        }

        public static string ExportLocation {
            get { return exportLocation; }
            set { exportLocation = value; }
        }


        private string CreateEmptyExcelFile(string filename) {
            lock (lockObject) {
                System.IO.File.Copy(ExportLocation + "/" + EmptyExcelFile, ExportLocation + "/" + filename);
            }
            return ExportLocation + "/" + filename;
        }

        public void WriteDataTableToNewFile(string filename, DataTable dt) {
            string fullFilePath = CreateEmptyExcelFile(filename);
            WriteDataTableAsSheet(GetOleToExcelConnectionString(fullFilePath), dt);
        }

        private void WriteDataTableAsSheet(string connectionString, DataTable dt) {
            using (OleDbConnection conn = new OleDbConnection(connectionString)) {
                conn.Open();
                //we decided to just use the formatted sheet in excel
                //using (OleDbCommand cmd = conn.CreateCommand()) {
                //    cmd.CommandText = BuildCreateSheetTabStatement(dt);
                //    cmd.ExecuteNonQuery();
                //}
                foreach (DataRow dr in dt.Rows) {
                    using (OleDbCommand cmd = conn.CreateCommand()) {
                        cmd.CommandText = BuildInsertStatement(dr);
                        cmd.ExecuteNonQuery();
                    }
                }

            }
        }

        private string BuildInsertStatement(DataRow dr) {
            string insertStatement = "Insert INTO [" + tableName + "$]";

            StringBuilder insertColumns = new StringBuilder();
            StringBuilder insertValues = new StringBuilder();

            int colCount = dr.Table.Columns.Count;
            if (colCount > 0) {

                insertColumns.Append(" ([" + CleanStringForODBC(dr.Table.Columns[0].ColumnName));
                insertValues.Append(" Values ('" + CleanStringForODBC(dr.GetStringFromDataRow(dr.Table.Columns[0].ColumnName, "")));

                for (int currCol = 1; currCol < colCount; currCol++) {
                    insertColumns.Append("],[" + CleanStringForODBC(dr.Table.Columns[currCol].ColumnName));
                    insertValues.Append("','" + CleanStringForODBC(dr.GetStringFromDataRow(dr.Table.Columns[currCol].ColumnName, "")));

                }
                insertColumns.Append("])");
                insertValues.Append("')");

            }
            return insertStatement + insertColumns.ToString() + insertValues.ToString();
        }

        private static string CleanStringForODBC(string input) {
            if (input.Contains("'")) {
                return input.Replace("'", "''");
            }
            return input;
        }

        private static string BuildCreateSheetTabStatement(DataTable dt) {
            StringBuilder createStatement = new StringBuilder(string.Format("CREATE TABLE [{0}] ", dt.TableName));
            int colCount = dt.Columns.Count;
            if (colCount > 0) {
                createStatement.Append(string.Format("([{0}] TEXT", dt.Columns[0].ColumnName));
                for (int currCol = 1; currCol < colCount; currCol++) {
                    createStatement.Append(string.Format(", [{0}] TEXT", dt.Columns[currCol].ColumnName));
                }
                createStatement.Append(")");
            }
            return createStatement.ToString();
        }
    }
}
