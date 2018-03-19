/* Author: Matthew Baker
 * Date Created: 03/17/2018
 * Date Modified: 03/17/2018
 * File: DistressTypes.cs
 * Program: Y2Y Comparison
 * Description: This file is used to create the database and format the SQL Querys used in this program.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace Y2Y_Comparison.Classes
{
    class DistressTypes

    {
        // Declare Variables used
        private static string RequiredDatabase;
        public string ExcelFile;
        public string excelFile2;
        public OleDbConnection myConnection;
        public OleDbConnection ExcelConnection;


        /* Function: DistressTypes()
         * Description: Default constructor used to set the variables and to create the database used in this program
         */

        public DistressTypes()
        {
            // Set the files used in the program
            RequiredDatabase = Environment.CurrentDirectory + @"\Resources\Database.mdb";
            ExcelFile = Environment.CurrentDirectory + @"\Resources\ExcelFormat.xls";

            // Create a connection to those files so SQL commands can be used from this program to manipulate those files.
            myConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + RequiredDatabase);
            ExcelConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data source=" + ExcelFile + "; Extended Properties=Excel 8.0;");

            // Clear Database
            DropTables();

            // Create the Tables in the new database.
            CreateDatabase();

            //Create a copy of the Excel file to have a fresh start next time
            CopyExcelFile();
        }

        /* Function: CopyExcelFile()
         * Description: Copy the excel file to a new name so that it can be copied back at the end for a fresh start.
         */

        private void CopyExcelFile()
        {
            excelFile2 = Environment.CurrentDirectory + @"\Resources\ExcelFormat2.xls";
            System.IO.File.Copy(ExcelFile, excelFile2, true);
        }
        /* Function: CopyExcelFileBack()
         * Description: Overwrite original excel with copy and delete the copy for a fresh start.
         */
        public void CopyExcelFileBack()
        {
            System.IO.File.Copy(excelFile2, ExcelFile, true);
            
        }

        /* Function: DropTables()
         * Description: Will clear the database of the ACP, CRCP, and JRCP tables if they exist
         *              Uses a try catch block to clear the database since OLEDB does not recognize
         *              the SQL statement IF EXISTS. If there is an error clearing it the try catch
         *              will just ignore the error and allow the program to continue.
         */

        public void DropTables()
        {
            myConnection.Open();
            OleDbCommand command = new OleDbCommand();
            command.Connection = myConnection;
            try
            {
                command.CommandText = "DROP TABLE ACP , CRCP, JRCP";
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }

            myConnection.Close();
        }


        /* Function: GetRequired()
         * Description: Return the path to the database if needed
         */

        public static string GetRequiredDB()
        {
            return RequiredDatabase;
        }


        /* Function: CreateDatabase()
         * Description: Will Create the tables needed in the new database
         */

        public void CreateDatabase()
        {
            // open a connection the database and create a command variable
            myConnection.Open();
            OleDbCommand Command = new OleDbCommand();
            Command.Connection = myConnection;

            // Create SQL Command for ACP Table and execute the command
            Command.CommandText = "Create Table ACP(" + ACP.CreateTableFormat() + ")";
            Command.ExecuteNonQuery();


            // Create SQL Command for CRCP Table and execute the command
            Command.CommandText = "Create Table CRCP(" + CRCP.CreateTableFormat() + ")";
            Command.ExecuteNonQuery();

            // Create SQL Command for JRCP Table and execute the command
            Command.CommandText = "Create Table JRCP(" + JRCP.CreateTableFormat() + ")";
            Command.ExecuteNonQuery();

            // Close Connection to Database.
            Command.Connection.Close();
        }
    }




    public class ACP
    {
        ACP()
        {

        }

        /* Function: CreateTableFormat()
         * Description: Will Create the headers for the ACP Table and declare their types
         */

        public static string CreateTableFormat()
        {
            return "Y2Y Text, NLENGTH NUMBER, NT_CR1_LF NUMBER,  NT_CR2_LF NUMBER, NL_CR1_LF NUMBER, NL_CR2_LF NUMBER,   NL_JT1_LF NUMBER, NL_JT2_LF NUMBER,   NRT_CR1_LF NUMBER,  NRT_CR2_LF NUMBER,  NRT_CR3_LF NUMBER,  NRL_CR1_LF NUMBER,  NRL_CR2_LF NUMBER,  NRL_CR3_LF NUMBER,  NA_CR1_SF NUMBER, NA_CR2_SF NUMBER,  NA_CR3_SF NUMBER, NPA_WP_SF NUMBER, NPA_NWP_SF NUMBER, NPOT_NO NUMBER, NDELAM_SF NUMBER, NBLEED1_SF NUMBER, NBLEED2_SF NUMBER, N_LDR NUMBER,  N_NDR NUMBER ";
        }

        /* Function: ItemsToGrab()
         * Description: Used to create the SQL Query for what information to get from the FWG Databases.
         */
        public static string ItemsToGrab()
        {
            return " NLENGTH , NT_CR1_LF ,  NT_CR2_LF , NL_CR1_LF , NL_CR2_LF ,   NL_JT1_LF , NL_JT2_LF ,   NRT_CR1_LF ,  NRT_CR2_LF ,  NRT_CR3_LF ,  NRL_CR1_LF ,  NRL_CR2_LF ,  NRL_CR3_LF ,  NA_CR1_SF , NA_CR2_SF ,  NA_CR3_SF , NPA_WP_SF , NPA_NWP_SF , NPOT_NO , NDELAM_SF , NBLEED1_SF , NBLEED2_SF , N_LDR ,  N_NDR ";

        }


        /* Function: ItemsSumAvg()
         * Description: Used to get the information that we want for each type in the FWG Database. Most of them
         *              are SUM of the Columns.
         */

        public static string ItemsSumAvg()
        {
            return "  SUM(NLENGTH ), SUM( NT_CR1_LF ), SUM(  NT_CR2_LF ), SUM( NL_CR1_LF ), SUM( NL_CR2_LF ), SUM(   NL_JT1_LF ), SUM( NL_JT2_LF ), SUM(   NRT_CR1_LF ), SUM(  NRT_CR2_LF ), SUM(  NRT_CR3_LF ), SUM(  NRL_CR1_LF ), SUM(  NRL_CR2_LF ), SUM(  NRL_CR3_LF ), SUM(  NA_CR1_SF ), SUM( NA_CR2_SF ), SUM(  NA_CR3_SF ), SUM( NPA_WP_SF ), SUM( NPA_NWP_SF ), SUM( NPOT_NO ), SUM( NDELAM_SF ), SUM( NBLEED1_SF ), SUM( NBLEED2_SF ), AVG(N_LDR ), AVG(N_NDR) ";
        }

        /* Function: Query()
         * Description: Will Format the Query that we want for transferring info from FWG database to our database
         */

        public static string Query(string FilePath)
        {
            return " INSERT INTO ACP ( " + ItemsToGrab()+") SELECT " + ItemsSumAvg() + " FROM [MS Access; DATABASE="+FilePath+"].ACP ";
        }

        /* Function: ExcelQuery()
         * Description: Will format the Query that we want for transferring info from our database to the Excel file
         */

        public static string ExcelQuery()
        {
            return "Insert into [ACP$] ( Y2Y," + ItemsToGrab() + ") SELECT Y2Y, " + ItemsToGrab() + " From [MS Access; DATABASE=" + DistressTypes.GetRequiredDB() + "].ACP";

        }
    };




    public class CRCP
    {
        CRCP()
        {

        }

        /* Function: CreateTableFormat()
         * Description: Will Create the headers for the CRCP Table and declare their types
         */

        public static string CreateTableFormat()
        {
            return "Y2Y Text, NT_CR1 NUMBER, NT_CR2 NUMBER, NT_CR3 NUMBER, NT_CR_NO NUMBER, NL_CR1 NUMBER, NL_CR2 NUMBER, NL_CR3 NUMBER, NCL_CR1_NO NUMBER, NCL_CR2_NO NUMBER, NCL_CR1_SF NUMBER, NCL_CR2_SF NUMBER, NCL_J_SP_LF NUMBER, CCL_J_SEAL CHAR, NPUNCH_NO NUMBER, NPUNCH_SF NUMBER, NC_PAT1_SF NUMBER, NC_PAT2_SF NUMBER , NC_PAT3_SF NUMBER, NA_PAT_SF NUMBER, CDR NUMBER, CPR NUMBER, N_CCI NUMBER ";
        }

        /* Function: ItemsToGrab()
         * Description: Used to create the SQL Query for what information to get from the FWG Databases.
         */

        public static string ItemsToGrab()
        {
            return "NT_CR1 , NT_CR2 , NT_CR3 , NT_CR_NO , NL_CR1 , NL_CR2 , NL_CR3 , NCL_CR1_NO , NCL_CR2_NO , NCL_CR1_SF , NCL_CR2_SF , NCL_J_SP_LF ,  NPUNCH_NO , NPUNCH_SF , NC_PAT1_SF , NC_PAT2_SF , NC_PAT3_SF , NA_PAT_SF , CDR , CPR , N_CCI  ";

        }


        /* Function: ItemsSumAvg()
         * Description: Used to get the information that we want for each type in the FWG Database. Most of them
         *              are SUM of the Columns.
         */

        public static string ItemsSumAvg()
        {
            return "SUM(NT_CR1 ), SUM( NT_CR2 ), SUM( NT_CR3 ), SUM( NT_CR_NO ), SUM( NL_CR1 ), SUM( NL_CR2 ), SUM( NL_CR3 ), SUM( NCL_CR1_NO ), SUM( NCL_CR2_NO ), SUM( NCL_CR1_SF ), SUM( NCL_CR2_SF ), SUM( NCL_J_SP_LF ), SUM(  NPUNCH_NO ), SUM( NPUNCH_SF ), SUM( NC_PAT1_SF ), SUM( NC_PAT2_SF ), SUM( NC_PAT3_SF ), AVG( NA_PAT_SF ), AVG( CDR ), SUM( CPR ), AVG( N_CCI ) ";
        }


        /* Function: Query()
         * Description: Will Format the Query that we want for transferring info from FWG database to our database
         */

        public static string Query(string FilePath)
        {
            return " INSERT INTO CRCP ( " + ItemsToGrab() + ") SELECT " + ItemsSumAvg() + " FROM [MS Access; DATABASE=" + FilePath + "].CRCP ";
        }


        /* Function: ExcelQuery()
         * Description: Will format the Query that we want for transferring info from our database to the Excel file
         */

        public static string ExcelQuery()
        {
            return "Insert into [CRCP$] ( Y2Y," + ItemsToGrab() + ") SELECT Y2Y, " + ItemsToGrab() + " From [MS Access; DATABASE=" + DistressTypes.GetRequiredDB() + "].CRCP";

        }
    };



    public class JRCP
    {
        JRCP()
        {

        }

        /* Function: CreateTableFormat()
         * Description: Will Create the headers for the JRCP Table and declare their types
         */

        public static string CreateTableFormat()
        {
            return " Y2Y Text, NT_CR1_NS NUMBER, NT_CR2_NS NUMBER, NL_CR1_NS NUMBER, NL_CR2_NS NUMBER,	NC_PAT1_NS NUMBER, NC_PAT2_NS NUMBER, NC_PAT3_NS NUMBER, NA_PAT_NS NUMBER, NNO_T_JTS NUMBER, NSLAB_AVG NUMBER, NT_J_SP_NS NUMBER, NL_J_SP_NS NUMBER, NC_BRK1_NS NUMBER, NC_BRK2_NS NUMBER, NDIV_NS NUMBER, NBLOW_NS NUMBER, SDR NUMBER  ";
        }

        /* Function: ItemsToGrab()
         * Description: Used to create the SQL Query for what information to get from the FWG Databases.
         */

        public static string ItemsToGrab()
        {
            return "NT_CR1_NS, NT_CR2_NS, NL_CR1_NS, NL_CR2_NS, NC_PAT1_NS, NC_PAT2_NS, NC_PAT3_NS, NA_PAT_NS, NNO_T_JTS, NSLAB_AVG, NT_J_SP_NS, NL_J_SP_NS, NC_BRK1_NS, NC_BRK2_NS, NDIV_NS, NBLOW_NS, SDR  ";

        }

        /* Function: ItemsSumAvg()
         * Description: Used to get the information that we want for each type in the FWG Database. Most of them
         *              are SUM of the Columns.
         */

        public static string ItemsSumAvg()
        {
            return "SUM(NT_CR1_NS), SUM(NT_CR2_NS), SUM(NL_CR1_NS), SUM(NL_CR2_NS), SUM(NC_PAT1_NS), SUM(NC_PAT2_NS), SUM(NC_PAT3_NS), SUM(NA_PAT_NS), SUM(NNO_T_JTS), AVG(NSLAB_AVG), SUM(NT_J_SP_NS), SUM(NL_J_SP_NS), SUM(NC_BRK1_NS), SUM(NC_BRK2_NS), SUM(NDIV_NS), SUM(NBLOW_NS), AVG(SDR) ";
        }

        /* Function: Query()
         * Description: Will Format the Query that we want for transferring info from FWG database to our database
         */

        public static string Query(string FilePath)
        { 
            return " INSERT INTO JRCP ( " + ItemsToGrab() + ") SELECT " + ItemsSumAvg() + " FROM [MS Access; DATABASE=" + FilePath + "].JRCP ";
        }

        /* Function: ExcelQuery()
         * Description: Will format the Query that we want for transferring info from our database to the Excel file
         */

        public static string ExcelQuery()
        {
            return "Insert into [JRCP$] ( Y2Y, " + ItemsToGrab() + ") SELECT Y2Y, " + ItemsToGrab() + " From [MS Access; DATABASE=" + DistressTypes.GetRequiredDB() + "].JRCP";

        }
    };
}

