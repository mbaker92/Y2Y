/* Author: Matthew Baker
 * Date Created: 03/17/2018
 * Date Modified: 03/17/2018
 * File: Form1.cs
 * Program: Y2Y Comparison
 * Description: This file controls the form used in this program. The form asks for the Current and
 *              the Previous FWG databases. It will then take the required information and create
 *              a database with the needed values. Then an Excel file is created and it is saved to
 *              the destination choosen by the user.
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Y2Y_Comparison
{
    public partial class Form1 : Form
    {

        // Declare Variables used in this class
        private Classes.DistressTypes Distress = new Classes.DistressTypes();
        public string FWGCurrent;
        public string FWGPrevious;
        private bool CurrentSelected = false;
        private bool PreviousSelected = false;

        public Form1()
        {
            InitializeComponent();

            // Set the text that shows the users choosen databases to blank
            PrevText.Text = "";
            CurrentText.Text = "";
        }


        /* Function : Browse1_Click
         * Description: Will get the filepath of the current database from the user
         */
            
        private void Browse1_Click(object sender, EventArgs e)
        {
            
            // Show OpenFileDialog to get the current database filename
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // set FWGCurrent to user specified file
                FWGCurrent = openFileDialog1.FileName;

                // Set CurrentSelected to true and show the filename on the form 
                CurrentSelected = true;
                CurrentText.Text = openFileDialog1.SafeFileName;
            }
            

        }

        /* Function : Browse2_Click
         * Description: Will get the filepath of the Previous database from the user
         */
        private void Browse2_Click(object sender, EventArgs e)
        {
            
            // Show OpenFileDialog to get the previous database filename
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                // set FWGPrevious to user specified file
                FWGPrevious = openFileDialog2.FileName;

                // Set PreviousSelected to true and show the filename on the form
                PreviousSelected = true;
                PrevText.Text = openFileDialog2.SafeFileName;
            }
        }

        /* Function : StartButton_Click
         * Description: If both PreviousSelected and CurrentSelected true, then get the required data from
         *              both databases and insert it into the new database. Format the tables as needed. 
         *              Export the info from the database into an excel. Reset the variables and show a 
         *              messagebox to tell user the file is ready.
         */

        private void StartButton_Click(object sender, EventArgs e)
        {

            if ((CurrentSelected == true) && (PreviousSelected == true))
            {
                // Insert required info from Current FWG database into the new database
                Insert(FWGCurrent);

                // Add "CURRENT" to the Y2Y column in the database for that Line
                ADDYear(0);

                // Insert required info from Previous FWG database into the new database
                Insert(FWGPrevious);

                // Add "PREVIOUS" to the Y2Y column in the Database for that Line
                ADDYear(1);

                // Format the tables as needed
                FormatTables();

                // Export the tables from the new Database into an Excel file
                ExportExcel();

                // Reset Variables for next run
                ResetVariables();

                // Alert User that Excel file is ready.
                MessageBox.Show(this, "Excel File is Ready for Viewing", "Finished", MessageBoxButtons.OK);
            }

        }

        /* Function: Insert(string DB)
         * Description: The Insert function will Insert the Required information into the ACP, CRCP, and JRCP
         *              tables from the filepath that is passed into the function.
         */

        private void Insert(string DB)
        {
            // Start DB connection to the new Database
            Distress.myConnection.Open();
            
            // Set SQL Command used to insert  ACP info from FWG database to new database and execute it
            OleDbCommand InsertACP = new OleDbCommand(Classes.ACP.Query(DB), Distress.myConnection);
            InsertACP.ExecuteNonQuery();

            // Set SQL Command used to insert  CRCP info from FWG database to new database and execute it
            OleDbCommand InsertCRCP = new OleDbCommand(Classes.CRCP.Query(DB), Distress.myConnection);
            InsertCRCP.ExecuteNonQuery();

            // Set SQL Command used to insert  JRCP info from FWG database to new database and execute it
            OleDbCommand InsertJRCP = new OleDbCommand(Classes.JRCP.Query(DB), Distress.myConnection);
            InsertJRCP.ExecuteNonQuery();

            // Close Database connection
            Distress.myConnection.Close();
            
        }

        /* Function: ADDYear(int current)
         * Description: This function will update the Y2Y column in the new database tables with current or previous.
         *              The current function should be run before the previous based on the SQL commands used. 
         */

        private void ADDYear(int current)
        {
            // Open Connection create command variable
            Distress.myConnection.Open();
            OleDbCommand Command = new OleDbCommand();
            Command.Connection = Distress.myConnection;

            // If current
            if (current == 0)
            {
                // Update the Y2Y for each line in every table to show CURRENT
                Command.CommandText = "UPDATE ACP SET Y2Y='CURRENT'";
                Command.ExecuteNonQuery();
                Command.CommandText = "UPDATE CRCP SET Y2Y='CURRENT'";
                Command.ExecuteNonQuery();
                Command.CommandText = "UPDATE JRCP SET Y2Y='CURRENT'";
                Command.ExecuteNonQuery();
            }
            else
            {
                // Update the Y2Y for Lines that are not set to current or where Y2Y is Null for every table
                Command.CommandText = "UPDATE ACP SET Y2Y='PREVIOUS' WHERE  Y2Y<>'CURRENT' OR Y2Y IS NULL ";
                Command.ExecuteNonQuery();
                Command.CommandText = "UPDATE CRCP SET Y2Y='PREVIOUS' WHERE Y2Y<>'CURRENT' OR Y2Y IS NULL ";
                Command.ExecuteNonQuery();
                Command.CommandText = "UPDATE JRCP SET Y2Y='PREVIOUS' WHERE Y2Y<>'CURRENT' OR Y2Y IS NULL ";
                Command.ExecuteNonQuery();
            }

            // Close Connection
            Distress.myConnection.Close();
        }

        /* Function: FormatTables()
         * Description: This will update a JRCP Variable to be divided by 100. Other Formatting needs can be
         *              put here.
         */

        private void FormatTables()
        {
            Distress.myConnection.Open();
            OleDbCommand Command = new OleDbCommand();
            Command.Connection = Distress.myConnection;
            Command.CommandText = "UPDATE JRCP SET NNO_T_JTS = NNO_T_JTS/100";
            Command.ExecuteNonQuery();
            Distress.myConnection.Close();
        }

        /* Function: ExportExcel()
         * Description: Export exccel will treat the Excel file as a database and copy the tables from new database to that
         *              Excel file. The SQL Query is done in the DistressTypes Class. It will then call OutputExcel() to prompt
         *              user for a folder to copy the Excel file to.
         */

        private void ExportExcel()
        {
            // Open connection to excel file
            Distress.ExcelConnection.Open();

            // Get ACP SQL Query and execute it
            OleDbCommand ACPExcelCommand = new OleDbCommand(Classes.ACP.ExcelQuery(), Distress.ExcelConnection);
            ACPExcelCommand.ExecuteNonQuery();

            // Get CRCP SQL Query and execute it
            OleDbCommand CRCPExcelCommand = new OleDbCommand(Classes.CRCP.ExcelQuery(), Distress.ExcelConnection);
            CRCPExcelCommand.ExecuteNonQuery();

            // Get JRCP SQL Query and execute it
            OleDbCommand JRCPExcelCommand = new OleDbCommand(Classes.JRCP.ExcelQuery(), Distress.ExcelConnection);
            JRCPExcelCommand.ExecuteNonQuery();

            // Close Connection to Excel File
            Distress.ExcelConnection.Close();

            // Prompt user for folder to save Excel
            OutputExcel();


        }

        /* Function: OutputExcel()
         * Description: Prompt user for a directory to store new excel file and copy it to the directory.
         *              If the File Already Exists, rename the oldfile with the time appended to the end.
         */

        private void OutputExcel()
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                if (System.IO.File.Exists(folderBrowserDialog1.SelectedPath + @"\Y2YOutput.xls"))
                {
                    System.IO.File.Copy(folderBrowserDialog1.SelectedPath + @"\Y2YOutput.xls", folderBrowserDialog1.SelectedPath + @"\Y2YOutput" + DateTime.Now.ToString("hhmmss") + ".xls");
                }
                System.IO.File.Copy(Distress.ExcelFile, folderBrowserDialog1.SelectedPath + @"\Y2YOutput.xls", true);

            }
        }

        /* Function: ResetVariables
         * Description: Reset all variables used to their initial states so that the program can be run again.
         */

        private void ResetVariables()
        {
            CurrentSelected = false;
            PreviousSelected = false;
            PrevText.Text = "";
            CurrentText.Text = "";

            // Reset the Excel File
            Distress.CopyExcelFileBack();

            // Reset Database
            Distress.DropTables();
            Distress.CreateDatabase();
        }
    }
}
