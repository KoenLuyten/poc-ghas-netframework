using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Text = "Form1";

            // create new excel application through interop
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            // create new workbook
            Workbook workbook = excelApp.Workbooks.Add();
            // create new worksheet
            Worksheet worksheet = workbook.Worksheets.Add();
            // set cell value
            worksheet.Cells[1, 1] = "Hello World!";
            // save workbook
            workbook.SaveAs("C:\\Users\\user\\Desktop\\test.xlsx");
            // close workbook
            workbook.Close();
            // close excel application
            excelApp.Quit();
        }

        #endregion
    }
}

