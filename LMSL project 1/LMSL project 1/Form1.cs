using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace LMSL_project_1
{
    public partial class form1 : Form
    {
        

        public form1()
        {
            InitializeComponent();
        }

        private void btncal_Click(object sender, EventArgs e)
        {
            // Calculate button click event code
            try
            {
                // Get input values
                double feedRate = Convert.ToDouble(txt_feedrate.Text);
                double operationHours = Convert.ToDouble(txt_OH.Text);
                double feedAssay = Convert.ToDouble(txt_feedassay.Text);
                double concentrateAssay = Convert.ToDouble(txt_conassay.Text);
                double tailingsAssay = Convert.ToDouble(txt_tailassay.Text);

                // Calculate Concentrate, Tailings, and Recovery
                double feedr = (feedRate * operationHours);
                double concentrate = (0.75 *feedr * (feedAssay - tailingsAssay))/(concentrateAssay - tailingsAssay);
                double tailings = (0.75*feedr) - concentrate;
                double recovery = (100 * concentrateAssay * (feedAssay - tailingsAssay))/(feedAssay*(concentrateAssay - tailingsAssay));
                
                

                // Display results
                txt_conout.Text = $"{concentrate:F2}";
                txt_tailout.Text = $"{tailings:F2}";
                txt_recovery.Text = $"{recovery:F2}";
                txt_feedout.Text = $"{feedr:F2}";
            }
            catch (FormatException)
            {
                MessageBox.Show("Please enter valid numeric values for all input fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnclearfld_Click(object sender, EventArgs e)
        {
            txt_feedrate.Clear();
            txt_OH.Clear();
            txt_feedassay.Clear();
            txt_conassay.Clear();
            txt_tailassay.Clear();
            txt_feedout.Clear();
            txt_conout.Clear();
            txt_tailout.Clear();
            txt_recovery.Clear();
            txtfeedravg.Clear();
            txtohavg.Clear();
            txtfeed_avg.Clear();
            txtcon_avg.Clear();
            txttail_avg.Clear();
            txtfeedasavg.Clear();
            txtconasavg.Clear();
            txttailasavg.Clear();
            txtrec_avg.Clear();
            txtcon_tot.Clear();
            txtfeed_tot.Clear();
            
            txtohtot.Clear();


        }

        private void btnexcel_Click(object sender, EventArgs e)
        {

            try
            {
                // Check if the file already exists
                string filePath = "D:\\LECTURES\\SEMISTER 6\\Industrial training\\vs project\\Book1.xlsx";
                bool fileExists = System.IO.File.Exists(filePath);

                // Create a new Excel application
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false; // Set Excel to be invisible

                // Open the existing workbook or add a new one
                Excel.Workbook workbook;

                if (fileExists)
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                }
                else
                {
                    workbook = excelApp.Workbooks.Add();
                }

                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                // Find the last used row in the worksheet
                Excel.Range foundRange = worksheet.Cells.Find("*", Type.Missing, Type.Missing, Type.Missing,
                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, Type.Missing, Type.Missing);

                int lastRow = (foundRange != null) ? foundRange.Row : 0;

                // Set headers if the worksheet is empty
                if (lastRow == 0)
                {
                    // ... (unchanged)

                    worksheet.Cells[1, 2] = "Date";
                    worksheet.Cells[1, 3] = "feed rate (T/h)";
                    worksheet.Cells[1, 4] = "OH (h)";
                    worksheet.Cells[1, 5] = "feed assay (%)";
                    worksheet.Cells[1, 6] = "con assay (%)";
                    worksheet.Cells[1, 7] = "tail assay (%)";
                    worksheet.Cells[1, 8] = "Feed (T)";
                    worksheet.Cells[1, 9] = "Concentrate (T)";
                    worksheet.Cells[1, 10] = "Tailings (T)";
                    worksheet.Cells[1, 11] = "Recovery (%)";
                    lastRow = 1; // Headers were added, set lastRow to 1
                }

                // Get output values
                DateTime currentDate;

                // Check if a valid date is provided in the txtdate textbox
                if (DateTime.TryParse(txtdate.Text, out currentDate))
                {
                    double feedRate1 = Convert.ToDouble(txt_feedrate.Text);
                    double oh1 = Convert.ToDouble(txt_OH.Text);
                    double feedas1 = Convert.ToDouble(txt_feedassay.Text);
                    double conas1 = Convert.ToDouble(txt_conassay.Text);
                    double tailas1 = Convert.ToDouble(txt_tailassay.Text);
                    double feedRateOut = Convert.ToDouble(txt_feedout.Text);
                    double concentrateOut = Convert.ToDouble(txt_conout.Text);
                    double tailingsOut = Convert.ToDouble(txt_tailout.Text);
                    double recoveryOut = Convert.ToDouble(txt_recovery.Text);

                    // Check if the date already exists in the worksheet
                    Excel.Range dateRange = worksheet.Cells.Find(currentDate.ToShortDateString(), Type.Missing, Type.Missing,
                        Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);

                    if (dateRange != null)
                    {
                        // Date already exists, ask for confirmation to replace
                        DialogResult result = MessageBox.Show("The data for this date already exists. Do you want to replace it?", "Data Exists",
                                                              MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            // Get the row number for the existing date
                            int existingRow = dateRange.Row;

                            // Replace existing data in the same row
                            worksheet.Cells[existingRow, 2] = currentDate.ToShortDateString();
                            worksheet.Cells[existingRow, 3] = feedRate1;
                            worksheet.Cells[existingRow, 4] = oh1;
                            worksheet.Cells[existingRow, 5] = feedas1;
                            worksheet.Cells[existingRow, 6] = conas1;
                            worksheet.Cells[existingRow, 7] = tailas1;
                            worksheet.Cells[existingRow, 8] = feedRateOut;
                            worksheet.Cells[existingRow, 9] = concentrateOut;
                            worksheet.Cells[existingRow, 10] = tailingsOut;
                            worksheet.Cells[existingRow, 11] = recoveryOut;

                            // Sort the worksheet by date
                            worksheet.UsedRange.Sort(worksheet.Columns[2, Type.Missing], Excel.XlSortOrder.xlAscending, Type.Missing, Type.Missing,
                                Excel.XlSortOrder.xlAscending, Type.Missing, Excel.XlSortOrder.xlAscending, Excel.XlYesNoGuess.xlGuess, Type.Missing, Type.Missing,
                                Excel.XlSortOrientation.xlSortColumns, Excel.XlSortMethod.xlPinYin, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal,
                                Excel.XlSortDataOption.xlSortNormal);

                            // Save the workbook
                            workbook.SaveAs(filePath);

                            // Close the workbook without saving changes (changes are already saved)
                            workbook.Close(false);

                            // Quit Excel application
                            excelApp.Quit();

                            // Release resources
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                            MessageBox.Show("Data exported to Excel successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            // User chose not to replace, exit without saving
                            return;
                        }
                    }
                    else
                    {
                        // Date does not exist, add a new row
                        lastRow++; // Increment lastRow to add a new record
                        worksheet.Cells[lastRow, 2] = currentDate.ToShortDateString();
                        worksheet.Cells[lastRow, 3] = feedRate1;
                        worksheet.Cells[lastRow, 4] = oh1;
                        worksheet.Cells[lastRow, 5] = feedas1;
                        worksheet.Cells[lastRow, 6] = conas1;
                        worksheet.Cells[lastRow, 7] = tailas1;
                        worksheet.Cells[lastRow, 8] = feedRateOut;
                        worksheet.Cells[lastRow, 9] = concentrateOut;
                        worksheet.Cells[lastRow, 10] = tailingsOut;
                        worksheet.Cells[lastRow, 11] = recoveryOut;

                        // Sort the worksheet by date
                        worksheet.UsedRange.Sort(worksheet.Columns[2, Type.Missing], Excel.XlSortOrder.xlAscending, Type.Missing, Type.Missing,
                            Excel.XlSortOrder.xlAscending, Type.Missing, Excel.XlSortOrder.xlAscending, Excel.XlYesNoGuess.xlGuess, Type.Missing, Type.Missing,
                            Excel.XlSortOrientation.xlSortColumns, Excel.XlSortMethod.xlPinYin, Excel.XlSortDataOption.xlSortNormal, Excel.XlSortDataOption.xlSortNormal,
                            Excel.XlSortDataOption.xlSortNormal);

                        // Save the workbook
                        workbook.SaveAs(filePath);

                        // Close the workbook without saving changes (changes are already saved)
                        workbook.Close(false);

                        // Quit Excel application
                        excelApp.Quit();

                        // Release resources
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                        MessageBox.Show("Data exported to Excel successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("enter valid date in the format MM/DD/YYYY.", "Invalid Date", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void btnget_Click(object sender, EventArgs e)
        {
            try
            {
                // Create a new Excel application
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false; // Make Excel visible

                // Open the existing workbook
                string filePath = "D:\\LECTURES\\SEMISTER 6\\Industrial training\\vs project\\Book1.xlsx";
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                // Find the last used row in the worksheet
                Excel.Range foundRange = worksheet.Cells.Find(txtdate.Text, Type.Missing, Type.Missing,
                    Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);

                int lastRow = (foundRange != null) ? foundRange.Row : 0;

                // Check if there are data for the specified date
                if (lastRow > 0)
                {
                    // Read data from Excel and display it in the form
                    double feedrate2 = Convert.ToDouble(worksheet.Cells[lastRow, 3].Value);
                    double oh2 = Convert.ToDouble(worksheet.Cells[lastRow, 4].Value);
                    double feedas2 = Convert.ToDouble(worksheet.Cells[lastRow, 5].Value);
                    double conas2 = Convert.ToDouble(worksheet.Cells[lastRow, 6].Value);
                    double tailas2 = Convert.ToDouble(worksheet.Cells[lastRow, 7].Value);
                    double feedR = Convert.ToDouble(worksheet.Cells[lastRow, 8].Value);
                    double concentrate = Convert.ToDouble(worksheet.Cells[lastRow, 9].Value);
                    double tailings = Convert.ToDouble(worksheet.Cells[lastRow, 10].Value);
                    double recovery = Convert.ToDouble(worksheet.Cells[lastRow, 11].Value);

                    // Display the data in your form
                    txt_feedrate.Text = $"{feedrate2:F2}";
                    txt_OH.Text = $"{oh2:F2}";
                    txt_feedassay.Text = $"{feedas2:F2}";
                    txt_conassay.Text = $"{conas2:F2}";
                    txt_tailassay.Text = $"{tailas2:F2}";
                    txt_feedout.Text = $"{feedR:F2}";
                    txt_conout.Text = $"{concentrate:F2}";
                    txt_tailout.Text = $"{tailings:F2}";
                    txt_recovery.Text = $"{recovery:F2}";
                }
                else
                {
                    MessageBox.Show("No data found for the specified date.", "Data Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                // Release resources
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        

        private void btn_cum_avg_Click(object sender, EventArgs e)
        {
            try
            {
                // Create a new Excel application
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false; // Make Excel invisible

                // Open the existing workbook
                string filePath = "D:\\LECTURES\\SEMISTER 6\\Industrial training\\vs project\\Book1.xlsx";
                Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                // Find the last used row in the worksheet
                Excel.Range lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                int lastRow = lastCell.Row;

                // Check if there are data in the worksheet
                if (lastRow > 0)
                {
                    DateTime startDate, endDate;

                    // Check if valid dates are provided in the txtfrom and txtto textboxes
                    if (DateTime.TryParse(txtfrom.Text, out startDate) && DateTime.TryParse(txtto.Text, out endDate))
                    {
                        // Initialize cumulative variables
                        double feedrateTotal = 0, ohTotal = 0, feedasTotal = 0, conasTotal = 0, tailasTotal = 0,
                               feedTotal = 0, conTotal = 0, tailTotal = 0, recTotal = 0;

                        // Initialize count variable for each category
                        int feedrateCount = 0, ohCount = 0, feedasCount = 0, conasCount = 0, tailasCount = 0,
                            feedCount = 0, conCount = 0, tailCount = 0, recCount = 0;

                        // Loop through each row within the specified date range and calculate cumulative values
                        for (int row = 2; row <= lastRow; row++) // Assuming the data starts from row 2 (headers are in row 1)
                        {
                            DateTime currentDate = Convert.ToDateTime(worksheet.Cells[row, 2].Value);

                            // Check if the current date is within the specified range
                            if (currentDate >= startDate && currentDate <= endDate)
                            {
                                feedrateTotal += Convert.ToDouble(worksheet.Cells[row, 3].Value);
                                ohTotal += Convert.ToDouble(worksheet.Cells[row, 4].Value);
                                feedasTotal += Convert.ToDouble(worksheet.Cells[row, 5].Value);
                                conasTotal += Convert.ToDouble(worksheet.Cells[row, 6].Value);
                                tailasTotal += Convert.ToDouble(worksheet.Cells[row, 7].Value);
                                feedTotal += Convert.ToDouble(worksheet.Cells[row, 8].Value);
                                conTotal += Convert.ToDouble(worksheet.Cells[row, 9].Value);
                                tailTotal += Convert.ToDouble(worksheet.Cells[row, 10].Value);
                                recTotal += Convert.ToDouble(worksheet.Cells[row, 11].Value);

                                // Increment the count for each category
                                feedrateCount++;
                                ohCount++;
                                feedasCount++;
                                conasCount++;
                                tailasCount++;
                                feedCount++;
                                conCount++;
                                tailCount++;
                                recCount++;
                            }
                        }

                        // Calculate averages
                        double avgFeedRate = feedrateCount > 0 ? feedrateTotal / feedrateCount : 0;
                        double avgOH = ohCount > 0 ? ohTotal / ohCount : 0;
                        double avgFeedAssay = feedasCount > 0 ? feedasTotal / feedasCount : 0;
                        double avgConAssay = conasCount > 0 ? conasTotal / conasCount : 0;
                        double avgTailAssay = tailasCount > 0 ? tailasTotal / tailasCount : 0;
                        double avgFeed = feedCount > 0 ? feedTotal / feedCount : 0;
                        double avgCon = conCount > 0 ? conTotal / conCount : 0;
                        double avgTail = tailCount > 0 ? tailTotal / tailCount : 0;
                        double avgRec = recCount > 0 ? recTotal / recCount : 0;

                        // Display cumulative values in respective textboxes
                        txtohtot.Text = $"{ohTotal:F2}";
                        txtfeed_tot.Text = $"{feedTotal:F2}";
                        txtcon_tot.Text = $"{conTotal:F2}";
                        txttail_tot.Text = $"{tailTotal:F2}";

                        // Display averages in respective textboxes
                        txtfeedravg.Text = $"{avgFeedRate:F2}";
                        txtohavg.Text = $"{avgOH:F2}";
                        txtfeedasavg.Text = $"{avgFeedAssay:F2}";
                        txtconasavg.Text = $"{avgConAssay:F2}";
                        txttailasavg.Text = $"{avgTailAssay:F2}";
                        txtfeed_avg.Text = $"{avgFeed:F2}";
                        txtcon_avg.Text = $"{avgCon:F2}";
                        txttail_avg.Text = $"{avgTail:F2}";
                        txtrec_avg.Text = $"{avgRec:F2}";
                    }
                    else
                    {
                        MessageBox.Show("Please enter valid dates in the format MM/DD/YYYY.", "Invalid Date", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No data found in the Excel file.", "No Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                // Release resources
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        
        

        
    }
}
