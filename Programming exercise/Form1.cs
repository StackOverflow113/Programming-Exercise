using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;

namespace Programming_exercise
{

  public partial class Form1 : Form
  {
    public Form1()
    {
      InitializeComponent();
    }

    private void button1_Click(object sender, EventArgs e)
    {

      if (FileExplorer.ShowDialog() == DialogResult.OK)
      {
        txtOrigin.Text = FileExplorer.SelectedPath;
      }

    }

    private void button3_Click(object sender, EventArgs e)
    {
      Cursor.Current = Cursors.WaitCursor;
      if (txtOrigin.Text == "")
      {
        MessageBox.Show("First select the folder please...",
    "Error", MessageBoxButtons.OK,
    MessageBoxIcon.Exclamation);
        button1.Focus();
        Cursor.Current = Cursors.Default;
        return;
      }

      string folderPath = txtOrigin.Text;
      if (!Directory.Exists(folderPath + "\\Processed"))
      {
        Directory.CreateDirectory(txtOrigin.Text + "\\Processed");
        Directory.CreateDirectory(txtOrigin.Text + "\\Not applicable");
      }
      else
      {
        MessageBox.Show("The folder already exists",
     "Error", MessageBoxButtons.OK,
     MessageBoxIcon.Exclamation);
        Cursor.Current = Cursors.Default;
        return;
      }

      string[] dirs = Directory.GetFiles(txtOrigin.Text, "*.*");
      int cantidad = dirs.Length;
      int incrementable = 1;
      Excel.Application excel = new Excel.Application();

      Excel.Workbook masterbook = excel.Workbooks.Add();
      string routebook = txtOrigin.Text + "\\Processed\\";
      excel.ActiveWorkbook.SaveAs(routebook + "MASTERBOOK.xls", Excel.XlFileFormat.xlWorkbookNormal);

      foreach (string i in dirs)
      {
        string path = i;
        string filename = null;
        filename = Path.GetFileName(path);

        if (filename.EndsWith(".xls") || filename.EndsWith(".xlsx"))
        {

          string finalroute = txtOrigin.Text + "\\Processed\\" + filename;
          string destino = routebook + "MASTERBOOK.xls";

          File.Move(i, finalroute);

          Excel.Workbook wbSource = excel.Workbooks.Open(finalroute, 0, false, 1, "", "", false, Excel.XlPlatform.xlWindows, 9, true, false, 0, true, false, false);
          int cantHojasOrigin = wbSource.Sheets.Count;

          string pathFileDestination = destino;
          Excel.Workbook wbDest = excel.Workbooks.Open(pathFileDestination, 0, false, 1, "", "", false, Excel.XlPlatform.xlWindows, 9, true, false, 0, true, false, false);
          int cantHojas = wbDest.Sheets.Count;

          do
          {
            Excel.Worksheet WorksheetSource = wbSource.Sheets[incrementable];
            Excel.Worksheet WorksheetDestination = wbDest.Sheets[1];
            WorksheetSource.UsedRange.Copy(Missing.Value);
            WorksheetDestination.UsedRange.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, true);
            wbDest.Worksheets.Add();
            wbDest.Save();
            incrementable++;
          } while (incrementable <= cantHojasOrigin);


          wbSource.Close();
          excel.Quit();
          incrementable = 1;
        }

        else
        {
          string finalroute = txtOrigin.Text + "\\Not applicable\\" + filename;
          File.Move(i, finalroute);
        }

      }

      /*KILL THE ZOMBIE PROCESS*/
      KillSpecificExcelFileProcess("EXCEL");
      Cursor.Current = Cursors.WaitCursor;
      MessageBox.Show("A total of:" + " " + cantidad + " " + "files were processed.");
      txtOrigin.Text = string.Empty;
      button1.Focus();
    }

    private void KillSpecificExcelFileProcess(string excelFileName)
    {
      var processes = from p in Process.GetProcessesByName("EXCEL")
                      select p;

      foreach (var process in processes)
      {
        if (process.ProcessName == excelFileName)
        {
          process.Kill();
        }
      }
    }

  }
}

