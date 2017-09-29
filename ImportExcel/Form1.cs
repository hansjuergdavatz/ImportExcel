using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;




namespace ImportExcel
{

  public partial class Form1 : Form
  {
    DataSet dsLocal = null;
    List<DBField> fieldList = new List<DBField>();

    string _connectionString = string.Empty;

    public Form1()
    {
      InitializeComponent();
      _connectionString = ConfigurationManager.ConnectionStrings["ImportConnectionString"].ConnectionString;
      SqlConnection con = new SqlConnection(_connectionString);
      this.Text = "Datenimport in DB: " + con.Database;
      con.Dispose();
    }

    private void btnFile_Click(object sender, EventArgs e)
    {
      OpenFileDialog openFileDialog1 = new OpenFileDialog();
      openFileDialog1.Filter = "Excel Files|*.xlsx";
      openFileDialog1.Title = "Excel-File auswählen";

      if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
      {
        txtImportFile.Text = openFileDialog1.FileName;
        dataGridView1.DataSource = null;
        Cursor.Current = Cursors.WaitCursor;
        Application.DoEvents();

        ReadExcel();
        Cursor.Current = Cursors.Default;
      }
    }
    private string GetExcelConnectionString(string fileName)
    {
      Dictionary<string, string> props = new Dictionary<string, string>();

      // XLSX - Excel 2007, 2010, 2012, 2013
      props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
      props["Extended Properties"] = "Excel 12.0 XML";
      props["Data Source"] = fileName; // "C:\\MyExcel.xlsx";

      // XLS - Excel 2003 and Older
      //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
      //props["Extended Properties"] = "Excel 8.0";
      //props["Data Source"] = "C:\\MyExcel.xls";

      StringBuilder sb = new StringBuilder();

      foreach (KeyValuePair<string, string> prop in props)
      {
        sb.Append(prop.Key);
        sb.Append('=');
        sb.Append(prop.Value);
        sb.Append(';');
      }

      return sb.ToString();
    }
    private void ReadExcel()
    {
      string connectionString = GetExcelConnectionString(txtImportFile.Text);

      dsLocal = new DataSet();
      try
      {
        using (OleDbConnection conn = new OleDbConnection(connectionString))
        {
          conn.Open();
          OleDbCommand cmd = new OleDbCommand();
          cmd.Connection = conn;

          // Get all Sheets in Excel File
          System.Data.DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

          // Loop through all Sheets to get data
          string sheetName = string.Empty;
          foreach (DataRow dr in dtSheet.Rows)
          {
            sheetName = dr["TABLE_NAME"].ToString();

            if (!sheetName.EndsWith("$"))
              continue;

            // Get all rows from the Sheet
            cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.TableName = sheetName;

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.Fill(dt);

            dsLocal.Tables.Add(dt);
          }

          cmd = null;
          conn.Close();
        }

        if (dsLocal.Tables.Count > 1)
        {
          MessageBox.Show("Das verwendete Importfile hat mehrere Tabellen und kann nicht verarbeitet werden.", "Import-Problem", MessageBoxButtons.OK, MessageBoxIcon.Error);
          return;
        }

      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
      }

      if (dsLocal.Tables.Count == 1)
      {
        var tables = dsLocal.Tables;
        dataGridView1.AutoGenerateColumns = true;
        dataGridView1.DataSource = dsLocal;
        dataGridView1.DataMember = tables[0].TableName;
      }
    }

    private System.Data.DataTable Read1<T>(string query) where T : IDbConnection, new()
    {
      using (var conn = new T())
      {
        using (var cmd = conn.CreateCommand())
        {
          cmd.CommandText = query;
          cmd.Connection.ConnectionString = _connectionString;
          cmd.Connection.Open();
          var table = new System.Data.DataTable();
          table.Load(cmd.ExecuteReader());
          return table;
        }
      }
    }
    private System.Data.DataTable ReadTableUpdate(string tableName, int idMaster)
    {
      System.Data.DataTable dt = new System.Data.DataTable();

      string keyField = "ID_Master";
      if (tableName != "PD_Master")
        keyField = "fID_Master";

      dt = Read1<SqlConnection>(string.Format("select * from {0} where {1} = {2}", tableName, keyField, idMaster));
      dt.TableName = tableName;
      return dt;
    }
    private System.Data.DataTable ReadTableInsert(string tableName)
    {
      System.Data.DataTable dt = new System.Data.DataTable();
      dt = Read1<SqlConnection>(string.Format("select top 0 * from {0}", tableName));
      dt.TableName = tableName;
      AddColumenToFieldList(dt);
      return dt;
    }
    private int GetMaxId(string tableName)
    {
      string sql = string.Empty;
      switch (tableName)
      {
        case "PD_MaDynAtt":
          sql = "select max(ID_MaDynAtt) FROM [dbo].[PD_MaDynAtt]";
          break;
        case "PD_MaMedia":
          sql = "select max(ID_MaMedia) FROM [dbo].[PD_MaMedia]";
          break;
        case "PD_Master":
          sql = "select max(ID_master) FROM [dbo].[PD_Master]";
          break;

        default:
          break;
      }

      SqlConnection connection = new SqlConnection(_connectionString);
      connection.Open();
      using (SqlCommand cmd1 = new SqlCommand(sql, connection))
      {
        object obj = cmd1.ExecuteScalar();
        if (obj.ToString().Length == 0)
          return 0;
        else
          return Convert.ToInt32(obj);
      }
    }
    private int GetIdMaster(string ArtNr)
    {
      string sql = string.Format("select ID_master FROM [dbo].[PD_Master] where [1000 - ArtNr] = '{0}'", ArtNr);
      SqlConnection connection = new SqlConnection(_connectionString);
      connection.Open();
      using (SqlCommand cmd1 = new SqlCommand(sql, connection))
      {
        if (cmd1.ExecuteScalar() != null)
        {
          int maxId = Convert.ToInt32(cmd1.ExecuteScalar());
          return maxId;
        }
        else
          return -1;
      }
    }
    private void AddColumenToFieldList(System.Data.DataTable table)
    {
      foreach (DataColumn column in table.Columns)
      {
        // Es wird die 4-stellige Nummer als Feldname gespeichert
        fieldList.Add(new DBField(column.ColumnName.Substring(0, 4), table.TableName, column.ColumnName, column.ReadOnly));
      }

    }
    private void CheckFields()
    {

    }


    private void InsertTableToSql(System.Data.DataTable table)
    {
      using (var bulkCopy = new SqlBulkCopy(_connectionString, SqlBulkCopyOptions.KeepIdentity))
      {
        // my DataTable column names match my SQL Column names, so I simply made this loop. However if your column names don't match, just pass in which datatable name matches the SQL column name in Column Mappings
        foreach (DataColumn col in table.Columns)
        {
          bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
        }
        bulkCopy.BulkCopyTimeout = 600;
        bulkCopy.DestinationTableName = table.TableName;
        bulkCopy.WriteToServer(table);
      }
    }
    private void UpdateTableToSql(System.Data.DataTable table, string tableName, int idMaster)
    {
      string sql = string.Empty;
      switch (tableName)
      {
        case "PD_Master":
          sql = string.Format("SELECT * FROM PD_Master Where ID_master = {1}", tableName, idMaster);
          break;
        case "PD_MaMedia":
          sql = string.Format("SELECT * FROM PD_MaMedia Where fID_master = {1}", tableName, idMaster);
          break;
        case "PD_MaDynAtt":
          sql = string.Format("SELECT * FROM PD_MaDynAtt Where fID_master = {1}", tableName, idMaster);
          break;
      }

      SqlConnection sqlConn = new SqlConnection(_connectionString);
      SqlDataAdapter adapter = new SqlDataAdapter(sql, sqlConn);
      adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey;   // needs to be set to apply table schema from db to datatable

      using (new SqlCommandBuilder(adapter))
      {
        try
        {
          System.Data.DataTable dtPrimary = new System.Data.DataTable();
          adapter.Fill(dtPrimary);
          foreach (DataRow row in table.Rows)
          {
            for (int i = 0; i < table.Columns.Count; i++)
            {
              string colName = Extensions.GetColumn(row, i);
              if (colName != "ID_master" && colName != "ID_MaMedia" && colName != "fID_Master" && colName != "ID_MaDynAtt")
                dtPrimary.Rows[0][colName] = row[i];
            }
          }
          sqlConn.Open();
          adapter.Update(dtPrimary);
          sqlConn.Close();
        }
        catch (Exception es)
        {
          Console.WriteLine(es.Message);
          Console.Read();
        }
      }
    }

    private int DeleteTableSQL(string tableName)
    {
      try
      {
        using (var sc = new SqlConnection(_connectionString))
        using (var cmd = sc.CreateCommand())
        {
          sc.Open();
          cmd.CommandText = string.Format("DELETE FROM {0}", tableName);
          int rc = cmd.ExecuteNonQuery();
          return rc;
        }
      }
      catch (Exception e)
      {
        MessageBox.Show(e.Message);
        return 0;
      }

    }
    private void btImport_Click(object sender, EventArgs e)
    {
      List<string> fehlendeFelder = new List<string>();
      int countErfasst = 0;
      int countGeaendert = 0;

      DataRow workRowMaster = null;
      DataRow workRowMedia = null;
      DataRow workRowAtt = null;

      System.Data.DataTable dtMasterUpdate = null;
      System.Data.DataTable dtMediaUpdate = null;
      System.Data.DataTable dtAttUpdate = null;

      try
      {
        System.Data.DataTable dtFromExcel = dsLocal.Tables[0]; // Excel-Daten

        fieldList.Clear();
        System.Data.DataTable dtMasterInsert = ReadTableInsert("PD_Master");
        System.Data.DataTable dtMediaInsert = ReadTableInsert("PD_MaMedia");
        System.Data.DataTable dtAttInsert = ReadTableInsert("PD_MaDynAtt");

        int maxIdMaster = GetMaxId("PD_Master");
        int workIdMaster = maxIdMaster;

        int maxIdAtt = GetMaxId("PD_MaDynAtt");
        int maxIdMedia = GetMaxId("PD_MaMedia");

        for (int i = 0; i < dtFromExcel.Columns.Count; i++)
        {
          string colName = Extensions.GetColumn(dtFromExcel.Rows[0], i);
          var value = fieldList.Find(item => item.Columen == colName.Substring(0, 4));
          if (value == null)
          {
            object exist = fehlendeFelder.Find(item => item == colName);
            if (exist == null)
              fehlendeFelder.Add(colName);
            break;
          }
        }

        if (fehlendeFelder.Count > 0)
        {
          string felder = string.Empty;
          foreach (var item in fehlendeFelder)
          {
            felder = felder + item + "\r\n";
          }

          string info = string.Format("Im Importfile sind unbekannte Felder:\r\n\r\n{0}\r\nDaten trotzdem importieren?.", felder);
          var rc = MessageBox.Show(info, "Unbekannte Felder", MessageBoxButtons.YesNo, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
          if (rc == DialogResult.No)
            return;
        }


        foreach (DataRow row in dtFromExcel.Rows)
        {
          bool hasDataMedia = false;
          bool hasDataAtt = false;

          // Kontrolle ob Datensatz vorhanden ist
          int existId = -1;
          for (int i = 0; i < dtFromExcel.Columns.Count; i++)
          {
            string colName = Extensions.GetColumn(row, i);
            if (colName == "1000 - ArtNr")
            {
              existId = GetIdMaster(row[i].ToString());
              break;
            }
          }

          if (existId > -1)
          {
            dtMasterUpdate = ReadTableUpdate("PD_Master", existId);
            workRowMaster = dtMasterUpdate.Select("ID_master=" + existId.ToString()).FirstOrDefault();

            dtMediaUpdate = ReadTableUpdate("PD_MaMedia", existId);
            workRowMedia = dtMediaUpdate.Select("fID_Master=" + existId.ToString()).FirstOrDefault();

            dtAttUpdate = ReadTableUpdate("PD_MaDynAtt", existId);
            workRowAtt = dtAttUpdate.Select("fID_Master=" + existId.ToString()).FirstOrDefault();
          }
          else
          {
            workRowMaster = dtMasterInsert.NewRow();
            workRowMedia = dtMediaInsert.NewRow();
            workRowAtt = dtAttInsert.NewRow();
          }

          for (int i = 0; i < dtFromExcel.Columns.Count; i++)
          {
            string colName = Extensions.GetColumn(row, i);
            var value = fieldList.Find(item => item.Columen == colName.Substring(0, 4));
            if (value == null)
            {
              //nur vorhandene Felder verarbeiten
              break;
            }

            if (colName != value.ColumenOrig)
              colName = value.ColumenOrig;

            switch (value.TableName)
            {
              case "PD_Master":
                if (colName == "1097 - Lagerware") // Ist das einige not Null Feld; Hat "-" als Inhalt; wenn nicht 1 dann false
                {
                  string s = row[i].ToString();
                  if (s == "1")
                    workRowMaster[colName] = true;
                  else
                    workRowMaster[colName] = false;
                }
                else
                  workRowMaster[colName] = row[i];
                break;

              case "PD_MaDynAtt":
                workRowAtt[colName] = row[i];
                hasDataAtt = true;
                break;
              case "PD_MaMedia":
                workRowMedia[colName] = row[i];
                hasDataMedia = true;
                break;
              default:
                break;
            }
          }

          if (existId == -1)
          {
            maxIdMaster++;
            workIdMaster = maxIdMaster;

            workRowMaster["ID_master"] = workIdMaster;
            dtMasterInsert.Rows.Add(workRowMaster);
            countErfasst++;

            if (hasDataAtt)
            {
              maxIdAtt++;
              workRowAtt["ID_MaDynAtt"] = maxIdAtt;
              workRowAtt["fID_Master"] = workIdMaster;
              dtAttInsert.Rows.Add(workRowAtt);
            }

            if (hasDataMedia)
            {
              maxIdMedia++;
              workRowMedia["ID_MaMedia"] = maxIdMedia;
              workRowMedia["fID_Master"] = workIdMaster;
              workRowMedia["IsMainVariantImages"] = false; //TODO woher??
              dtMediaInsert.Rows.Add(workRowMedia);
            }

          }
          else
          {
            workIdMaster = existId;
            UpdateTableToSql(dtMasterUpdate, "PD_Master", existId);
            countGeaendert++;

            if (hasDataAtt)
              UpdateTableToSql(dtAttUpdate, "PD_MaDynAtt", existId);

            if (hasDataMedia)
            {
              workRowMedia["IsMainVariantImages"] = false; //TODO woher??
              UpdateTableToSql(dtMediaUpdate, "PD_MaMedia", existId);
            }
          }
        }

        InsertTableToSql(dtMasterInsert);
        InsertTableToSql(dtMediaInsert);
        InsertTableToSql(dtAttInsert);

        //if (fehlendeFelder.Count > 0)
        //{
        //  string felder = string.Empty;
        //  foreach (var item in fehlendeFelder)
        //  {
        //    felder = felder + item + "\r\n";
        //  }

        //  string info = string.Format("Im Importfile sind unbekannte Felder:\r\n{0}\r\nBitte kontrollieren.", felder);
        //  MessageBox.Show(info, "Unbekannte Felder", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //}

        string info2 = string.Format("Erfasst: {0} Geändert:{1}", countErfasst, countGeaendert);
        MessageBox.Show(info2, "Import abgeschlossen", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
    }

    public bool IsNumeric(string input)
    {
      int test;
      return int.TryParse(input, out test);
    }

    private void button1_Click(object sender, EventArgs e)
    {
    }

    private void btnVorlage_Click(object sender, EventArgs e)
    {
      fieldList.Clear();
      System.Data.DataTable dtMasterInsert = ReadTableInsert("PD_Master");
      System.Data.DataTable dtMediaInsert = ReadTableInsert("PD_MaMedia");
      System.Data.DataTable dtAttInsert = ReadTableInsert("PD_MaDynAtt");

      System.Data.DataTable table = new System.Data.DataTable();
      table.TableName = "Produkte";
      foreach (var item in fieldList)
      {
        if (IsNumeric(item.ColumenOrig.Substring(0, 4)))
          table.Columns.Add(item.ColumenOrig, typeof(string));
      }

      XLWorkbook workbook = new XLWorkbook();
      workbook.Worksheets.Add(table);

      SaveFileDialog saveFileDialog1 = new SaveFileDialog();
      saveFileDialog1.Filter = "Excel file|*.xlsx";
      saveFileDialog1.Title = "Peicher ein Excel-File";
      saveFileDialog1.ShowDialog();
      if (saveFileDialog1.FileName != "")
      {
        workbook.SaveAs(saveFileDialog1.FileName);
      }

    }

    private void btnDeleteDB_Click(object sender, EventArgs e)
    {
      var rc = MessageBox.Show("Wollen Sie die Daten auf dem Server löschen?", "Löschen Daten", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
      if (rc == DialogResult.Yes)
      {
        int delCount = DeleteTableSQL("PD_Master");
        delCount = DeleteTableSQL("PD_MaMedia");
        delCount = DeleteTableSQL("PD_MaDynAtt");
        MessageBox.Show("Daten gelöscht.");
      }
    }
  }
}
