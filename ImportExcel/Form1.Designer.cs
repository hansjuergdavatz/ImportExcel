namespace ImportExcel
{
  partial class Form1
  {
    /// <summary>
    /// Erforderliche Designervariable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Verwendete Ressourcen bereinigen.
    /// </summary>
    /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
    protected override void Dispose(bool disposing)
    {
      if (disposing && (components != null))
      {
        components.Dispose();
      }
      base.Dispose(disposing);
    }

    #region Vom Windows Form-Designer generierter Code

    /// <summary>
    /// Erforderliche Methode für die Designerunterstützung.
    /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
    /// </summary>
    private void InitializeComponent()
    {
      this.btnFile = new System.Windows.Forms.Button();
      this.txtImportFile = new System.Windows.Forms.TextBox();
      this.btImport = new System.Windows.Forms.Button();
      this.dataGridView1 = new System.Windows.Forms.DataGridView();
      this.btnVorlage = new System.Windows.Forms.Button();
      this.btnDeleteDB = new System.Windows.Forms.Button();
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
      this.SuspendLayout();
      // 
      // btnFile
      // 
      this.btnFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
      this.btnFile.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.btnFile.Location = new System.Drawing.Point(547, 24);
      this.btnFile.Name = "btnFile";
      this.btnFile.Size = new System.Drawing.Size(32, 23);
      this.btnFile.TabIndex = 0;
      this.btnFile.Text = "...";
      this.btnFile.UseVisualStyleBackColor = true;
      this.btnFile.Click += new System.EventHandler(this.btnFile_Click);
      // 
      // txtImportFile
      // 
      this.txtImportFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.txtImportFile.Location = new System.Drawing.Point(12, 24);
      this.txtImportFile.Name = "txtImportFile";
      this.txtImportFile.Size = new System.Drawing.Size(529, 20);
      this.txtImportFile.TabIndex = 1;
      // 
      // btImport
      // 
      this.btImport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.btImport.Location = new System.Drawing.Point(466, 209);
      this.btImport.Name = "btImport";
      this.btImport.Size = new System.Drawing.Size(112, 23);
      this.btImport.TabIndex = 2;
      this.btImport.Text = "Import in die DB";
      this.btImport.UseVisualStyleBackColor = true;
      this.btImport.Click += new System.EventHandler(this.btImport_Click);
      // 
      // dataGridView1
      // 
      this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridView1.Location = new System.Drawing.Point(12, 50);
      this.dataGridView1.Name = "dataGridView1";
      this.dataGridView1.Size = new System.Drawing.Size(567, 147);
      this.dataGridView1.TabIndex = 5;
      // 
      // btnVorlage
      // 
      this.btnVorlage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.btnVorlage.Location = new System.Drawing.Point(12, 210);
      this.btnVorlage.Name = "btnVorlage";
      this.btnVorlage.Size = new System.Drawing.Size(146, 23);
      this.btnVorlage.TabIndex = 6;
      this.btnVorlage.Text = "Import Vorlage erstellen";
      this.btnVorlage.UseVisualStyleBackColor = true;
      this.btnVorlage.Click += new System.EventHandler(this.btnVorlage_Click);
      // 
      // btnDeleteDB
      // 
      this.btnDeleteDB.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
      this.btnDeleteDB.Location = new System.Drawing.Point(256, 210);
      this.btnDeleteDB.Name = "btnDeleteDB";
      this.btnDeleteDB.Size = new System.Drawing.Size(112, 23);
      this.btnDeleteDB.TabIndex = 7;
      this.btnDeleteDB.Text = "DB Daten löschen";
      this.btnDeleteDB.UseVisualStyleBackColor = true;
      this.btnDeleteDB.Click += new System.EventHandler(this.btnDeleteDB_Click);
      // 
      // Form1
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(594, 245);
      this.Controls.Add(this.btnDeleteDB);
      this.Controls.Add(this.btnVorlage);
      this.Controls.Add(this.dataGridView1);
      this.Controls.Add(this.btImport);
      this.Controls.Add(this.txtImportFile);
      this.Controls.Add(this.btnFile);
      this.Name = "Form1";
      this.Text = "Import Daten";
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

    }

    #endregion

    private System.Windows.Forms.Button btnFile;
    private System.Windows.Forms.TextBox txtImportFile;
    private System.Windows.Forms.Button btImport;
    private System.Windows.Forms.DataGridView dataGridView1;
    private System.Windows.Forms.Button btnVorlage;
    private System.Windows.Forms.Button btnDeleteDB;
  }
}

