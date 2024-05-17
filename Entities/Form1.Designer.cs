using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;


namespace ExcelImageDownloader
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnSelectFile = new Button();
            txtFilePath = new TextBox();
            txtFolderPath = new TextBox();
            btnSelectFolder = new Button();
            btnDownload = new Button();
            progressBar = new ProgressBar();
            lblStatus = new Label();
            SuspendLayout();
            // 
            // btnSelectFile
            // 
            btnSelectFile.Location = new Point(12, 12);
            btnSelectFile.Name = "btnSelectFile";
            btnSelectFile.Size = new Size(100, 23);
            btnSelectFile.TabIndex = 0;
            btnSelectFile.Text = "Select File";
            btnSelectFile.UseVisualStyleBackColor = true;
            btnSelectFile.Click += btnSelectFile_Click;
            // 
            // txtFilePath
            // 
            txtFilePath.Location = new Point(118, 14);
            txtFilePath.Name = "txtFilePath";
            txtFilePath.ReadOnly = true;
            txtFilePath.Size = new Size(270, 23);
            txtFilePath.TabIndex = 1;
            // 
            // txtFolderPath
            // 
            txtFolderPath.Location = new Point(118, 43);
            txtFolderPath.Name = "txtFolderPath";
            txtFolderPath.ReadOnly = true;
            txtFolderPath.Size = new Size(270, 23);
            txtFolderPath.TabIndex = 2;
            // 
            // btnSelectFolder
            // 
            btnSelectFolder.Location = new Point(12, 41);
            btnSelectFolder.Name = "btnSelectFolder";
            btnSelectFolder.Size = new Size(100, 23);
            btnSelectFolder.TabIndex = 3;
            btnSelectFolder.Text = "Select folder";
            btnSelectFolder.UseVisualStyleBackColor = true;
            btnSelectFolder.Click += btnSelectFolder_Click;
            // 
            // btnDownload
            // 
            btnDownload.Location = new Point(12, 70);
            btnDownload.Name = "btnDownload";
            btnDownload.Size = new Size(100, 23);
            btnDownload.TabIndex = 4;
            btnDownload.Text = "Download";
            btnDownload.UseVisualStyleBackColor = true;
            btnDownload.Click += btnDownload_Click;
            // 
            // progressBar
            // 
            progressBar.Location = new Point(12, 99);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(376, 23);
            progressBar.TabIndex = 5;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(12, 125);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(39, 15);
            lblStatus.TabIndex = 6;
            lblStatus.Text = "Status";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(450, 223);
            Controls.Add(lblStatus);
            Controls.Add(progressBar);
            Controls.Add(btnDownload);
            Controls.Add(btnSelectFolder);
            Controls.Add(txtFolderPath);
            Controls.Add(txtFilePath);
            Controls.Add(btnSelectFile);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form1";
            Text = "ExcelSystemDowload";
            TopMost = true;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnSelectFile;
        private TextBox txtFilePath;
        private TextBox txtFolderPath;
        private Button btnSelectFolder;
        private Button btnDownload;
        private ProgressBar progressBar;
        private Label lblStatus;

        // 
        // btnSelectFile
        // 


    }
}
