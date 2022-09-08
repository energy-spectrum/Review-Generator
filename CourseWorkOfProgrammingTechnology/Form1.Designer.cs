
namespace CourseWorkOfProgrammingTechnology
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
            this.buttonOk = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.pathInputDataFile = new System.Windows.Forms.Label();
            this.SelectingInputDataFile = new System.Windows.Forms.Button();
            this.SelectingFolderToSave = new System.Windows.Forms.Button();
            this.pathFolderToSave = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SelectingTeacherSignature = new System.Windows.Forms.Button();
            this.pathToTeacherSignature = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonOk
            // 
            this.buttonOk.Location = new System.Drawing.Point(342, 358);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(94, 29);
            this.buttonOk.TabIndex = 0;
            this.buttonOk.Text = "Ok";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.Ok_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialogInputData";
            // 
            // pathInputDataFile
            // 
            this.pathInputDataFile.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.pathInputDataFile.Image = global::CourseWorkOfProgrammingTechnology.Properties.Resources.BackgroundForPath;
            this.pathInputDataFile.Location = new System.Drawing.Point(126, 52);
            this.pathInputDataFile.Name = "pathInputDataFile";
            this.pathInputDataFile.Size = new System.Drawing.Size(627, 25);
            this.pathInputDataFile.TabIndex = 1;
            // 
            // SelectingInputDataFile
            // 
            this.SelectingInputDataFile.Location = new System.Drawing.Point(26, 36);
            this.SelectingInputDataFile.Name = "SelectingInputDataFile";
            this.SelectingInputDataFile.Size = new System.Drawing.Size(94, 53);
            this.SelectingInputDataFile.TabIndex = 2;
            this.SelectingInputDataFile.Text = "Выберите файл";
            this.SelectingInputDataFile.UseVisualStyleBackColor = true;
            this.SelectingInputDataFile.Click += new System.EventHandler(this.SelectingInputDataFile_Click);
            // 
            // SelectingFolderToSave
            // 
            this.SelectingFolderToSave.Location = new System.Drawing.Point(26, 183);
            this.SelectingFolderToSave.Name = "SelectingFolderToSave";
            this.SelectingFolderToSave.Size = new System.Drawing.Size(94, 53);
            this.SelectingFolderToSave.TabIndex = 4;
            this.SelectingFolderToSave.Text = "Выберите папку";
            this.SelectingFolderToSave.UseVisualStyleBackColor = true;
            this.SelectingFolderToSave.Click += new System.EventHandler(this.SelectingFolderToSave_Click);
            // 
            // pathFolderToSave
            // 
            this.pathFolderToSave.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.pathFolderToSave.Image = global::CourseWorkOfProgrammingTechnology.Properties.Resources.BackgroundForPath;
            this.pathFolderToSave.Location = new System.Drawing.Point(126, 195);
            this.pathFolderToSave.Name = "pathFolderToSave";
            this.pathFolderToSave.Size = new System.Drawing.Size(627, 25);
            this.pathFolderToSave.TabIndex = 3;
            // 
            // SelectingTeacherSignature
            // 
            this.SelectingTeacherSignature.Location = new System.Drawing.Point(26, 110);
            this.SelectingTeacherSignature.Name = "SelectingTeacherSignature";
            this.SelectingTeacherSignature.Size = new System.Drawing.Size(94, 53);
            this.SelectingTeacherSignature.TabIndex = 6;
            this.SelectingTeacherSignature.Text = "Выберите подпись";
            this.SelectingTeacherSignature.UseVisualStyleBackColor = true;
            this.SelectingTeacherSignature.Click += new System.EventHandler(this.SelectingTeacherSignature_Click);
            // 
            // pathToTeacherSignature
            // 
            this.pathToTeacherSignature.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.pathToTeacherSignature.Image = global::CourseWorkOfProgrammingTechnology.Properties.Resources.BackgroundForPath;
            this.pathToTeacherSignature.Location = new System.Drawing.Point(126, 122);
            this.pathToTeacherSignature.Name = "pathToTeacherSignature";
            this.pathToTeacherSignature.Size = new System.Drawing.Size(627, 25);
            this.pathToTeacherSignature.TabIndex = 5;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.SelectingTeacherSignature);
            this.Controls.Add(this.pathToTeacherSignature);
            this.Controls.Add(this.SelectingFolderToSave);
            this.Controls.Add(this.pathFolderToSave);
            this.Controls.Add(this.SelectingInputDataFile);
            this.Controls.Add(this.pathInputDataFile);
            this.Controls.Add(this.buttonOk);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Review Generator";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label pathInputDataFile;
        private System.Windows.Forms.Button SelectingInputDataFile;
        private System.Windows.Forms.Button SelectingFolderToSave;
        private System.Windows.Forms.Label pathFolderToSave;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button SelectingTeacherSignature;
        private System.Windows.Forms.Label pathToTeacherSignature;
    }
}

