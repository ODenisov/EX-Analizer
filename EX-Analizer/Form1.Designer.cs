namespace EX_Analizer
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.b_openInput = new System.Windows.Forms.Button();
            this.b_startProcess = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.openExcelFile = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // b_openInput
            // 
            this.b_openInput.Location = new System.Drawing.Point(108, 12);
            this.b_openInput.Name = "b_openInput";
            this.b_openInput.Size = new System.Drawing.Size(75, 23);
            this.b_openInput.TabIndex = 0;
            this.b_openInput.Text = "Вручную";
            this.b_openInput.UseVisualStyleBackColor = true;
            this.b_openInput.Click += new System.EventHandler(this.b_openInput_Click);
            // 
            // b_startProcess
            // 
            this.b_startProcess.Location = new System.Drawing.Point(108, 91);
            this.b_startProcess.Name = "b_startProcess";
            this.b_startProcess.Size = new System.Drawing.Size(75, 23);
            this.b_startProcess.TabIndex = 1;
            this.b_startProcess.Text = "Старт";
            this.b_startProcess.UseVisualStyleBackColor = true;
            this.b_startProcess.Click += new System.EventHandler(this.b_startProcess_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(129, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "label1";
            // 
            // openExcelFile
            // 
            this.openExcelFile.FileName = "openFileDialog1";
            this.openExcelFile.Multiselect = true;
            this.openExcelFile.FileOk += new System.ComponentModel.CancelEventHandler(this.openExcelFile_FileOk);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(289, 128);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.b_startProcess);
            this.Controls.Add(this.b_openInput);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button b_openInput;
        private System.Windows.Forms.Button b_startProcess;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openExcelFile;
    }
}

