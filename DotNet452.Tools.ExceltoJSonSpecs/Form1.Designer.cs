namespace ExceltoJSonSpecs
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
            this.UploadSpec = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.InsertColor = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // UploadSpec
            // 
            this.UploadSpec.Location = new System.Drawing.Point(50, 71);
            this.UploadSpec.Name = "UploadSpec";
            this.UploadSpec.Size = new System.Drawing.Size(196, 86);
            this.UploadSpec.TabIndex = 0;
            this.UploadSpec.Text = "ExceltoJson Updateby Name";
            this.UploadSpec.UseVisualStyleBackColor = true;
            this.UploadSpec.Click += new System.EventHandler(this.button1_Click);
            // 
            // InsertColor
            // 
            this.InsertColor.Location = new System.Drawing.Point(300, 71);
            this.InsertColor.Name = "InsertColor";
            this.InsertColor.Size = new System.Drawing.Size(196, 86);
            this.InsertColor.TabIndex = 0;
            this.InsertColor.Text = "Insert Storage and  Color";
            this.InsertColor.UseVisualStyleBackColor = true;
            this.InsertColor.Click += new System.EventHandler(this.InsertColor_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(824, 428);
            this.Controls.Add(this.InsertColor);
            this.Controls.Add(this.UploadSpec);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button UploadSpec;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button InsertColor;
    }
}

