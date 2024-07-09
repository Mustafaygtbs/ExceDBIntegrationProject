namespace ExceDBIntegrationProject
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            richTextBox1 = new RichTextBox();
            btnVTdenOku = new Button();
            btnExceldenOku = new Button();
            richTextBox2 = new RichTextBox();
            SuspendLayout();
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(35, 234);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(427, 177);
            richTextBox1.TabIndex = 0;
            richTextBox1.Text = "";
            // 
            // btnVTdenOku
            // 
            btnVTdenOku.Location = new Point(529, 291);
            btnVTdenOku.Name = "btnVTdenOku";
            btnVTdenOku.Size = new Size(248, 73);
            btnVTdenOku.TabIndex = 1;
            btnVTdenOku.Text = "read from database and write to excel";
            btnVTdenOku.UseVisualStyleBackColor = true;
            btnVTdenOku.Click += btnVTdenOku_Click;
            // 
            // btnExceldenOku
            // 
            btnExceldenOku.Location = new Point(529, 56);
            btnExceldenOku.Name = "btnExceldenOku";
            btnExceldenOku.Size = new Size(248, 70);
            btnExceldenOku.TabIndex = 2;
            btnExceldenOku.Text = "read from excel and write to database ";
            btnExceldenOku.UseVisualStyleBackColor = true;
            btnExceldenOku.Click += btnExceldenOku_Click;
            // 
            // richTextBox2
            // 
            richTextBox2.BackColor = Color.White;
            richTextBox2.Location = new Point(35, 12);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(427, 170);
            richTextBox2.TabIndex = 3;
            richTextBox2.Text = "";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.Red;
            ClientSize = new Size(857, 504);
            Controls.Add(richTextBox2);
            Controls.Add(btnExceldenOku);
            Controls.Add(btnVTdenOku);
            Controls.Add(richTextBox1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
        }

        #endregion

        private RichTextBox richTextBox1;
        private Button btnVTdenOku;
        private Button btnExceldenOku;
        private RichTextBox richTextBox2;
    }
}
