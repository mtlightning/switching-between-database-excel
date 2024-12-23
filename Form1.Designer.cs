namespace ExcelVTEntegrasyonProjesi
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
            buttonVTdenOku = new Button();
            richTextBox1 = new RichTextBox();
            richTextBox2 = new RichTextBox();
            buttonExceldenOku = new Button();
            SuspendLayout();
            // 
            // buttonVTdenOku
            // 
            buttonVTdenOku.BackColor = SystemColors.AppWorkspace;
            buttonVTdenOku.Location = new Point(453, 60);
            buttonVTdenOku.Name = "buttonVTdenOku";
            buttonVTdenOku.Size = new Size(173, 55);
            buttonVTdenOku.TabIndex = 0;
            buttonVTdenOku.Text = "Veri Tabanından Oku ve Excel'e Yaz";
            buttonVTdenOku.UseVisualStyleBackColor = false;
            buttonVTdenOku.Click += buttonVTdenOku_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(12, 32);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(390, 120);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // richTextBox2
            // 
            richTextBox2.Location = new Point(12, 170);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(390, 120);
            richTextBox2.TabIndex = 2;
            richTextBox2.Text = "";
            // 
            // buttonExceldenOku
            // 
            buttonExceldenOku.BackColor = SystemColors.AppWorkspace;
            buttonExceldenOku.Location = new Point(453, 211);
            buttonExceldenOku.Name = "buttonExceldenOku";
            buttonExceldenOku.Size = new Size(173, 55);
            buttonExceldenOku.TabIndex = 3;
            buttonExceldenOku.Text = "Excel'den Oku ve Veritabanına Yaz";
            buttonExceldenOku.UseVisualStyleBackColor = false;
            buttonExceldenOku.Click += buttonExceldenOku_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(buttonExceldenOku);
            Controls.Add(richTextBox2);
            Controls.Add(richTextBox1);
            Controls.Add(buttonVTdenOku);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
        }

        #endregion

        private Button buttonVTdenOku;
        private RichTextBox richTextBox1;
        private RichTextBox richTextBox2;
        private Button buttonExceldenOku;
    }
}
