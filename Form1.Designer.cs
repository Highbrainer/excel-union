
namespace ExcelUnion
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.file1Button = new System.Windows.Forms.Button();
            this.file1TextBox = new System.Windows.Forms.TextBox();
            this.buttonLaunch = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBarA = new System.Windows.Forms.ProgressBar();
            this.keyColumn1ComboBox = new System.Windows.Forms.ComboBox();
            this.sheet1ComboBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.chosenCols1ListBox = new System.Windows.Forms.CheckedListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.chosenCols2ListBox = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.sheet2ComboBox = new System.Windows.Forms.ComboBox();
            this.keyColumn2ComboBox = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.file2TextBox = new System.Windows.Forms.TextBox();
            this.file2Button = new System.Windows.Forms.Button();
            this.saveFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // file1Button
            // 
            this.file1Button.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.file1Button.Location = new System.Drawing.Point(448, 22);
            this.file1Button.Name = "file1Button";
            this.file1Button.Size = new System.Drawing.Size(34, 23);
            this.file1Button.TabIndex = 0;
            this.file1Button.Text = "...";
            this.file1Button.UseVisualStyleBackColor = true;
            this.file1Button.Click += new System.EventHandler(this.file1Button_Click);
            // 
            // file1TextBox
            // 
            this.file1TextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.file1TextBox.Location = new System.Drawing.Point(6, 22);
            this.file1TextBox.Name = "file1TextBox";
            this.file1TextBox.Size = new System.Drawing.Size(436, 23);
            this.file1TextBox.TabIndex = 1;
            this.file1TextBox.TextChanged += new System.EventHandler(this.file1TextBox_TextChanged);
            // 
            // buttonLaunch
            // 
            this.buttonLaunch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonLaunch.Location = new System.Drawing.Point(933, 389);
            this.buttonLaunch.Name = "buttonLaunch";
            this.buttonLaunch.Size = new System.Drawing.Size(90, 26);
            this.buttonLaunch.TabIndex = 5;
            this.buttonLaunch.Text = "Go !";
            this.buttonLaunch.UseVisualStyleBackColor = true;
            this.buttonLaunch.Click += new System.EventHandler(this.buttonLaunch_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 57);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 15);
            this.label2.TabIndex = 8;
            this.label2.Text = "Onglet";
            // 
            // progressBarA
            // 
            this.progressBarA.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBarA.Location = new System.Drawing.Point(9, 389);
            this.progressBarA.Name = "progressBarA";
            this.progressBarA.Size = new System.Drawing.Size(916, 26);
            this.progressBarA.TabIndex = 9;
            // 
            // keyColumn1ComboBox
            // 
            this.keyColumn1ComboBox.FormattingEnabled = true;
            this.keyColumn1ComboBox.Location = new System.Drawing.Point(262, 54);
            this.keyColumn1ComboBox.Name = "keyColumn1ComboBox";
            this.keyColumn1ComboBox.Size = new System.Drawing.Size(180, 23);
            this.keyColumn1ComboBox.TabIndex = 11;
            // 
            // sheet1ComboBox
            // 
            this.sheet1ComboBox.FormattingEnabled = true;
            this.sheet1ComboBox.Location = new System.Drawing.Point(55, 54);
            this.sheet1ComboBox.Name = "sheet1ComboBox";
            this.sheet1ComboBox.Size = new System.Drawing.Size(108, 23);
            this.sheet1ComboBox.TabIndex = 12;
            this.sheet1ComboBox.TextChanged += new System.EventHandler(this.sheet1ComboBox_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(182, 57);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 15);
            this.label3.TabIndex = 13;
            this.label3.Text = "Colonne clef";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.chosenCols1ListBox);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.sheet1ComboBox);
            this.groupBox1.Controls.Add(this.keyColumn1ComboBox);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.file1TextBox);
            this.groupBox1.Controls.Add(this.file1Button);
            this.groupBox1.Location = new System.Drawing.Point(9, 8);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(495, 372);
            this.groupBox1.TabIndex = 14;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Fichier 1";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 94);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(108, 15);
            this.label5.TabIndex = 15;
            this.label5.Text = "Colonnes choisies :";
            // 
            // chosenCols1ListBox
            // 
            this.chosenCols1ListBox.FormattingEnabled = true;
            this.chosenCols1ListBox.Location = new System.Drawing.Point(6, 111);
            this.chosenCols1ListBox.Name = "chosenCols1ListBox";
            this.chosenCols1ListBox.Size = new System.Drawing.Size(436, 238);
            this.chosenCols1ListBox.TabIndex = 14;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.chosenCols2ListBox);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.sheet2ComboBox);
            this.groupBox2.Controls.Add(this.keyColumn2ComboBox);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.file2TextBox);
            this.groupBox2.Controls.Add(this.file2Button);
            this.groupBox2.Location = new System.Drawing.Point(522, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(501, 368);
            this.groupBox2.TabIndex = 15;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Fichier 2";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 89);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(108, 15);
            this.label6.TabIndex = 16;
            this.label6.Text = "Colonnes choisies :";
            // 
            // chosenCols2ListBox
            // 
            this.chosenCols2ListBox.FormattingEnabled = true;
            this.chosenCols2ListBox.Location = new System.Drawing.Point(6, 107);
            this.chosenCols2ListBox.Name = "chosenCols2ListBox";
            this.chosenCols2ListBox.Size = new System.Drawing.Size(442, 238);
            this.chosenCols2ListBox.TabIndex = 15;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(182, 57);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(74, 15);
            this.label1.TabIndex = 13;
            this.label1.Text = "Colonne clef";
            // 
            // sheet2ComboBox
            // 
            this.sheet2ComboBox.FormattingEnabled = true;
            this.sheet2ComboBox.Location = new System.Drawing.Point(55, 54);
            this.sheet2ComboBox.Name = "sheet2ComboBox";
            this.sheet2ComboBox.Size = new System.Drawing.Size(108, 23);
            this.sheet2ComboBox.TabIndex = 12;
            this.sheet2ComboBox.TextChanged += new System.EventHandler(this.sheet2ComboBox_TextChanged);
            // 
            // keyColumn2ComboBox
            // 
            this.keyColumn2ComboBox.FormattingEnabled = true;
            this.keyColumn2ComboBox.Location = new System.Drawing.Point(262, 54);
            this.keyColumn2ComboBox.Name = "keyColumn2ComboBox";
            this.keyColumn2ComboBox.Size = new System.Drawing.Size(186, 23);
            this.keyColumn2ComboBox.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 57);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(43, 15);
            this.label4.TabIndex = 8;
            this.label4.Text = "Onglet";
            // 
            // file2TextBox
            // 
            this.file2TextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.file2TextBox.Location = new System.Drawing.Point(6, 22);
            this.file2TextBox.Name = "file2TextBox";
            this.file2TextBox.Size = new System.Drawing.Size(442, 23);
            this.file2TextBox.TabIndex = 1;
            this.file2TextBox.TextChanged += new System.EventHandler(this.file2TextBox_TextChanged);
            // 
            // file2Button
            // 
            this.file2Button.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.file2Button.Location = new System.Drawing.Point(454, 22);
            this.file2Button.Name = "file2Button";
            this.file2Button.Size = new System.Drawing.Size(34, 23);
            this.file2Button.TabIndex = 0;
            this.file2Button.Text = "...";
            this.file2Button.UseVisualStyleBackColor = true;
            this.file2Button.Click += new System.EventHandler(this.file2Button_Click);
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.CheckFileExists = false;
            this.saveFileDialog.DefaultExt = "xlsx";
            this.saveFileDialog.FileName = "export-union.xlsx";
            this.saveFileDialog.Title = "Fichier à générer";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1035, 427);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.progressBarA);
            this.Controls.Add(this.buttonLaunch);
            this.Name = "Form1";
            this.Text = "Onglets à unir";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button file1Button;
        private System.Windows.Forms.TextBox file1TextBox;
        private System.Windows.Forms.Button buttonLaunch;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBarA;
        private System.Windows.Forms.ComboBox keyColumn1ComboBox;
        private System.Windows.Forms.ComboBox sheet1ComboBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox sheet2ComboBox;
        private System.Windows.Forms.ComboBox keyColumn2ComboBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox file2TextBox;
        private System.Windows.Forms.Button file2Button;
        private System.Windows.Forms.CheckedListBox chosenCols1ListBox;
        private System.Windows.Forms.CheckedListBox chosenCols2ListBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.OpenFileDialog saveFileDialog;
    }
}

