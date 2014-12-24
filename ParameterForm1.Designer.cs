namespace FinaquantInExcel
{
    partial class ParameterForm1
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
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.FuncTitle = new System.Windows.Forms.Label();
            this.FuncDescription = new System.Windows.Forms.RichTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(12, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(80, 80);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // FuncTitle
            // 
            this.FuncTitle.AutoSize = true;
            this.FuncTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FuncTitle.Location = new System.Drawing.Point(98, 14);
            this.FuncTitle.Name = "FuncTitle";
            this.FuncTitle.Size = new System.Drawing.Size(61, 16);
            this.FuncTitle.TabIndex = 1;
            this.FuncTitle.Text = "functitle";
            // 
            // FuncDescription
            // 
            this.FuncDescription.Location = new System.Drawing.Point(101, 33);
            this.FuncDescription.Name = "FuncDescription";
            this.FuncDescription.ReadOnly = true;
            this.FuncDescription.Size = new System.Drawing.Size(325, 60);
            this.FuncDescription.TabIndex = 2;
            this.FuncDescription.Text = "";
            this.FuncDescription.TextChanged += new System.EventHandler(this.FuncDescription_TextChanged);
            // 
            // ParameterForm1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(434, 362);
            this.Controls.Add(this.FuncDescription);
            this.Controls.Add(this.FuncTitle);
            this.Controls.Add(this.pictureBox1);
            this.Name = "ParameterForm1";
            this.Text = "ParameterForm1";
            this.Load += new System.EventHandler(this.ParameterForm1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label FuncTitle;
        private System.Windows.Forms.RichTextBox FuncDescription;



    }
}
