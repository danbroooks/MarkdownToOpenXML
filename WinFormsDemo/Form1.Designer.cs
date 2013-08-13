namespace WinFormsDemo
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
            this.markdown = new System.Windows.Forms.TextBox();
            this.file = new System.Windows.Forms.TextBox();
            this.mdlabel = new System.Windows.Forms.Label();
            this.pathlabel = new System.Windows.Forms.Label();
            this.save = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // markdown
            // 
            this.markdown.Location = new System.Drawing.Point(12, 39);
            this.markdown.Multiline = true;
            this.markdown.Name = "markdown";
            this.markdown.Size = new System.Drawing.Size(332, 123);
            this.markdown.TabIndex = 0;
            this.markdown.Text = "_Text Document_\r\n\r\nThis is a test document.";
            // 
            // file
            // 
            this.file.Location = new System.Drawing.Point(51, 187);
            this.file.Name = "file";
            this.file.Size = new System.Drawing.Size(205, 20);
            this.file.TabIndex = 1;
            this.file.Text = "C:\\test.docx";
            // 
            // mdlabel
            // 
            this.mdlabel.AutoSize = true;
            this.mdlabel.Location = new System.Drawing.Point(12, 13);
            this.mdlabel.Name = "mdlabel";
            this.mdlabel.Size = new System.Drawing.Size(60, 13);
            this.mdlabel.TabIndex = 2;
            this.mdlabel.Text = "Markdown:";
            // 
            // pathlabel
            // 
            this.pathlabel.AutoSize = true;
            this.pathlabel.Location = new System.Drawing.Point(13, 190);
            this.pathlabel.Name = "pathlabel";
            this.pathlabel.Size = new System.Drawing.Size(32, 13);
            this.pathlabel.TabIndex = 3;
            this.pathlabel.Text = "Path:";
            // 
            // save
            // 
            this.save.Location = new System.Drawing.Point(269, 187);
            this.save.Name = "save";
            this.save.Size = new System.Drawing.Size(75, 20);
            this.save.TabIndex = 4;
            this.save.Text = "Save";
            this.save.UseVisualStyleBackColor = true;
            this.save.Click += new System.EventHandler(this.save_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(356, 226);
            this.Controls.Add(this.save);
            this.Controls.Add(this.pathlabel);
            this.Controls.Add(this.mdlabel);
            this.Controls.Add(this.file);
            this.Controls.Add(this.markdown);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "Form1";
            this.Text = "MD2OXML";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox markdown;
        private System.Windows.Forms.TextBox file;
        private System.Windows.Forms.Label mdlabel;
        private System.Windows.Forms.Label pathlabel;
        private System.Windows.Forms.Button save;
    }
}

