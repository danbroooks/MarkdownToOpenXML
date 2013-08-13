
namespace WinFormsDemo
{
    using MarkdownToOpenXML;
    using System;
    using System.Diagnostics;
    using System.Windows.Forms;

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void save_Click(object sender, EventArgs e)
        {
            MD2OXML.CreateDocX(markdown.Text, file.Text);
            Process.Start(file.Text);
        }
    }
}
