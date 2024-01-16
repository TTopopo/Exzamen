using Microsoft.Office.Interop.Word;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
namespace Exzamen
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        } 
        private void button1_Click_1(object sender, EventArgs e)
        {
            // Создание нового документа Word
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Add();

            // Заполнение документа данными из текстового поля
            Word.Paragraph paragraph = doc.Content.Paragraphs.Add();

            paragraph.Range.Text = " " + label1.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = " " + label2.Text;
            paragraph.Range.InsertParagraphAfter(); 
            paragraph.Range.Text = "  " + textBox1.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + textBox2.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + textBox3.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + label3.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + textBox4.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + label4.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + label5.Text;
            paragraph.Range.InsertParagraphAfter(); 
            paragraph.Range.Text = "  " + textBox4.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + label6.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + textBox5.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + label7.Text;
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.Text = "  " + textBox7.Text;
            paragraph.Range.InsertParagraphAfter();
            // Сохранение документа
            object fileName = "5.doc";
            doc.SaveAs2(ref fileName);
            doc.Close();
            wordApp.Quit();

            MessageBox.Show("Документ успешно сохранен.");
        }
    }
}