using Microsoft.Office.Interop.Word;
using System;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string Fill()
        {
            string document_text = String.Format(
                "Уважаемый {0}, \r\n" +
                "Компания \"{1}\" благодарит Вас за сотрудничество, выражает глубокую признательность за верность и доверие, которое Вы оказываете нам, пользуясь нашими продуктами и методами. " +
                "Для поддерживания высокого качества услуг мы постоянно инвестируем средства на усовершенствование методоа и разработку новых. \r\n" +
                "Наряду с этим, из-за повышения количества желаюхих обучиться и расширения обучающей базы, мы вынужденны повысить цены на обучение и некоторые продукты. \r\n" +
                "Мы хотим подчеркнуть, что данный шаг необходим, прежде всего, для обеспечения стабильности, качества услуг и продуктов. \r\n\r\n" +
                "С {2} повысилась цена на услугу \"{3}\". \r\n\r\n" +
                "C уважением,\r\n " +
                "{4}.",
                textBox1.Text,
                textBox2.Text,
                dateTimePicker1.Text,
                textBox4.Text,
                textBox3.Text
                );
            return document_text;
        }

        public bool Check()
        {
            bool flag = true;
            foreach (var tb in this.Controls.OfType<TextBox>())
            {
                if (tb.Text == "" || tb.Text == "...")
                {
                    flag = false;
                }
            }
            return flag;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox5.Text = Fill();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox5.Text = Fill();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox5.Text = Fill();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            textBox5.Text = Fill();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            textBox5.Text = Fill();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Check())
            {
                Object Template = Type.Missing;
                Object NewTemplate = Type.Missing;
                Object Typpe = Type.Missing;
                Object Visible = Type.Missing;
                var Word1 = new Microsoft.Office.Interop.Word.Application();
                Word1.Visible = true;
                var newDocument = Word1.Documents.Add(ref Template, ref NewTemplate, ref Typpe, ref Visible);

                newDocument.Words.First.InsertBefore(Fill());

                foreach (Paragraph paragrph in newDocument.Paragraphs)
                {
                    paragrph.Range.Font.Color = WdColor.wdColorBlack;
                    paragrph.Range.Font.Size = 16;
                    paragrph.Range.Font.Name = "Times New Roman";
                    paragrph.Range.Font.Italic = 0;
                    paragrph.Range.Font.Bold = 0;
                }

                object begin = 0;
                object end = 0;
                Range wordrange = null;

                begin = Fill().IndexOf(textBox1.Text);
                end = (int)begin + textBox1.Text.Length;
                wordrange = newDocument.Range(ref begin, ref end);
                wordrange.Select();
                wordrange.Font.Italic = 1;

                begin = Fill().IndexOf(textBox2.Text);
                end = (int)begin + textBox2.Text.Length;
                wordrange = newDocument.Range(ref begin, ref end);
                wordrange.Select();
                wordrange.Font.Italic = 1;

                begin = Fill().IndexOf(dateTimePicker1.Text);
                end = (int)begin + dateTimePicker1.Text.Length;
                wordrange = newDocument.Range(ref begin, ref end);
                wordrange.Select();
                wordrange.Font.Italic = 1;

                begin = Fill().IndexOf(textBox4.Text);
                end = (int)begin + textBox4.Text.Length;
                wordrange = newDocument.Range(ref begin, ref end);
                wordrange.Select();
                wordrange.Font.Italic = 1;

                Object FileName = @"D:\" + textBox1.Text + ".doc";
                Object FileFormat = Type.Missing;
                Object LockComment = Type.Missing;
                Object Password = Type.Missing;
                Object AddToResentFile = Type.Missing;
                Object WritePass = Type.Missing;
                Object ReadOnlyRecommended = Type.Missing;
                Object EmbedTrue = Type.Missing;
                Object SaveNative = Type.Missing;
                Object SaveFormsData = Type.Missing;
                Object SaveAs = Type.Missing;
                Object Encoding = Type.Missing;
                Object InsertLine = Type.Missing;
                Object AllowSub = Type.Missing;
                Object LineEnd = Type.Missing;
                Object AddBi = Type.Missing;
                Object SaveChang = Type.Missing;
                Object OriginalFormat = Type.Missing;
                Object RouteDoc = Type.Missing;
                Word1.ActiveDocument.SaveAs(ref FileName,
                    ref FileFormat,
                    ref LockComment,
                    ref Password,
                    ref AddToResentFile,
                    ref WritePass,
                    ref ReadOnlyRecommended,
                    ref EmbedTrue,
                    ref SaveNative,
                    ref SaveFormsData,
                    ref SaveAs,
                    ref Encoding,
                    ref InsertLine,
                    ref AllowSub,
                    ref LineEnd,
                    ref AddBi);
                Word1.Application.Quit(ref SaveChang, ref OriginalFormat, ref RouteDoc);
            }
            else
            {
                MessageBox.Show("Заполните все необходимые поля");
            }
        }
    }
}
