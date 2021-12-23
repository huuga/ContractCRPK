using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace ContractCRPK
{
    public partial class Form2 : Form
    {
        private readonly string TemplateFileName = Directory.GetCurrentDirectory() + @"\test.docx";
        public Form2()
        {
            InitializeComponent();
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            var lastName = textBox1.Text;
            var firstName = textBox2.Text;
            var middleName = textBox3.Text;
            var psprt_ser = textBox4.Text;
            var psprt_num = textBox5.Text;
            var psprt_dprt = textBox6.Text;
            var psprt_date = textBox12.Text;
            var inn = textBox7.Text;
            var snils = textBox8.Text;
            var addr1 = textBox9.Text;
            var phone = textBox11.Text;

            var wordApp = new Word.Application();

            try
            {
                var wordDocument = wordApp.Documents.Open(TemplateFileName);

                char[] chars1 = firstName.ToCharArray();
                char[] chars2 = middleName.ToCharArray();
                char first1 = chars1[0];
                char first2 = chars2[0];
                string lastNameAndInitials = lastName + " " + first1 + "." + first2 + ".";

                for (int i = 0; i < 6; i++)
                {
                    ReplaceWordStub("{lastName}", lastName, wordDocument);
                    ReplaceWordStub("{firstName}", firstName, wordDocument);
                    ReplaceWordStub("{middleName}", middleName, wordDocument);
                    ReplaceWordStub("{signature}", lastNameAndInitials, wordDocument);
                }

                for (int i = 0; i < 5; i++)
                {
                    ReplaceWordStub("{signature}", lastNameAndInitials, wordDocument);
                }

                ReplaceWordStub("{psprt_ser}", psprt_ser, wordDocument);
                ReplaceWordStub("{psprt_num}", psprt_num, wordDocument);
                ReplaceWordStub("{psprt_dprt}", psprt_dprt, wordDocument);
                ReplaceWordStub("{psprt_date}", psprt_date, wordDocument);
                ReplaceWordStub("{inn}", inn, wordDocument);
                ReplaceWordStub("{snils}", snils, wordDocument);
                ReplaceWordStub("{addr1}", addr1, wordDocument);
                ReplaceWordStub("{phone}", phone, wordDocument);

                for (int i = 0; i < 3; i++)
                {
                    ReplaceWordStub("{sum}", (countOfParticipants * 1500).ToString(), wordDocument);
                }


                object objMiss = Missing.Value;
                Microsoft.Office.Interop.Word.Table tab = wordDocument.Tables[2];
                tab.Rows.Add(objMiss);
                tab.Cell(2, 1).Range.Text = "1";                
                tab.Cell(2, 2).Range.Text = comboBox3.Text;
                tab.Cell(2, 3).Range.Text = comboBox1.Text;
                tab.Cell(2, 4).Range.Text = textBox14.Text;
                tab.Cell(2, 5).Range.Text = comboBox2.Text;
                tab.Cell(2, 6).Range.Text = "1500";

                for (int i = 3; i < countOfParticipants + 2; i++)
                {
                    tab.Rows.Add(objMiss);
                    tab.Cell(i, 1).Range.Text = (i - 1).ToString();
                    tab.Cell(i, 2).Range.Text = comboBoxCompetitions[i - 3].Text;
                    tab.Cell(i, 3).Range.Text = comboBoxCategories[i - 3].Text;
                    tab.Cell(i, 4).Range.Text = textBoxesNames[i - 3].Text;
                    tab.Cell(i, 5).Range.Text = comboBoxStatuses[i - 3].Text;
                    tab.Cell(i, 6).Range.Text = "1500";
                }

                

                wordDocument.SaveAs2(System.Environment.GetEnvironmentVariable("USERPROFILE")+@"\Desktop\ДоговорЦРПК.docx");                
                wordApp.Documents.Close();

                MessageBox.Show("ВАШ ДОГОВОР СОХРАНЕН НА РАБОЧЕМ СТОЛЕ.\nМОЖЕТЕ ЗАКРЫТЬ WORD.",
                    "Сообщение",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
            }
            catch
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        int multiplier = 0;
        int countOfParticipants = 1;
        List<TextBox> textBoxesNames = new List<TextBox>();
        List<ComboBox> comboBoxCompetitions = new List<ComboBox>();
        List<ComboBox> comboBoxCategories = new List<ComboBox>();
        List<ComboBox> comboBoxStatuses = new List<ComboBox>();
        private void button2_Click(object sender, EventArgs e)
        {            
            TextBox txt = new TextBox();
            this.Controls.Add(txt);
            txt.Location = new Point(410, 531 + multiplier);
            txt.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            txt.Width = 165;
            txt.AutoSize = true;
            textBoxesNames.Add(txt);

            ComboBox cmb = new ComboBox();
            this.Controls.Add(cmb);
            cmb.Location = new Point(55, 531 + multiplier);
            cmb.Width = 224;
            cmb.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb.Items.AddRange(comboBox3.Items.Cast<object>().ToArray());
            comboBoxCompetitions.Add(cmb);

            ComboBox cmb2 = new ComboBox();
            this.Controls.Add(cmb2);
            cmb2.Location = new Point(285, 531 + multiplier);
            cmb2.Width = 119;
            cmb2.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb2.Items.AddRange(comboBox1.Items.Cast<object>().ToArray());
            comboBoxCategories.Add(cmb2);

            ComboBox cmb3 = new ComboBox();
            this.Controls.Add(cmb3);
            cmb3.Location = new Point(581, 531 + multiplier);
            cmb3.Width = 104;
            cmb3.DropDownStyle = ComboBoxStyle.DropDownList;
            cmb3.Items.AddRange(comboBox2.Items.Cast<object>().ToArray());
            comboBoxStatuses.Add(cmb3);

            Label lbl = new Label();
            this.Controls.Add(lbl);
            lbl.Location = new Point(35, 534 + multiplier);
            lbl.Text = (countOfParticipants + 1).ToString();

            Label lbl2 = new Label();
            this.Controls.Add(lbl2);
            lbl2.Location = new Point(705, 534 + multiplier);
            lbl2.Font = txt.Font = new Font("Arial", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            lbl2.Text = "1500";

            multiplier += 30;
            countOfParticipants += 1;
            btnDynTextBox.SetBounds(btnDynTextBox.Location.X, btnDynTextBox.Location.Y + 30, btnDynTextBox.Width, btnDynTextBox.Height);
            button1.SetBounds(button1.Location.X, button1.Location.Y + 30, button1.Width, button1.Height);
            this.Size = new Size(this.Size.Width, this.Size.Height + 50);
        }
    }
}
