using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace MediaAlunosExcel
{
    public partial class Form1 : Form
    {
        private string aluno, aprovado, livro, documento;
        private int freq1, freq2;
        private float media;
        private Excel.Application excelApp;
        private Word.Application wordApp;

        public Form1()
        {
            InitializeComponent();

            txtNome.Text = null;

            cbNota1.DropDownStyle = ComboBoxStyle.DropDownList;
            cbNota2.DropDownStyle = ComboBoxStyle.DropDownList;

            cbNota1.Items.Add("");
            cbNota2.Items.Add("");

            int aval;
            for (aval = 0; aval <= 20; aval++)
            {
                cbNota1.Items.Add(aval);
                cbNota2.Items.Add(aval);
            }

            cbNota1.Text = null;
            cbNota2.Text = null;

            txtMedia.Text = null;
            txtMedia.ReadOnly = true;

            IniciarExcel();

            if (VerificarPlanilha() == false)
            {
                CriarPlanilha();
            }
            else
            {
                AbrirPlanilha();
            }
        }

        private void IniciarExcel()
        {
            excelApp = new Excel.Application();
            excelApp.Visible = true;
        }

        private bool VerificarPlanilha()
        {
            livro = Application.StartupPath + "/alunos.xlsx";

            if (File.Exists(livro))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private bool VerificarDocumento()
        {
            if (File.Exists(documento))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void CriarPlanilha()
        {
            excelApp.Workbooks.Add();
            excelApp.Sheets["Planilha1"].Select();

            excelApp.Range["A1"].Value = "Nome";
            excelApp.Range["B1"].Value = "Nota1";
            excelApp.Range["C1"].Value = "Nota2";
            excelApp.Range["D1"].Value = "Média";
            excelApp.Range["E1"].Value = "Situação";

            FormatarPlanilha();
        }

        private void FormatarPlanilha()
        {
            excelApp.Range["A1:E1"].Font.Bold = true;
            excelApp.Range["A1:E1"].Interior.Color = Excel.Constants.xlGray50;
            excelApp.Range["A1:E1"].HorizontalAlignment = Excel.Constants.xlCenter;

            excelApp.Range["A:A"].ColumnWidth = 30;
            excelApp.Range["B:B"].ColumnWidth = 10;
            excelApp.Range["C:C"].ColumnWidth = 10;
            excelApp.Range["D:D"].ColumnWidth = 10;
            excelApp.Range["E:E"].ColumnWidth = 20;

            if (VerificarPlanilha() == false)
            {
                SalvarPlanilhaComo();
            }
            else
            {
                SalvarPlanilha();
            }
        }

        private void SalvarPlanilhaComo()
        {
            excelApp.ActiveWorkbook.SaveAs(livro);
        }

        private void SalvarPlanilha()
        {
            excelApp.ActiveWorkbook.Save();
        }

        private void SalvarDocumentoComo()
        {
            documento = Application.StartupPath + "/alunos.docx";
            wordApp.ActiveDocument.SaveAs2(documento);
        }

        private void AbrirPlanilha()
        {
            try
            {
                excelApp.Workbooks.Open(livro);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha : " + ex.Message);
            }
        }

        private bool ValidarDados()
        {
            try
            {
                if (txtNome.Text.Length < 2)
                {
                    MessageBox.Show("Introduza o nome do Aluno");
                    txtNome.SelectAll();
                    txtNome.Focus();
                    return false;
                }
                if (cbNota1.Text.Length == 0)
                {
                    MessageBox.Show("Selecione a Nota 1");
                    cbNota1.Focus();
                    return false;
                }
                if (cbNota2.Text.Length == 0)
                {
                    MessageBox.Show("Selecione a Nota 2");
                    cbNota2.Focus();
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha: " + ex.Message);
            }
            return true;
        }

        private void CalcularMedia()
        {
            aluno = txtNome.Text;
            freq1 = int.Parse(cbNota1.Text);
            freq2 = int.Parse(cbNota2.Text);
            media = (freq1 + freq2) / 2;

            if (freq1 < 8 && freq2 < 8)
            {
                aprovado = "Reprovado";
            }
            else if (freq1 < 8 && freq2 >= 8)
            {
                aprovado = "Repetir Nota1";
            }
            else if (freq2 < 8 && freq1 >= 8)
            {
                aprovado = "Repetir Nota 2";
            }
            else if (freq1 >= 8 && freq2 >= 8)
            {
                if (media >= 9.5)
                {
                    aprovado = "Aprovado";
                }
                else
                {
                    aprovado = "Prova oral";
                }
            }

            txtMedia.Text = ($"{media} ({aprovado})");
            ExportarExcel();
        }

        private void ExportarExcel()
        {
            excelApp.Sheets["Planilha1"].Select();

            int linhaExcel = 2;
            bool valor = true;
            while (valor == true)
            {
                if (excelApp.Range["A" + linhaExcel].Value != null)
                {
                    valor = true;
                    linhaExcel = linhaExcel + 1;
                }
                else
                {
                    valor = false;
                }
            }

            excelApp.Range["A" + linhaExcel].Value = aluno;
            excelApp.Range["B" + linhaExcel].Value = freq1;
            excelApp.Range["C" + linhaExcel].Value = freq2;
            excelApp.Range["D" + linhaExcel].Value = media;
            excelApp.Range["E" + linhaExcel].Value = aprovado;
            FormatarPlanilha();
            Cores();
        }

        private void Cores()
        {
            excelApp.Sheets["Planilha1"].Select();

            int linhaExcel = 2;
            bool valor = true;
            string sit;
            while (valor == true)
            {
                if (excelApp.Range["E" + linhaExcel].Value != null)
                {
                    sit = excelApp.Range["E" + linhaExcel].Value;

                    switch (sit)
                    {
                        case "Aprovado":
                            excelApp.Range["A" + linhaExcel + ":" + "E" + linhaExcel].Interior.Color = Color.FromArgb(0, 255, 0);
                            excelApp.Range["A" + linhaExcel + ":" + "E" + linhaExcel].Font.Color = Color.FromArgb(0, 0, 0);
                            break;

                        case "Reprovado":
                            excelApp.Range["A" + linhaExcel + ":" + "E" + linhaExcel].Interior.Color = Color.FromArgb(255, 0, 0);
                            excelApp.Range["A" + linhaExcel + ":" + "E" + linhaExcel].Font.Color = Color.FromArgb(255, 255, 255);
                            break;

                        case "Prova Oral":
                            excelApp.Range["A" + linhaExcel + ":" + "E" + linhaExcel].Interior.Color = Color.FromArgb(255, 255, 0);
                            excelApp.Range["A" + linhaExcel + ":" + "E" + linhaExcel].Font.Color = Color.FromArgb(0, 0, 0);
                            break;

                        default:
                            excelApp.Range["A" + linhaExcel + ":" + "E" + linhaExcel].Interior.Color = Excel.Constants.xlNone;
                            excelApp.Range["A" + linhaExcel + ":" + "E" + linhaExcel].Font.Color = Color.FromArgb(0, 0, 0);
                            break;
                    }
                    valor = true;
                    linhaExcel = linhaExcel + 1;
                }
                else
                {
                    valor = false;
                }
            }
            SalvarPlanilha();
        }

        private void Imprimir()
        {
            try
            {
                if (VerificarPlanilha() == true)
                {
                    SalvarPlanilha();
                    int LinhaExcel = 2;
                    bool valor = true;

                    while (valor == true)
                    {
                        if (excelApp.Range["E" + LinhaExcel].Value == null)
                        {
                            valor = false;
                        }
                        else
                        {
                            LinhaExcel = LinhaExcel + 1;
                        }
                    }

                    excelApp.Sheets["Planilha1"].Range["A1:E" + LinhaExcel].Copy();

                    wordApp = new Word.Application();
                    wordApp.Visible = true;
                    wordApp.Documents.Add();

                    wordApp.Selection.Font.Size = 22;
                    wordApp.Selection.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    wordApp.Selection.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wordApp.Selection.Font.Bold = 1;
                    wordApp.Selection.TypeText("Avaliações");

                    wordApp.Selection.Font.Size = 12;
                    wordApp.Selection.TypeParagraph();
                    wordApp.Selection.TypeParagraph();

                    wordApp.Selection.Paste();

                    wordApp.ActiveDocument.PrintOut();

                    SalvarDocumentoComo();
                }
                else
                {
                    MessageBox.Show("A planilha Excel não foi criada ainda!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha: " + ex.Message);
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (ValidarDados() == true)
            {
                CalcularMedia();
            }
        }

        private void btnImprimir_Click(object sender, EventArgs e)
        {
            Imprimir();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                SalvarPlanilha();
                SalvarDocumentoComo();
                excelApp.Quit();
                wordApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Falha: " + ex.Message);
            }
        }
    }
}