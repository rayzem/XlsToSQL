using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using MaterialDesignThemes.Wpf;

namespace XlsToSQL
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Title = "Selecione sua planilha";
            open.Filter = "Planilha Excel|*.xls;*.xlsx";
            open.ShowDialog(this);
            if (open.CheckFileExists)
            {
                txtPath.Text = open.FileName;
            }
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (String.IsNullOrWhiteSpace(txtTabela.Text))
                {
                    MessageBox.Show("Campo tabela é necessário.", "Campo não preenchido.", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (String.IsNullOrWhiteSpace(txtPath.Text))
                {
                    MessageBox.Show("Não foi apontado um caminho até a tabela.", "o arquivo não foi selecionado.", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                using (var stream = File.Open(txtPath.Text, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        if (comboTipo.SelectedIndex == 0)
                        {
                            string query = "INSERT INTO " + txtTabela.Text + " (";
                            for (int i = 0; i < result.Tables[0].Rows[0].ItemArray.Length; i++)
                            {
                                if (i + 1 == result.Tables[0].Rows[0].ItemArray.Length)
                                {
                                    query += "'" + result.Tables[0].Rows[0].ItemArray.GetValue(i).ToString() + "')";
                                }
                                else
                                {
                                    query += "'" + result.Tables[0].Rows[0].ItemArray.GetValue(i).ToString() + "',";
                                }
                            }
                            query += Environment.NewLine + " VALUES ";
                            for (int i = 1; i < result.Tables[0].Rows.Count; i++)
                            {
                                query += "( ";
                                for (int h = 0; h < result.Tables[0].Rows[i].ItemArray.Length; h++)
                                {
                                    if (h + 1 == result.Tables[0].Rows[i].ItemArray.Length && i + 1 == result.Tables[0].Rows.Count)
                                    {
                                        query += "'" + result.Tables[0].Rows[i].ItemArray.GetValue(h).ToString() + "');";
                                    }
                                    else if (h + 1 == result.Tables[0].Rows[i].ItemArray.Length)
                                    {
                                        query += "'" + result.Tables[0].Rows[i].ItemArray.GetValue(h).ToString() + "')," + Environment.NewLine;
                                    }
                                    else
                                    {
                                        query += "'" + result.Tables[0].Rows[i].ItemArray.GetValue(h).ToString() + "',";
                                    }
                                }
                                txtResult.Text = query;
                            }
                        }
                        else
                        {
                            string query = "Update " + txtTabela.Text + " SET ";
                            for (int x = 1; x < result.Tables[0].Rows.Count; x++)
                            {
                                for (int i = 0; i < result.Tables[0].Columns.Count; i++)
                                {
                                    if (i + 1 == result.Tables[0].Columns.Count)
                                    {
                                        query += result.Tables[0].Rows[0][i].ToString() + " = '" + result.Tables[0].Rows[x][i].ToString() + "'; " + Environment.NewLine;
                                        if (x + 1 != result.Tables[0].Rows.Count)
                                            query += "Update " + txtTabela.Text + " SET ";
                                    }

                                    else
                                    {
                                        query += result.Tables[0].Rows[0][i].ToString() + " = '" + result.Tables[0].Rows[x][i].ToString() + "', ";
                                    }
                                }
                            }
                            txtResult.Text = query;
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("O arquivo não é uma tabela ou está danificado.");
            }
        }

        private void txtResult_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            string text = (sender as TextBox).Text;
            (sender as TextBox).SelectAll();
            Clipboard.SetText(text);
        }

        private void btnLimpar_Click(object sender, RoutedEventArgs e)
        {
            txtPath.Text = String.Empty;
            txtResult.Text = String.Empty;
            txtTabela.Text = "";
        }
        private void btnExportar_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.AddExtension = true;
            save.Filter = "Arquivo Sql|*.sql;";
            save.ShowDialog(this);
            if (save.CheckPathExists)
            {
                try
                {
                    File.WriteAllText(save.FileName, txtResult.Text);
                    MessageBox.Show("Arquivo Salvo!", "Confirmação", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception)
                {

                }
            }
        }

    }
}
