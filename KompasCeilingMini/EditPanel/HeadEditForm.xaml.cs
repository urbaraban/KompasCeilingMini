using KompasLib.Tools;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;


namespace KompasCeilingMini.EditPanel
{
    /// <summary>
    /// Логика взаимодействия для HeadEditForm.xaml
    /// </summary>
    public partial class HeadEditForm : Window
    {
        private KmpsDoc tempDoc;
        private ItemCollection ic;
        public HeadEditForm(KmpsDoc TempDoc, ItemCollection IC)
        {
            InitializeComponent();
            this.tempDoc = TempDoc;
            this.ic = IC;


            DataSet ds = comboDataSet("SELECT Name, InText FROM dbo.Variable WHERE Visible='True'", "Variable");

            VariableList.DisplayMemberPath = "Name";
            VariableList.SelectedValuePath = "InText";
            VariableList.ItemsSource = ds.Tables["Variable"].DefaultView;
            
            DataSet comboDataSet(string sqlCmd, string TableName)
            {
                ds = new DataSet();

                string sqlCon = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\DefaultValue.mdf;Integrated Security=True";

                using (SqlConnection sqlConnection = new SqlConnection(sqlCon))
                {
                    SqlDataAdapter da = new SqlDataAdapter(sqlCmd, sqlConnection);
                    sqlConnection.Open();
                    da.Fill(ds, TableName);

                    sqlConnection.Close();
                }
                return ds;
            }
        }

        private void VariableList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (VisualTreeHelper.GetParent(e.OriginalSource as UIElement) is ListBoxItem)
            {
                ListBoxItem item = (ListBoxItem)VisualTreeHelper.GetParent(e.OriginalSource as UIElement);
                if (item == null) return;
                if (item.IsSelected)
                {
                    e.Handled = true;
                }
            }
            if (VariableList.SelectedItems.Count > 0)
            {

                DataRowView mySelectedItem = VariableList.SelectedItem as DataRowView;
                if (mySelectedItem != null)
                {
                    DragDrop.DoDragDrop(VariableList, "{" + mySelectedItem["InText"].ToString().Replace(" ", string.Empty) + "}", DragDropEffects.Copy);
                }
            }
        }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.HeadText = textBox.Text;
            KompasCeilingMini.Properties.Settings.Default.Save();
        }

        private void TextBox_Drop(object sender, DragEventArgs e)
        {
            string tstring;
            tstring = e.Data.GetData(DataFormats.StringFormat).ToString().Trim(' ');
            textBox.Text = tstring;
        }

        private void TextBox_DragEnter(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if ((bool)refreshChek.IsChecked) ExampleLabel.Content = tempDoc.GiveMeText(textBox.Text, ic, KompasCeilingMini.Properties.Settings.Default.sufix, string.Empty);
        }

        private void refreshBtn_Click(object sender, RoutedEventArgs e)
        {
            ExampleLabel.Content = tempDoc.GiveMeText(textBox.Text, ic, KompasCeilingMini.Properties.Settings.Default.sufix, string.Empty);
        }
    }
}
