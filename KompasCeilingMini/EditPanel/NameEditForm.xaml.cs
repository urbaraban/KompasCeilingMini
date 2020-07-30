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
    /// Логика взаимодействия для NameEditForm.xaml
    /// </summary>
    public partial class NameEditForm : Window
    {
        private KmpsDoc tempDoc;
        private ItemCollection ic;
        public NameEditForm(KmpsDoc TempDoc, ItemCollection IC)
        {
            InitializeComponent();
            this.tempDoc = TempDoc;
            this.ic = IC;

            DataSet ds = MainWindow.DefDataSet;
            
            MyList.DisplayMemberPath = "Name";
            MyList.SelectedValuePath = "InText";
            MyList.ItemsSource = ds.Tables["Variable"].DefaultView;
        }
        /*
        private DataSet comboDataSet(string sqlCmd, string TableName)
        {
            DataSet ds = new DataSet();

            string sqlCon = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\DefaultValue.mdf;Integrated Security=True";

            using (SqlConnection sqlConnection = new SqlConnection(sqlCon))
            {
                SqlDataAdapter da = new SqlDataAdapter(sqlCmd, sqlConnection);
                sqlConnection.Open();
                da.Fill(ds, TableName);

                sqlConnection.Close();
            }
            return ds;
        }*/

        private void MyList_PreviewMouseDown(object sender, MouseButtonEventArgs e)
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
            if (MyList.SelectedItems.Count > 0)
            {

                DataRowView mySelectedItem = MyList.SelectedItem as DataRowView;
                if (mySelectedItem != null)
                {
                    DragDrop.DoDragDrop(MyList, "{" + mySelectedItem["InText"].ToString().Replace(" ",string.Empty) + "}", DragDropEffects.Copy);
                }
            }
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

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.NameFileText = textBox.Text;
            KompasCeilingMini.Properties.Settings.Default.Save();
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ExampleLabel.Content = tempDoc.GiveMeText(textBox.Text, ic, KompasCeilingMini.Properties.Settings.Default.variable_suffix, string.Empty);
        }
    }
}
