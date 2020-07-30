
using System.Diagnostics;
using System.IO;
using System.Windows;
using reference = System.Int32;
using System.Xml.Linq;
using System.Linq;
using System.Data.SqlClient;
using System;
using KAPITypes;
using Kompas6API7;
using Kompas6API5;
using Kompas6Constants;
using System.Windows.Media;
using System.Windows.Controls;
using System.Collections.Generic;

using System.Windows.Threading;

using KompasLib.Tools;
using System.Windows.Data;
using BlueTest.BlueClass;
using Windows.Devices.Enumeration;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth;
using System.Windows.Input;
using System.Runtime.InteropServices;
using System.Threading;
using KompasLib.KompasTool;

namespace KompasCeilingMini
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        static Dispatcher dispatcher = System.Windows.Application.Current.Dispatcher;


        #region Custom declarations

        public static KmpsAppl API;

        public static int Comp = KompasCeilingMini.Properties.Settings.Default.stg_compensation_val;

        public static DefDataSet DefDataSet = new DefDataSet();

        #endregion Custom declarations

        public MainWindow()
        {
            
            InitializeComponent();
            DefDataSet.FillTable();
            StatusLabel += StatusUpdate;
            Number += NumberUpdate;
            Suffix += SuffixUpdate;
            CheckResize();
        }

        private void NumberUpdate(object sender, int e)
        {
            NumberUpDn.Value = e;
        }

        private void StatusUpdate(object sender, string e)
        {
            KmpStatLbl.Content = e;
        }

        private void SuffixUpdate(object sender, string e)
        {
            sufixBox.Text = e;
        }


        #region MainVoid
        //выбираем текущий домент
        internal void targetApi()
        {
            if (KmpsAppl.KompasAPI == null)
            {
                KmpsAppl.ChangeDoc += KmpsAppl_ChangeDoc;
                KmpsAppl.ConnectBoolEvent += new EventHandler<bool>(ConnectEvent);

                if (KmpsAppl.KompasAPI != null)
                {
                    KmpStatLbl.Content = "Компас: ОК";
                }
                else
                {
                    MessageBox.Show(this, "Не найден активный объект", "Сообщение");
                    KmpStatLbl.Content = "Компас: Нет";
                }

                KmpsAppl.SelectDoc();

            }
            else
            {
                KmpsAppl.ChangeDoc += KmpsAppl_ChangeDoc;
                KmpsAppl.SelectDoc();
            }

           if (KmpsAppl.Doc. != null)
              KmpsAppl.Doc.SelectDimenetion += new EventHandler<object>(SelectDimChange);

            void ConnectEvent(object sender, bool e)
            {
                dispatcher.Invoke(new Action(() =>
                {
                    if (e)
                    {

                    }
                    else
                    {
                        KmpStatLbl.Content = "Компас: Нет";
                    }
                }));
                System.Threading.Thread.Sleep(50);
            }
        }

        private void KmpsAppl_ChangeDoc(object sender, KmpsDoc e)
        {
            dispatcher.Invoke(new Action(() =>
            {
                UpdateForm(true);
                KmpsAppl.Doc.ST.ChangeListDimention += ListDimentionChange;
            }));
            System.Threading.Thread.Sleep(50);
        }

        private void SelectDimChange(object sender, object e)
        {
            dispatcher.Invoke(new Action(() =>
            {
                AddingLineToggle.IsChecked = false;
                IDrawingObject lineDimension = (IDrawingObject)e;
                VariableDimentionlistBox.SelectedValue = lineDimension.Reference;
                DimVariableBox.Text = Math.Round(SizeTool.ReturnValVariableDim((IDrawingObject1)e), 2).ToString();
            }));
            System.Threading.Thread.Sleep(50);
        }

        #endregion

        //Отправляет на образмеривание

        private void NewFileVoid()
        {
            if (API != null)
            {
                if (!API.CreateDoc())
                    MessageBox.Show("Файл не создался");
            }
            else
            {
                System.Windows.MessageBox.Show(this, "Объект не захвачен", "Сообщение");
            }
        }

        private void OpenFileVoid()
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();

            openFileDialog.Filter = "Kompas (.frw)|*.frw|All Files (*.*)|*.*";

            if (KompasCeilingMini.Properties.Settings.Default.stg_work_folders == null || KompasCeilingMini.Properties.Settings.Default.stg_work_folders == "")
            {
                System.Windows.Forms.FolderBrowserDialog WorkFolderSlct = new System.Windows.Forms.FolderBrowserDialog();

                if (WorkFolderSlct.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    KompasCeilingMini.Properties.Settings.Default.stg_work_folders = WorkFolderSlct.SelectedPath;
                    KompasCeilingMini.Properties.Settings.Default.Save();
                }
            }

            openFileDialog.InitialDirectory = KompasCeilingMini.Properties.Settings.Default.stg_work_folders;
            openFileDialog.FileName = null;
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (KmpsAppl.KompasAPI == null)
                {
                    Process.Start(openFileDialog.FileName);
                }
                // Открыть документ с диска
                // первый параметр - имя открываемого файла
                // второй параметр указывает на необходимость выдачи запроса "Файл изменен. Сохранять?" при закрытии файла
                // третий параметр - указатель на IDispatch, по которому График вызывает уведомления об изменении своего состояния
                // ф-ия возвращает HANDLE открытого документа
                if (KmpsAppl.KompasAPI != null)
                {
                    int type = KmpsAppl.KompasAPI.ksGetDocumentTypeByName(openFileDialog.FileName);
                    switch (type)
                    {

                        case (int)DocType.lt_DocSheetStandart:  //2d документы
                        case (int)DocType.lt_DocFragment:
                            KmpsAppl.Doc.D5 = (ksDocument2D)KmpsAppl.KompasAPI.Document2D();
                            if (KmpsAppl.Doc.D5 != null)
                                Process.Start(openFileDialog.FileName);

                            break;
                    }

                    int err = KmpsAppl.KompasAPI.ksReturnResult();
                    if (err != 0)
                        KmpsAppl.KompasAPI.ksResultNULL();

                }
                else
                {
                    targetApi();
                }
            }
        }




        private string SqlReturnValue(string sqlCmd, string columName)
        {
            using (SqlConnection sqlConnection = new SqlConnection(KompasCeilingMini.Properties.Settings.Default.EasyCeilingConnectionString))
            {
                string value = string.Empty;

                SqlCommand com = new SqlCommand
                {
                    Connection = sqlConnection
                };
                sqlConnection.Open();

                SqlDataReader reader = null;
                com.CommandText = sqlCmd;

                reader = com.ExecuteReader();
                if (reader.Read())
                    value = reader[columName].ToString();

                sqlConnection.Close();

                return value;
            }
        }



        private void NewFileBtn_Click(object sender, RoutedEventArgs e)
        {
            NewFileVoid();
        }

        /// <summary>
        /// Клас загрузки в ComboBox
        /// </summary>
        private class ComboData
        {
            public string Name { get; set; }
            public double ID { get; set; }

            public ComboData(string _name, double _id)
            {
                Name = _name;
                ID = _id;
            }
        }

        private void KmpStatLbl_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            targetApi();
        }

        private void KsContrForm_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            DateTime date = DateTime.Now;
            KompasCeilingMini.Properties.Settings.Default.lastDate = date.Month;
            KompasCeilingMini.Properties.Settings.Default.Save();
        }

        private void KsContrForm_Loaded(object sender, RoutedEventArgs e)
        {
            DateTime date = DateTime.Now;
            if ((KompasCeilingMini.Properties.Settings.Default.lastDate < date.Month) && (date.Month != 1))
                if (MessageBox.Show("Новый месяц!\n Сменим папку?", "Внимание", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.OK)
                    FolderSelectDialog();

            targetApi();
        }

        private void FolderSelectDialog()
        {
            DateTime date = DateTime.Now;

            System.Windows.Forms.FolderBrowserDialog WorkFolderSlct = new System.Windows.Forms.FolderBrowserDialog();

            if (WorkFolderSlct.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                KompasCeilingMini.Properties.Settings.Default.stg_work_folders = WorkFolderSlct.SelectedPath;
                KompasCeilingMini.Properties.Settings.Default.lastDate = date.Month;
                KompasCeilingMini.Properties.Settings.Default.Save();
            }
        }


        private void KsContrForm_Closed(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        //выбрать рабочую папку
        private void WorkFolder_Click(object sender, RoutedEventArgs e)
        {
            FolderSelectDialog();
        }

        private void Menu_NewFile_Click(object sender, RoutedEventArgs e)
        {
            NewFileVoid();
        }

        private void Menu_OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileVoid();
        }

        #region SaveFile
        internal void SaveMe(bool saveas = false)
        {
            string oldName = string.Empty;

           /*if ((SqareUBox.Text == string.Empty) || (SqareUBox.Text == "0"))
            {
                MessageBox.Show("Не проставленна площадь по усадке (красным)");
                return;
            }*/

            if (API != null)
            {
                string name = string.Empty;

                //Если номера не совпадают, то предлагаем сохранить с новым
                if (NumberUpDn.Value != KVariable.Give("Number", string.Empty))
                {
                    if (MessageBox.Show("Сохраняем с новым номером " + NumberUpDn.Value + "?", "Внимание", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                    {
                        KompasCeilingMini.Properties.Settings.Default.ZkzLastNumber = (int)NumberUpDn.Value;
                        KompasCeilingMini.Properties.Settings.Default.Save();
                        name = KmpsAppl.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, CeilingItems, KompasCeilingMini.Properties.Settings.Default.variable_suffix, string.Empty);
                    }

                    else return;
                }
                //Если имя пустое а мы сохраняем компас согласно настройкам
                else if ((KmpsAppl.Doc.D7.Name != string.Empty) && (KompasCeilingMini.Properties.Settings.Default.stg_frw_check))
                {
                    //Если имя равно генирруемому
                    if (KmpsAppl.Doc.D7.Name.Substring(0, KmpsAppl.Doc.D7.Name.Length - 4) == KmpsAppl.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, CeilingItems, KompasCeilingMini.Properties.Settings.Default.variable_suffix, string.Empty))
                        //То используем как есть
                        name = KmpsAppl.Doc.D7.Name.Substring(0, KmpsAppl.Doc.D7.Name.Length - 4);
                    else
                    {
                        //Нет — делаем новое.
                        oldName = KmpsAppl.Doc.D7.PathName;
                        name = KmpsAppl.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, CeilingItems, KompasCeilingMini.Properties.Settings.Default.variable_suffix, string.Empty);
                    }
                }

                else
                {
                    //Если текущий номер уже сохранян как использованный
                    if ((int)NumberUpDn.Value <= KompasCeilingMini.Properties.Settings.Default.ZkzLastNumber)
                    {
                        if (MessageBox.Show("Вы уже использовали номер " + NumberUpDn.Value.ToString().Split('/', '-')[0] + ", или меньше\nДалее?", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {
                            KompasCeilingMini.Properties.Settings.Default.ZkzLastNumber = (int)NumberUpDn.Value;
                            KompasCeilingMini.Properties.Settings.Default.Save();
                            name = KmpsAppl.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, CeilingItems, KompasCeilingMini.Properties.Settings.Default.variable_suffix, string.Empty);
                        }
                        else
                        {
                            return; //если хотим номер изменить
                        }
                    }
                    else
                    {
                        KompasCeilingMini.Properties.Settings.Default.ZkzLastNumber = (int)NumberUpDn.Value;
                        KompasCeilingMini.Properties.Settings.Default.Save();
                        name = KmpsAppl.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, CeilingItems, KompasCeilingMini.Properties.Settings.Default.variable_suffix, string.Empty);
                    }
                }
                string path = (KompasCeilingMini.Properties.Settings.Default.stg_work_folders[0] + "\\" + name);
                FileInfo fi1 = new FileInfo(path + ".frw");
                if ((fi1.Exists) && (saveas == false))
                {
                    if (MessageBox.Show("Да — заменить\nНет — создаст новый", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        WriteIni();
                    }
                    else return;
                }
                else WriteIni();

                void WriteIni(bool menu = false)
                {
                    KmpsAppl.ProgressBar.Start(0, 3, "Сохранение", true);
                    ksRasterFormatParam rasParam = KmpsAppl.Doc.D5.RasterFormatParam();
                    //jpeg парам

                    if (rasParam != null)
                    {
                        rasParam.format = 2;
                        rasParam.extResolution = 30;
                    }
                    //frw Компас
                    if ((KmpsAppl.Doc. != null) && (menu == false))
                    {
                        KmpsAppl.ProgressBar.SetProgress(1, "Пишем компас", true);

                        ksDocumentParam documentParam = (ksDocumentParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);

                        KmpsAppl.Doc.D5.ksGetObjParam(KmpsAppl.Doc.D5.reference, documentParam, ldefin2d.ALLPARAM);
                        documentParam.author = Environment.UserName;
                        documentParam.comment = "KMCm";
                        KmpsAppl.Doc.D5.ksSetObjParam(KmpsAppl.Doc.D5.reference, documentParam, ldefin2d.ALLPARAM);


                        if (KompasCeilingMini.Properties.Settings.Default.stg_frw_check) KmpsAppl.Doc.D5.ksSaveDocument(path + ".frw");

                        SaveStatLabel.Content = "Сохранили";
                        if (KompasCeilingMini.Properties.Settings.Default.stg_jpg_check) KmpsAppl.Doc.D5.SaveAsToRasterFormat(KompasCeilingMini.Properties.Settings.Default.stg_work_folders[0] + "\\jpg\\" + name + ".jpg", rasParam);

                    }

                    if (oldName != string.Empty)
                    {
                        if (File.Exists(oldName)) File.Delete(oldName);
                        if (File.Exists(oldName + ".bak")) File.Delete(oldName + ".bak");
                        if (File.Exists(oldName.Replace(".frw", ".ini"))) File.Delete(oldName.Replace(".frw", ".ini"));
                    }

                    KmpsAppl.ProgressBar.Stop("Сохранили", true);
                }
                //Записывает INI документ
            }
            else
            {
                MessageBox.Show(this, "Объект не захвачен", "Сообщение");
            }
        }
        #endregion

        private void ReCalcBtn_Click(object sender, RoutedEventArgs e)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                bool mashtab = (bool)mashtabChek.IsChecked;
                bool noMashtab = (bool)NoMashtab.IsChecked;

                ReCalcBtn.Background = Brushes.YellowGreen;
                for (int i = 0; i < CeilingItems.Count; i++)
                    ContourCalc.GetContour(i, CeilingItems, SelectedFactura, ((DefDataSet.FacturaRow)SelectedFactura).Width,
                          mashtab, noMashtab,
                    KompasCeilingMini.Properties.Settings.Default.stg_compensation_val, KompasCeilingMini.Properties.Settings.Default.stg_compensation_val, false);

                UpdateForm(false);
            }
            else
            {
                MessageBox.Show(this, "Объект не захвачен", "Сообщение");
            }
        }

        private void OpenFileBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileVoid();
        }






        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveMe();
        }

        private void HeadBtn_Click(object sender, RoutedEventArgs e)
        {
            if (API != null)
            {
                double x = 0, y = 0;
                double width = KVariable.Give("XGabarit", "1");
                ksRequestInfo info = (ksRequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);

                int j = KmpsAppl.Doc.D5.ksCursor(info, ref x, ref y, 0);


                    KmpsAppl.Doc.CreateText(KompasCeilingMini.Properties.Settings.Default.HeadText,
                        x, y, width, CeilingItems,
                        KompasCeilingMini.Properties.Settings.Default.variable_suffix,
                        KompasCeilingMini.Properties.Settings.Default.stg_text_auto, false);

                double height = KompasCeilingMini.Properties.Settings.Default.HeadText.Split('\n').Length * width / 30;

               // KmpsAppl.Doc.CreateQrCode(true, height, x + 400, y);

            }
            else
            {
                MessageBox.Show(this, "Не найден активный объект", "Сообщение");
                KmpStatLbl.Content = "Компас: Нет";
            }

        }

        //Переносим в компас чертеж потолка
        private void ExportEC_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.Filter = "Kompas (.svg)|*.svg|All Files (*.*)|*.*";

            openFileDialog.InitialDirectory = KompasCeilingMini.Properties.Settings.Default.stg_work_folders;
            openFileDialog.FileName = null;
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (openFileDialog.FileName != null)
                {
                    XDocument XDoc;
                    XDoc = XDocument.Load(openFileDialog.FileName).FirstNode.Document;

                    for (int i = 0; i < XDoc.Descendants().Count() - 1; i++)
                    {
                        if (XDoc.Root.Descendants().ElementAt(i).Name == "{" + XDoc.Root.Attribute("xmlns").Value + "}" + "polygon")
                        {
                            XElement xElement = XDoc.Root.Descendants().ElementAt(i);
                            DrawPolygon(xElement);
                        }
                    }
                }
            }

            void DrawPolygon(XElement xElement)
            {
                string[] point = xElement.Attribute("points").Value.ToString().Split(' ');
                if (point.Length > 1)
                {
                    if (MessageBox.Show("Да — Точки\nНет — Линии", "Варианты", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                    {
                        for (int i = 0; i < point.Length; i++)
                        {
                            if (point[i] != "") KmpsAppl.Doc.D5.ksPoint(double.Parse(point[i].Split(',')[0].Replace('.', ',')) / 10, double.Parse(point[i].Split(',')[1].Replace('.', ',')) / 10, 1);
                        }
                    }
                    else
                    {
                        for (int i = 1; i < point.Length; i++)
                        {
                            if (point[i] != "")
                            {
                                double X1, Y1, X2, Y2;
                                X1 = Math.Round(double.Parse(point[i].Split(',')[0].Replace('.', ',')) / Comp, 4);
                                Y1 = -Math.Round(double.Parse(point[i].Split(',')[1].Replace('.', ',')) / Comp, 4);
                                X2 = Math.Round(double.Parse(point[i - 1].Split(',')[0].Replace('.', ',')) / Comp, 4);
                                Y2 = -Math.Round(double.Parse(point[i - 1].Split(',')[1].Replace('.', ',')) / Comp, 4);

                                KmpsAppl.Doc.D5.ksLineSeg(X1, Y1, X2, Y2, 2);
                            }
                            else
                            {
                                double X1, Y1, X2, Y2;
                                X1 = Math.Round(double.Parse(point[point.Length - 2].Split(',')[0].Replace('.', ',')) / Comp, 4);
                                Y1 = -Math.Round(double.Parse(point[point.Length - 2].Split(',')[1].Replace('.', ',')) / Comp, 4);
                                X2 = Math.Round(double.Parse(point[0].Split(',')[0].Replace('.', ',')) / Comp, 4);
                                Y2 = -Math.Round(double.Parse(point[0].Split(',')[1].Replace('.', ',')) / Comp, 4);
                                KmpsAppl.Doc.D5.ksLineSeg(X1, Y1, X2, Y2, 2);
                            }
                        }
                    }
                }
                else MessageBox.Show("Что-то пошло не так.");
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://money.yandex.ru/to/410011060113741");
        }

        private async void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.variable_suffix = sufixBox.Text;
             await KVariable.UpdateNote("Suffix", sufixBox.Text, string.Empty);
            KompasCeilingMini.Properties.Settings.Default.Save();
        }


        private void menuConnect_Click(object sender, RoutedEventArgs e)
        {
            targetApi();
        }

        private async void paramBox_1_LostFocus(object sender, RoutedEventArgs e)
        {
            if (KmpsAppl.Doc. != null)
                await KVariable.UpdateNote("Comment1", paramBox_1.Text, string.Empty);
        }

        private async void paramBox_2_LostFocus(object sender, RoutedEventArgs e)
        {
            if (KmpsAppl.Doc. != null)
                await KVariable.UpdateNote("Comment2", paramBox_2.Text, string.Empty);
        }

        private async void NumberUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (API != null)
                await KVariable.UpdateAsync("Number", (double)NumberUpDn.Value, string.Empty);
        }

        private void MakeOrd_Click(object sender, RoutedEventArgs e)
        {
            if (API != null)
                SizeTool.Coordinate(SelectedFactura.ToString(),
                    ((DefDataSet.FacturaRow)SelectedFactura).Width, 
                    KompasCeilingMini.Properties.Settings.Default.stg_crd_dopusk,
                    KompasCeilingMini.Properties.Settings.Default.stg_text_auto ? 
                    KVariable.Give("XGabarit", SelectedFactura.ToString()) / 40 * 2 :
                    KompasCeilingMini.Properties.Settings.Default.stg_text_size);
        }


        private void KsContrForm_Deactivated(object sender, EventArgs e)
        {
            if ((this.Focusable == true)&&(MainFrame.ResizeMode == ResizeMode.NoResize))
            {
                this.Opacity = 0.4;
            }
        }

        private void KsContrForm_Activated(object sender, EventArgs e)
        {
            if (MainFrame.ResizeMode == ResizeMode.NoResize)
            {
                this.Opacity = 1;
            }
        }


        //Класс для возврата
        public class ComboBT
        {
            public string Name { get; set; }
            public string Mac { get; set; }

            public ComboBT(string _name, string _mac)
            {
                Name = _name;
                Mac = _mac;
            }
        }

        public void ChangeDoc(object sender, KmpsAppl e)
        {

        }

        private void WidthUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (MainWindow.API != null)
            {
                double size = KVariable.Give("XGabarit", facturaListCheck.SelectedItem.ToString());
                if ((double)WidthUpDn.Value / size < 1) XUpDn.Value = Math.Round(((1 - (double)WidthUpDn.Value / size) * 100), 1);
            }
        }

        private void HeightUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (MainWindow.API != null)
            {
                double size = KVariable.Give("YGabarit", facturaListCheck.SelectedItem.ToString());
                if ((double)HeightUpDn.Value / size < 1) YUpDn.Value = Math.Round(((1 - (double)HeightUpDn.Value / size) * 100), 1);
            }
        }

        private async void UpdateForm(bool First = true)
        {
            if ((KmpsAppl.KompasAPI != null) && (KmpsAppl.Appl.ActiveDocument != null))
            {

                if (KmpsAppl.Appl.ActiveDocument.Name != string.Empty)
                    StatusLabel(this, KmpsAppl.Appl.ActiveDocument.Name);
                else StatusLabel(this, "Новый документ");


                IVariable7 var;

                if (First == true)
                {
                    if (!KmpsAppl.Doc.IsVariableNameValid("Number"))
                    {
                        facturaListCheck.Items.Clear();

                        var = KmpsAppl.Doc.Variable[false, "Number"];
                        if (var.Note != "Number")
                        {
                            string[] numberList = var.Note.Split(';');
                            for (int i = 0; i < numberList.Length; i++) facturaListCheck.Items.Add(numberList[i]);
                            facturaListCheck.SelectedIndex = 0;
                        }
                        else
                        {
                            var.Note = "1";
                            KmpsAppl.Doc.UpdateVariables();
                        }
                    }
                    else
                    {
                        facturaListCheck.Items.Clear();
                        KVariable.Add("Number", 1, "Номера");
                        var = KmpsAppl.Doc.Variable[false, "Number"];
                        var.Note = "1";
                        KmpsAppl.Doc.UpdateVariables();
                        facturaListCheck.Items.Add("1");
                        facturaListCheck.SelectedIndex = 0;
                    }

                    mashtabChek.IsChecked = false;
                }

                string index = facturaListCheck.SelectedItem.ToString();


                //Number
                if (KmpsAppl.Doc.IsVariableNameValid("Number")) KVariable.Add("Number", 0, "Площадь");
                var = KVariable. Variable[false, "Number"];
                Number(this, (int)var.Value);
                //suffix
                if (KmpsAppl.Doc.IsVariableNameValid("Suffix")) KVariable.Add("Suffix", 0, "Площадь");
                var = KmpsAppl.Doc.Variable[false, "Suffix"];
                Suffix(this, var.Note);

                //площадь
                if (KmpsAppl.Doc.IsVariableNameValid("Sqare" + index)) KVariable.Add("Sqare" + index, 0, "Площадь");
                var = KmpsAppl.Doc.Variable[false, "Sqare" + index];
                SqareBox.Text = var.Value.ToString();

                //площадь с усадкой
                if (KmpsAppl.Doc.IsVariableNameValid("SqareU" + index)) KVariable.Add("SqareU" + index, 0, "ПлощадьУ");
                var = KmpsAppl.Doc.Variable[false, "SqareU" + index];
                SqareUBox.Text = var.Value.ToString();

                //периметр
                if (KmpsAppl.Doc.IsVariableNameValid("Perimetr" + index)) KVariable.Add("Perimetr" + index, 0, "Perimetr");
                var = KmpsAppl.Doc.Variable[false, "Perimetr" + index];
                PerimetrBox.Text = var.Value.ToString();

                //периметрУ
                if (KmpsAppl.Doc.IsVariableNameValid("PerimetrU" + index)) KVariable.Add("PerimetrU" + index, 0, "PerimetrU");
                //периметр прямых
                if (KmpsAppl.Doc.IsVariableNameValid("LineP" + index)) KVariable.Add("LineP" + index, 0, "LineP");
                var = KmpsAppl.Doc.Variable[false, "LineP" + index];

                //периметр кривых
                if (KmpsAppl.Doc.IsVariableNameValid("CurveP" + index)) KVariable.Add("CurveP" + index, 0, "CurveP");
                var = KmpsAppl.Doc.Variable[false, "CurveP" + index];

                //углы
                if (KmpsAppl.Doc.IsVariableNameValid("Angle")) KVariable.Add("Angle", 4, "Angle");
                var = KmpsAppl.Doc.Variable[false, "Angle"];
                AngUpDn.Value = (byte)var.Value;

                //шов
                if (KmpsAppl.Doc.IsVariableNameValid("Shov" + index)) KVariable.Add("Shov" + index, 0, "Shov");
                var = KmpsAppl.Doc.Variable[false, "Shov" + index];
                ShovBox.Text = var.Value.ToString();


                //XGabarit
                if (KmpsAppl.Doc.IsVariableNameValid("XGabarit" + index)) KVariable.Add("XGabarit" + index, 0, "XGabarit");
                var = KmpsAppl.Doc.Variable[false, "XGabarit" + index];
                WidthUpDn.Value = Math.Round((var.Value * ((100 - KVariable.Give("koefX", string.Empty)) / 100)), 1);

                //YGabarit
                if (KmpsAppl.Doc. KmpsAppl.Doc.IsVariableNameValid("YGabarit" + index)) KVariable.Add("YGabarit" + index, 0, "YGabarit");
                var = KmpsAppl.Doc.Variable[false, "YGabarit" + index];
                HeightUpDn.Value = Math.Round((var.Value * ((100 - KVariable.Give("koefY", string.Empty)) / 100)), 1);

                //Xcrd
                if (KmpsAppl.Doc.IsVariableNameValid("Xcrd" + index)) KVariable.Add("Xcrd" + index, 0, "Xcrd");
                var = KmpsAppl.Doc.Variable[false, "Xcrd" + index];

                //Ycrd
                if (KmpsAppl.Doc.IsVariableNameValid("Ycrd" + index)) KVariable.Add("Ycrd" + index, 0, "Ycrd");
                var = KmpsAppl.Doc.Variable[false, "Ycrd" + index];


                //X
                if (KmpsAppl.Doc.IsVariableNameValid("koefX")) KVariable.Add("koefX", 7, "Unchecked");
                var = KmpsAppl.Doc.Variable[false, "koefX"];
                XUpDn.Value = var.Value;
                if (var.Note == "koefX") await KVariable.UpdateNote("koefX", "Unchecked", "");
                //Статус усадки;
                if (First)
                {
                    if (var.Note == "Checked") mashtabChek.IsChecked = true;
                    else mashtabChek.IsChecked = false;
                }

                //Y
                if (KmpsAppl.Doc.IsVariableNameValid("koefY")) KVariable.Add("koefY", 7, "koefY");
                var = KmpsAppl.Doc.Variable[false, "koefY"];
                YUpDn.Value = Math.Round(var.Value, 1);


                //фотопечать
                if (KmpsAppl.Doc.IsVariableNameValid("photo" + index)) KVariable.Add("photo" + index, 0, "photo");
                var = KmpsAppl.Doc.Variable[false, "photo" + index];
                FpUpDn.Value = Math.Round(var.Value, 1);

                //длина
                if (KmpsAppl.Doc.IsVariableNameValid("lenth" + index)) KVariable.Add("lenth" + index, 0, "lenth");
                var = KmpsAppl.Doc.Variable[false, "lenth" + index];

                //фактура
                if (KmpsAppl.Doc.IsVariableNameValid("factura" + index)) KVariable.Add("factura" + index, -1, "factura");
                var = KmpsAppl.Doc.Variable[false, "factura" + index];

                double sizeX = KVariable.Give("XGabarit", index) * ((100 - KVariable.Give("koefX", string.Empty)) / 100);
                double sizeY = KVariable.Give("YGabarit", index) * ((100 - KVariable.Give("koefY", string.Empty)) / 100);
                DataSetFacturaRefresh("Width >= " + Math.Min(sizeX * 0.97, sizeY * 0.97).ToString().Replace(',', '.'), "NumberPP ASC");
                if (FacturaCombo.Items.Count == 0)
                    DataSetFacturaRefresh(string.Empty, "NumberPP ASC");


                //цвет
                if (KmpsAppl.Doc.IsVariableNameValid("color" + index)) KVariable.Add("color" + index, -1, "color");
                var = KmpsAppl.Doc.Variable[false, "color" + index];
                DataSetFacturaColorRefresh();

                if (var.Note == "color" + index) await KVariable.UpdateNote("color", ColorCombo.Text, index);

                //вырез
                if (KmpsAppl.Doc.IsVariableNameValid("cut" + index)) KVariable.Add("cut" + index, 0, "вырез");
                var = KmpsAppl.Doc.Variable[false, "cut" + index];
                CutBox.Text = var.Value.ToString();

                //Комментарий 1
                if (KmpsAppl.Doc.IsVariableNameValid("Comment1")) KVariable.Add("Comment1", 0, "Comment1");
                var = KmpsAppl.Doc.Variable[false, "Comment1"];
                paramBox_1.Text = var.Note;
                //Комментарий 2
                if (KmpsAppl.Doc.D5.IsVariableNameValid("Comment2")) KVariable.Add("Comment2", 0, "Comment2");
                var = KmpsAppl.Doc.Variable[false, "Comment2"];
                paramBox_2.Text = var.Note;

            }
            else if (KmpsAppl.Appl.ActiveDocument == null)
            {
                StatusLabel(this, "Пусто");
            }
        }


        private void DataSetFacturaRefresh(string StrFilter, string StrOrder)
        {
            KmpsAppl.someFlag = false;
            FacturaCombo.ItemsSource = null;
            FacturaCombo.Items.Clear();

            if (DefDataSet != null)
            {
                FacturaCombo.DisplayMemberPath = "DisplayName";
                FacturaCombo.SelectedValuePath = "IDFactura";
                FacturaCombo.ItemsSource = DefDataSet.Tables["Factura"].Select(StrFilter, StrOrder);

                FacturaCombo.SelectedValue = KVariable.Give("factura", facturaListCheck.SelectedItem.ToString());
            }
            else FacturaCombo.Items.Add("Не найдена база");
            KmpsAppl.someFlag = true;

            if (FacturaCombo.SelectedIndex < 0)
                FacturaCombo.SelectedIndex = 0;

            SelectedFactura = FacturaCombo.SelectedItem;

            if (ColorCombo.SelectedIndex < 0)
                ColorCombo.SelectedIndex = 0;

            SelectedColor = ColorCombo.SelectedItem;
        }


        private void CalcAllBtn_Click(object sender, RoutedEventArgs e)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                CalcAllBtn.IsEnabled = false;
                bool mashtab = (bool)mashtabChek.IsChecked;
                bool noMashtab = (bool)NoMashtab.IsChecked;
                int index = facturaListCheck.SelectedIndex;

                ContourCalc.GetContour(index, CeilingItems, SelectedFactura, ((DefDataSet.FacturaRow)SelectedFactura).Width,
                          mashtab, noMashtab,
                    KompasCeilingMini.Properties.Settings.Default.stg_compensation_val, KompasCeilingMini.Properties.Settings.Default.stg_compensation_val, true);

                KmpsAppl.Appl.StopCurrentProcess();
                UpdateForm(false);
                CalcAllBtn.IsEnabled = true;
            }
            else
            {
                MessageBox.Show("Объект не захвачен", "Сообщение");
            }
        }

        private async void PlusFacturaBtn_Copy_Click(object sender, RoutedEventArgs e)
        {
            if (facturaListCheck.Items.Count > 1)
            {
                await KVariable.ClearAsync(facturaListCheck.SelectedItem.ToString(), DefDataSet.Tables["Variable"]);
                KmpsAppl.Doc.Macro.RemoveCeilingMacro(facturaListCheck.SelectedItem.ToString());
                facturaListCheck.Items.Remove(facturaListCheck.SelectedItem);
            }
        }

        private void MashtabChek_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.RightButton == System.Windows.Input.MouseButtonState.Pressed)
                mashtabChek.IsChecked = !mashtabChek.IsChecked;
            else if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
            {

            }
        }

        private void mashtabChek_Click(object sender, RoutedEventArgs e)
        {
            if (MainWindow.API != null)
            {
                mashtabChek.IsChecked = KmpsAppl.Doc.Mashtab(!(bool)mashtabChek.IsChecked);
                //Пересчитываем контуры
                bool mashtab = (bool)mashtabChek.IsChecked;
                bool noMashtab = (bool)NoMashtab.IsChecked;

                for (int i = 0; i < CeilingItems.Count; i++)
                    ContourCalc.GetContour(i, CeilingItems, SelectedFactura, ((DefDataSet.FacturaRow)SelectedFactura).Width,
                          mashtab, noMashtab,
                    KompasCeilingMini.Properties.Settings.Default.stg_compensation_val, KompasCeilingMini.Properties.Settings.Default.stg_compensation_val, false);

                UpdateForm(false);
            }
            else
                MessageBox.Show("Объект не захвачен", "Сообщение");
        }

        private void PlusFacturaBtn_Click(object sender, RoutedEventArgs e)
        {
            if (KmpsAppl.Doc != null)
            {
                facturaListCheck.Items.Add(facturaListCheck.Items.Count + 1);

                {
                    IVariable7 var = KmpsAppl.Doc.Variable[false, "Number"];
                    var.Note = GiveNumberStroke();
                    KmpsAppl.Doc.UpdateVariables();
                }

                facturaListCheck.SelectedIndex = facturaListCheck.Items.Count - 1;

                string GiveNumberStroke()
                {
                    string number = facturaListCheck.Items[0].ToString();
                    for (int i = 1; i < facturaListCheck.Items.Count; i++)
                        number += ";" + facturaListCheck.Items[i].ToString();
                    return number;
                }
            }
        }

        private void DataSetFacturaColorRefresh()
        {
            if (FacturaCombo.SelectedIndex > -1)
            {

                ColorCombo.ItemsSource = null;
                KmpsAppl.someFlag = false;
                ColorCombo.Items.Clear();

                //DataSet ds = sqltool.ComboDataSet("SELECT IDFacCol, Name FROM dbo.FactruaColor WHERE IDFactura=" + FacturaCombo.SelectedValue + " ORDER BY NumberPP", "FactruaColor", KompasCeilingMini.Properties.Settings.Default.EasyCeilingConnectionString);

                if (DefDataSet != null)
                {
                    ColorCombo.DisplayMemberPath = "Name";
                    ColorCombo.SelectedValuePath = "IDFacCol";
                    ColorCombo.ItemsSource = DefDataSet.Tables["FacturaColor"].Select("IDFactura = " + FacturaCombo.SelectedValue, "NumberPP ASC");


                    ColorCombo.SelectedValue = KVariable.Give("color", facturaListCheck.SelectedItem.ToString());
                }
                else ColorCombo.Items.Add("Не найдена база");

                KmpsAppl.someFlag = true;

                if (ColorCombo.SelectedIndex < 0)
                    ColorCombo.SelectedIndex = 0;
            }

        }

        private void FacturaListCheck_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (facturaListCheck.Items.Count > 0)
                if (KmpsAppl.KompasAPI != null)
                {
                    if (facturaListCheck.SelectedIndex < 0) facturaListCheck.SelectedIndex = 0;

                    CeilingItems = facturaListCheck.Items;
                    SelectedCeiling = facturaListCheck.SelectedItem;

                    IVariable7 var = KmpsAppl.Doc.Variable[false, "Number"];
                    var.Note = GiveNumberStroke();
                    KmpsAppl.Doc.UpdateVariables();
                    UpdateForm(false);

                    //Возвращает строку всех фактур
                    string GiveNumberStroke()
                    {
                        string number = facturaListCheck.Items[0].ToString();
                        for (int i = 1; i < facturaListCheck.Items.Count; i++) number += ";" + facturaListCheck.Items[i].ToString();
                        return number;
                    }
                }
        }

        #region CalcAll

        //Изменили документ
        public static event EventHandler<String> StatusLabel;
        public static event EventHandler<Int32> Number;
        public static event EventHandler<String> Suffix;
        public static ItemCollection CeilingItems;
        public static object SelectedCeiling;
        public static bool Mashtab;
        public static object SelectedFactura;
        public static object SelectedColor;



        private async void FacturaCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if ((KmpsAppl.someFlag) && (KmpsAppl.KompasAPI != null) && (FacturaCombo.SelectedIndex > -1))
            {
                SelectedFactura = FacturaCombo.SelectedItem;
                await KVariable.UpdateAsync("factura", (double)((DefDataSet.FacturaRow)FacturaCombo.SelectedItem).IDFactura, facturaListCheck.SelectedItem.ToString());
                await KVariable.UpdateNote("factura", ((DefDataSet.FacturaRow)FacturaCombo.SelectedItem).DisplayName.ToString(), facturaListCheck.SelectedItem.ToString());
                DataSetFacturaColorRefresh();
            }
        }

        private async void ColorCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if ((KmpsAppl.someFlag) && (ColorCombo.SelectedIndex > -1))
            {
                SelectedColor = ColorCombo.SelectedItem;
                await KVariable.UpdateAsync("color", ((DefDataSet.FacturaColorRow)ColorCombo.SelectedItem).IDFacCol, facturaListCheck.SelectedItem.ToString());
                await KVariable.UpdateNote("color", ((DefDataSet.FacturaColorRow)ColorCombo.SelectedItem).Name, facturaListCheck.SelectedItem.ToString());
            }
        }

        private async void AngUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (MainWindow.API != null)
                await KVariable.UpdateAsync("Angle", (double)AngUpDn.Value, string.Empty);
        }
        #endregion

        #region DimPanel
        private BluetoothObserver observer;
        public void ListDimentionChange(object sender, List<List<SizeTool.ComboData>> e)
        {
            dispatcher.Invoke(new Action(() =>
            {
                VariableDimentionlistBox.DataContext = null;
                VariableDimentionlistBox.DisplayMemberPath = "Name";
                VariableDimentionlistBox.SelectedValuePath = "Reference";
                VariableDimentionlistBox.ItemsSource = e[0];

                FreeDimentionlistBox.DataContext = null;
                FreeDimentionlistBox.DisplayMemberPath = "Name";
                FreeDimentionlistBox.SelectedValuePath = "Reference";
                FreeDimentionlistBox.ItemsSource = e[1];

                FixDimentionlistBox.DataContext = null;
                FixDimentionlistBox.DisplayMemberPath = "Name";
                FixDimentionlistBox.SelectedValuePath = "Reference";
                FixDimentionlistBox.ItemsSource = e[2];
            }));
            System.Threading.Thread.Sleep(50);
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            ConnectObserver();
        }

        private void ConnectObserver()
        {
            if (observer == null)
            {
                observer = new BluetoothObserver();
                observer.statlabelChange += new EventHandler<string>(ChangeStatusBT);
                observer.onLastDimention += new EventHandler<double>(SetLastDimention);
                GC.KeepAlive(observer);
            }
            if (BTcombo.SelectedItem != null)
                if ((BluetoothObserver.Watcher == null) || (BluetoothObserver.Watcher.Status == Windows.Devices.Bluetooth.Advertisement.BluetoothLEAdvertisementWatcherStatus.Stopped) && (!observer.isFindDevice))
                    observer.Start(BTcombo.Text, BTcombo.SelectedValue.ToString(), "0000ffb0-0000-1000-8000-00805f9b34fb", "0000ffb2-0000-1000-8000-00805f9b34fb");

               /* else if ((observer.isFindDevice) || (BluetoothObserver.Watcher.Status == Windows.Devices.Bluetooth.Advertisement.BluetoothLEAdvertisementWatcherStatus.Stopped))
                {
                    BluetoothObserver.Watcher.Stop();
                    observer.isFindDevice = false;
                }*/
        }


        private void ChangeStatusBT(object sender, string e)
        {
            dispatcher.Invoke( new Action(() =>
            {
                ConnectBtBtn.Content = e;
            }));
            System.Threading.Thread.Sleep(50);
        }

        private void SetLastDimention(object sender, double e)
        {
            dispatcher.Invoke(new Action(() =>
            {
                KmpsAppl.Doc.ST.UpdateOrMakeLine(e * (1000/Comp), (bool)AddingLineToggle.IsChecked, true, KmpsAppl.Doc.GiveSelectOrChooseObj());
                if (VariableDimentionlistBox.SelectedItems.Count == 1)
                    VariableDimentionlistBox.SelectedIndex = VariableDimentionlistBox.SelectedIndex + 1 > VariableDimentionlistBox.Items.Count - 1 ? 0 : VariableDimentionlistBox.SelectedIndex + 1;

                KmpsAppl.ZoomAll();
            }));
            System.Threading.Thread.Sleep(100);
        }


        private async void DimentionlistBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ListBox listBox = (ListBox)sender;

            if (listBox.SelectedItems != null)
            {
                KmpsAppl.Doc.GetChooseContainer().UnchooseAll();
                // ConnectObserver();
                foreach (SizeTool.ComboData data in listBox.SelectedItems)
                {
                    DimVariableBox.Text = await Task<string>.Factory.StartNew(() =>
                    {
                        object Dim = (object)KmpsAppl.KompasAPI.TransferReference(data.Reference, KmpsAppl.Doc.D5.reference);
                        if (Dim != null)
                        {
                            KmpsAppl.Doc.GetChooseContainer().Choose(Dim);

                            IDrawingObject1 Dim1 = (IDrawingObject1)Dim;
                            Array arrayConstrait = (Array)Dim1.Constraints;

                            if (arrayConstrait != null)
                                foreach (IParametriticConstraint constraint in arrayConstrait)
                                    if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                                         return Math.Round(KVariable.Give(constraint.Variable, string.Empty),2).ToString(); //вместо 100 компенсатор

                            if (!KmpsAppl.Doc.GetSelectContainer().IsSelected((IDrawingObject)Dim)) KmpsAppl.Doc.GetSelectContainer().UnselectAll();
                        }
                        return string.Empty;
                    });

                    DimVariableBox.SelectAll();
                }
            }
        }

        private void DimVariableBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {

            if (Keyboard.Modifiers == ModifierKeys.None && e.Key == System.Windows.Input.Key.Enter)
            {
                KmpsAppl.Doc.ST.UpdateOrMakeLine(double.Parse(DimVariableBox.Text.Replace('.', ',')), (bool)AddingLineToggle.IsChecked, true, KmpsAppl.Doc.GiveSelectOrChooseObj());
                
                if (VariableDimentionlistBox.SelectedItems.Count == 1)
                    VariableDimentionlistBox.SelectedIndex = VariableDimentionlistBox.SelectedIndex + 1 > VariableDimentionlistBox.Items.Count - 1 ? 0 : VariableDimentionlistBox.SelectedIndex + 1;
            }
            //Если нажат Шифт
            else if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == System.Windows.Input.Key.Enter)
                KmpsAppl.Doc.ST.UpdateOrMakeLine(double.Parse(DimVariableBox.Text.Replace('.', ',')), (bool)AddingLineToggle.IsChecked, false, null);

            //Контрол
            else if (Keyboard.Modifiers == ModifierKeys.Shift && e.Key == System.Windows.Input.Key.Enter)
            {
                AddingLineToggle.IsChecked = true;
                KmpsAppl.Doc.ST.UpdateOrMakeLine(double.Parse(DimVariableBox.Text.Replace('.', ',')), (bool)AddingLineToggle.IsChecked, false, KmpsAppl.Doc.GiveSelectOrChooseObj());
            }

            if (e.Key == System.Windows.Input.Key.Enter)
                DimVariableBox.Select(0, DimVariableBox.Text.Length);

            KmpsAppl.ZoomAll();
        }

        //Обновляем лист боксы с размерами
        private void RefreshDimBtn_Click(object sender, RoutedEventArgs e)
        {
            RefreshDimListBox();
        }

        private async void RefreshDimListBox()
        {
            AddingLineToggle.IsChecked = false;
            await KmpsAppl.Doc.ST.GetLineDimentionListAsync();
        }

        private void DimVariableBox_GotFocus(object sender, RoutedEventArgs e)
        {
            DimVariableBox.SelectAll();
        }

        private void ChangePositionItem(bool up)
        {
            if (VariableDimentionlistBox.SelectedItems != null)
            {
                int step = up ? -1 : 1;

                for (int i = up ? 0 : VariableDimentionlistBox.SelectedItems.Count - 1; up ? i < VariableDimentionlistBox.SelectedItems.Count : i >= 0; i -= step)
                {
                    SizeTool.ComboData data = (SizeTool.ComboData)VariableDimentionlistBox.SelectedItems[i];
                    if ((data.NumberPP + step >= 0) && (data.NumberPP + step < VariableDimentionlistBox.Items.Count))
                    {
                        //Смещаем выбранный объект в порядке
                        data.NumberPP += step;

                        //Смещаем объект что стоял перед/за ним
                        ((SizeTool.ComboData)VariableDimentionlistBox.Items[VariableDimentionlistBox.Items.IndexOf(data) + step]).NumberPP -= step;

                        //Обновляем коллекцию
                        CollectionViewSource.GetDefaultView(VariableDimentionlistBox.ItemsSource).Refresh();
                    }
                    else break;
                }
            }
        }

        private async void BTcombo_Initialized(object sender, EventArgs e)
        {
            BTcombo.DisplayMemberPath = "Name";
            BTcombo.SelectedValuePath = "Mac";
            var diveceList = await PairedList();
            BTcombo.ItemsSource = diveceList;

            BTcombo.SelectedValue = KompasCeilingMini.Properties.Settings.Default.BlueAddres;
        }

        //Запрос LE paired устройств
        private async Task<List<ComboBT>> PairedList()
        {
            List<ComboBT> bTs = new List<ComboBT>();
            DeviceInformationCollection PairedBluetoothDevices =
            await DeviceInformation.FindAllAsync(BluetoothLEDevice.GetDeviceSelector());

            foreach (DeviceInformation device in PairedBluetoothDevices)
                bTs.Add(new ComboBT(device.Name, device.Id));

            return await Task.Factory.StartNew(() =>
            {
                return bTs;
            });
        }

        private void BTcombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (BTcombo.SelectedValue != null)
            {
                KompasCeilingMini.Properties.Settings.Default.BlueAddres = BTcombo.SelectedValue.ToString();
                KompasCeilingMini.Properties.Settings.Default.Save();
            }
        }

        private void DimUpBtn_Click(object sender, RoutedEventArgs e)
        {
            ChangePositionItem(true);
        }

        private void DimDnBtn_Click(object sender, RoutedEventArgs e)
        {
            ChangePositionItem(false);
        }

        private void ConstraitLineToPointBtn_Click(object sender, RoutedEventArgs e)
        {
            KmpsAppl.Doc.ST.ObjectsToObjectDim();
        }

        private void SplitLineBtn_Click(object sender, RoutedEventArgs e)
        {
            if (KmpsAppl.Doc.ST != null)
            {
                KmpsAppl.someFlag = false;
                KmpsAppl.Doc.ST.SplitLine((double)SplitLineUpDn.Value, (bool)DellBaseDimChek.IsChecked);
                KmpsAppl.someFlag = true;
            }
        }

        private void ToPointBtn_Click(object sender, RoutedEventArgs e)
        {
            KmpsAppl.Doc.ST.ConnectLineToLine();
        }

        private void XInvertBtn_Click(object sender, RoutedEventArgs e)
        {
            KmpsAppl.Doc.ST.InvertPointCoord(true, KompasCeilingMini.Properties.Settings.Default.stg_text_size);
        }

        private void YInvertBtn_Click(object sender, RoutedEventArgs e)
        {
            KmpsAppl.Doc.ST.InvertPointCoord(false, KompasCeilingMini.Properties.Settings.Default.stg_text_size);
        }

        private async void FixDimBtn_Click(object sender, RoutedEventArgs e)
        {
           await KmpsAppl.Doc.ST.SetVariableToDim(false, ksDimensionTextBracketsEnum.ksDimSquareBrackets);
        }

        private async void UnfixDimBtn_Click(object sender, RoutedEventArgs e)
        {
           await KmpsAppl.Doc.ST.SetVariableToDim(true, ksDimensionTextBracketsEnum.ksDimBrackets);
        }

        private async void VariableDimBtn_Click(object sender, RoutedEventArgs e)
        {
            await KmpsAppl.Doc.ST.SetVariableToDim(false, ksDimensionTextBracketsEnum.ksDimBracketsOff);
        }
        #endregion


        #region setting
        private void KmpsCheckBox_Click(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.Save();
        }


        private void NameSetting_Click(object sender, RoutedEventArgs e)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                EditPanel.NameEditForm nameEdit = new EditPanel.NameEditForm(KmpsAppl.Doc., CeilingItems);
                nameEdit.Show();
            }
        }

        private void HeadSetting_Click(object sender, RoutedEventArgs e)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                EditPanel.HeadEditForm headEdit = new EditPanel.HeadEditForm(KmpsAppl.Doc., CeilingItems);
                headEdit.Show();
            }
        }


        private void JpgCheckBox_Click(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.Save();
        }

        #endregion

        private void HamburgerButton_Click(object sender, RoutedEventArgs e)
        {
            MenuDoc.Width = 150;
        }

        private void MenuButton2_Click(object sender, RoutedEventArgs e)
        {
            DimPanel.Visibility = Visibility.Collapsed;
            CalculatPanel.Visibility = Visibility.Visible;
            SettingPanel.Visibility = Visibility.Visible;
        }

        private void MenuButton1_Click(object sender, RoutedEventArgs e)
        {
            DimPanel.Visibility = Visibility.Visible;
            CalculatPanel.Visibility = Visibility.Visible;
            SettingPanel.Visibility = Visibility.Visible;
        }

        private void MenuButton3_Click(object sender, RoutedEventArgs e)
        {
            DimPanel.Visibility = Visibility.Collapsed;
            CalculatPanel.Visibility = Visibility.Collapsed;
            SettingPanel.Visibility = Visibility.Visible;
        }

        private void CloseMainBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void DockPanel_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void SizeCheks_Click(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.stg_size_auto = (bool)SizeCheks.IsChecked;
        }

        private void ObtuseAngleBtn_Click(object sender, RoutedEventArgs e)
        {
            SizeTool.ObtuseAngle();
        }

        private void PerimetrBox_ToolTipOpening(object sender, ToolTipEventArgs e)
        {
            PLineBlock.Text = "Прямые: " + KVariable.Sum("LineP", CeilingItems);
            PCurveBlock.Text = "Кривые: " + KVariable.Sum("CurveP", CeilingItems);
        }

        private void SetDimBtn_Click(object sender, RoutedEventArgs e)
        {
            KmpsAppl.Doc.ST.UpdateOrMakeLine(double.Parse(DimVariableBox.Text.Replace('.', ',')), (bool)AddingLineToggle.IsChecked, true, KmpsAppl.Doc.GiveSelectOrChooseObj());
        }

        private void CloseDocUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            KompasCeilingMini.Properties.Settings.Default.stg_closelast_val = (int)CloseDocUpDn.Value;
        }

        private void SizeText_ValueChanged_1(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            KompasCeilingMini.Properties.Settings.Default.stg_text_size = (int)SizeText.Value;
        }

        private void DopuskUpDN_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            KompasCeilingMini.Properties.Settings.Default.stg_dopuskUsadka = (int)DopuskUpDN.Value;
        }

        private void DistanceUPDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            KompasCeilingMini.Properties.Settings.Default.stg_crd_dopusk = (int)DistanceUPDn.Value;
        }

        private void SizeVariant_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (KompasCeilingMini.Properties.Settings.Default.stg_compensation_index != SizeVariant.SelectedIndex)
            {
                KompasCeilingMini.Properties.Settings.Default.stg_compensation_index = SizeVariant.SelectedIndex;
                KompasCeilingMini.Properties.Settings.Default.stg_compensation_val = int.Parse(((ComboBoxItem)SizeVariant.SelectedItem).Content.ToString());
                Comp = KompasCeilingMini.Properties.Settings.Default.stg_compensation_val;
                KompasCeilingMini.Properties.Settings.Default.Save();
            }

        }

        private void SizeVariant_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            SizeVariant.SelectedIndex = KompasCeilingMini.Properties.Settings.Default.stg_compensation_index;
        }

        private async void XUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (MainWindow.API != null)
            {
                await KVariable.UpdateAsync("koefX", XUpDn.Value.Value, string.Empty);
                WidthUpDn.Value = Math.Round(KVariable.Give("XGabarit", facturaListCheck.SelectedItem.ToString()) * (100 - XUpDn.Value.Value) / 100, 2);
            }
        }

        private async void YUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double?> e)
        {
            if (MainWindow.API != null)
            {
                await KVariable.UpdateAsync("koefY", YUpDn.Value.Value, string.Empty);
                HeightUpDn.Value = Math.Round(KVariable.Give("YGabarit", facturaListCheck.SelectedItem.ToString()) * (100 - YUpDn.Value.Value) / 100, 2);
            }
        }

        private void MinimezeMainBtn_Click(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void LockBtn_Click(object sender, RoutedEventArgs e)
        {
            KmpsDoc.LockedLayerAsync(88, true, true);
        }

        private void DellObjBtn_Click(object sender, RoutedEventArgs e)
        {
            KmpsAppl.Doc.DeleteSelectObj();
        }

        private void MoveViewBtn_Click(object sender, RoutedEventArgs e)
        {
            KmpsAppl.ZoomAll();
        }

        private void LockFrameBtn_Click(object sender, RoutedEventArgs e)
        {
            CheckResize(true);
        }
        private void CheckResize(bool invert = false)
        {
            if (invert) KompasCeilingMini.Properties.Settings.Default.head_resize = !KompasCeilingMini.Properties.Settings.Default.head_resize;
            if (KompasCeilingMini.Properties.Settings.Default.head_resize)
            {
                LockFrameBtn.Content = "&#xE785;";
                LockFrameBtn.Foreground = Brushes.Black;
                MainFrame.ResizeMode = ResizeMode.CanResize;
                KompasCeilingMini.Properties.Settings.Default.head_resize = true;
            }
            else
            {
                LockFrameBtn.Content = "&#xE72E;";
                LockFrameBtn.Foreground = Brushes.White;
                MainFrame.ResizeMode = ResizeMode.NoResize;
                KompasCeilingMini.Properties.Settings.Default.head_resize = false;
            }

            KompasCeilingMini.Properties.Settings.Default.Save();
        }

        private void ZoomOut_Click(object sender, RoutedEventArgs e)
        {
            IDocumentFrame documentFrame = (IDocumentFrame)KmpsAppl.Appl.ActiveDocument.DocumentFrames[0];
            documentFrame.ZoomPrevNextOrAll(ZoomTypeEnum.ksZoomNext);
        }

        private void ZoomIn_Click(object sender, RoutedEventArgs e)
        {
            IDocumentFrame documentFrame = (IDocumentFrame)KmpsAppl.Appl.ActiveDocument.DocumentFrames[0];
            documentFrame.ZoomPrevNextOrAll(ZoomTypeEnum.ksZoomPrevious);
        }
    }


}
