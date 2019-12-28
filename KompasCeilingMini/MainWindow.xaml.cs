using System.Data;
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
using System.Threading.Tasks;

using BlueTest.BlueClass;
using KompasLib.Tools;

namespace KompasCeilingMini
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        #region Custom declarations

        private KmpsAppl api;
        private BluetoothObserver observer;
        private SQLTool sqltool = new SQLTool();

        private bool clicado = false;
        private Point lm = new Point();

        public KmpsAppl API => api;

        #endregion Custom declarations

        public MainWindow()
        {
            InitializeComponent();
        }


        #region MainVoid
        //выбираем текущий домент
        internal void targetApi()
        {
            if (KmpsAppl.KompasAPI == null)
            {
                api = new KmpsAppl(KompasCeilingMini.Properties.Settings.Default.CloseLastVal, KompasCeilingMini.Properties.Settings.Default.CloseLast);
                api.ChangeDoc += new EventHandler<KmpsAppl>(ChangeDoc);
                api.ConnectBoolEvent += new EventHandler<bool>(ConnectEvent);

                if (KmpsAppl.KompasAPI != null)
                {
                    KmpStatLbl.Content = "Компас: ОК";
                }
                else
                {
                    MessageBox.Show(this, "Не найден активный объект", "Сообщение");
                    KmpStatLbl.Content = "Компас: Нет";
                }

                api.SelectDoc();

            }
            else
            {
                api.ChangeDoc += new EventHandler<KmpsAppl>(ChangeDoc);
                api.SelectDoc();
            }
            void ConnectEvent(object sender, bool e)
            {
                Dispatcher.Invoke(new Action(() =>
                {
                    if (e)
                    {

                    }
                    else
                    {
                        KmpStatLbl.Content = "Компас: Нет";
                    }
                }));
            }
        }

        private void ChangeDoc(object sender, KmpsAppl e)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                UpdateForm(true);
            }));
        }


        private void UpdateForm(bool First = true)
        {
            if((KmpsAppl.KompasAPI != null)&&(KmpsAppl.Appl.ActiveDocument != null))
            {
                if (KmpsAppl.Appl.ActiveDocument.Name != string.Empty)
                    SaveStatLabel.Content = KmpsAppl.Appl.ActiveDocument.Name;
                else SaveStatLabel.Content = "Новый документ";


                IVariable7 var;

                if (First == true)
                {
                    if (!api.Doc.D71.IsVariableNameValid("Number"))
                    {
                        facturaListCheck.Items.Clear();

                        var = api.Doc.D71.Variable[false, "Number"];
                        if (var.Note != "Number")
                        {
                            string[] numberList = var.Note.Split(';');
                            for (int i = 0; i < numberList.Length; i++) facturaListCheck.Items.Add(numberList[i]);
                            facturaListCheck.SelectedIndex = 0;
                        }
                        else
                        {
                            var.Note = "1";
                            api.Doc.D71.UpdateVariables();
                        }
                    }
                    else
                    {
                        facturaListCheck.Items.Clear();
                        api.Doc.D71.AddVariable("Number", 1, "Номера");
                        var = api.Doc.D71.Variable[false, "Number"];
                        var.Note = "1";
                        api.Doc.D71.UpdateVariables();
                        facturaListCheck.Items.Add("1");
                        facturaListCheck.SelectedIndex = 0;
                    }

                    mashtabChek.IsChecked = false;
                }

                string index = facturaListCheck.SelectedItem.ToString();


                //Number
                if (api.Doc.D71.IsVariableNameValid("Number")) api.Doc.D71.AddVariable("Number", 0, "Площадь");
                var = api.Doc.D71.Variable[false, "Number"];
                NumberUpDn.Value = (int)var.Value;
                //suffix
                var = api.Doc.D71.Variable[false, "Number"];
                sufixBox.Text = var.Note;

                //площадь
                if (api.Doc.D71.IsVariableNameValid("Sqare" + index)) api.Doc.D71.AddVariable("Sqare" + index, 0, "Площадь");
                var = api.Doc.D71.Variable[false, "Sqare" + index];
                SqareBox.Text = var.Value.ToString();

                //площадь с усадкой
                if (api.Doc.D71.IsVariableNameValid("SqareU" + index)) api.Doc.D71.AddVariable("SqareU" + index, 0, "ПлощадьУ");
                var = api.Doc.D71.Variable[false, "SqareU" + index];
                SqareUBox.Text = var.Value.ToString();

                //периметр
                if (api.Doc.D71.IsVariableNameValid("Perimetr" + index)) api.Doc.D71.AddVariable("Perimetr" + index, 0, "Perimetr");
                var = api.Doc.D71.Variable[false, "Perimetr" + index];
                PerimetrBox.Text = var.Value.ToString();

                //периметрУ
                if (api.Doc.D71.IsVariableNameValid("PerimetrU" + index)) api.Doc.D71.AddVariable("PerimetrU" + index, 0, "PerimetrU");
                //периметр прямых
                if (api.Doc.D71.IsVariableNameValid("LineP" + index)) api.Doc.D71.AddVariable("LineP" + index, 0, "LineP");
                var = api.Doc.D71.Variable[false, "LineP" + index];

                //периметр кривых
                if (api.Doc.D71.IsVariableNameValid("CurveP" + index)) api.Doc.D71.AddVariable("CurveP" + index, 0, "CurveP");
                var = api.Doc.D71.Variable[false, "CurveP" + index];

                //углы
                if (api.Doc.D71.IsVariableNameValid("Angle")) api.Doc.D71.AddVariable("Angle", 0, "Angle");
                var = api.Doc.D71.Variable[false, "Angle"];
                AngUpDn.Value = (byte)var.Value;

                //шов
                if (api.Doc.D71.IsVariableNameValid("Shov" + index)) api.Doc.D71.AddVariable("Shov" + index, 0, "Shov");
                var = api.Doc.D71.Variable[false, "Shov" + index];
                ShovBox.Text = var.Value.ToString();


                //XGabarit
                if (api.Doc.D71.IsVariableNameValid("XGabarit" + index)) api.Doc.D71.AddVariable("XGabarit" + index, 0, "XGabarit");
                var = api.Doc.D71.Variable[false, "XGabarit" + index];
                WidthUpDn.Value = (var.Value * ((100 - api.Doc.Variable.Give("koefX", string.Empty)) / 100));

                //YGabarit
                if (api.Doc.D71.IsVariableNameValid("YGabarit" + index)) api.Doc.D71.AddVariable("YGabarit" + index, 0, "YGabarit");
                var = api.Doc.D71.Variable[false, "YGabarit" + index];
                HeightUpDn.Value = (var.Value * ((100 - api.Doc.Variable.Give("koefY", string.Empty)) / 100));

                //Xcrd
                if (api.Doc.D71.IsVariableNameValid("Xcrd" + index)) api.Doc.D71.AddVariable("Xcrd" + index, 0, "Xcrd");
                var = api.Doc.D71.Variable[false, "Xcrd" + index];

                //Ycrd
                if (api.Doc.D71.IsVariableNameValid("Ycrd" + index)) api.Doc.D71.AddVariable("Ycrd" + index, 0, "Ycrd");
                var = api.Doc.D71.Variable[false, "Ycrd" + index];


                //X
                if (api.Doc.D71.IsVariableNameValid("koefX")) api.Doc.D71.AddVariable("koefX", 7, "Unchecked");
                var = api.Doc.D71.Variable[false, "koefX"];
                XUpDn.Value = var.Value;
                if (var.Note == "koefX") api.Doc.Variable.UpdateNote("koefX", "Unchecked", "");
                //Статус усадки;
                if (First)
                {
                    if (var.Note == "Checked") mashtabChek.IsChecked = true;
                    else mashtabChek.IsChecked = false;
                }

                //Y
                if (api.Doc.D71.IsVariableNameValid("koefY")) api.Doc.D71.AddVariable("koefY", 7, "koefY");
                var = api.Doc.D71.Variable[false, "koefY"];
                YUpDn.Value = var.Value;


                //фотопечать
                if (api.Doc.D71.IsVariableNameValid("photo" + index)) api.Doc.D71.AddVariable("photo" + index, 0, "photo");
                var = api.Doc.D71.Variable[false, "photo" + index];
                FpCalc.Value = (decimal)var.Value;

                //длина
                if (api.Doc.D71.IsVariableNameValid("lenth" + index)) api.Doc.D71.AddVariable("lenth" + index, 0, "lenth");
                var = api.Doc.D71.Variable[false, "lenth" + index];

                //фактура
                if (api.Doc.D71.IsVariableNameValid("factura" + index)) api.Doc.D71.AddVariable("factura" + index, -1, "factura");
                var = api.Doc.D71.Variable[false, "factura" + index];

                double sizeX = api.Doc.Variable.Give("XGabarit", index) * ((100 - api.Doc.Variable.Give("koefX", string.Empty)) / 100);
                double sizeY = api.Doc.Variable.Give("YGabarit", index) * ((100 - api.Doc.Variable.Give("koefY", string.Empty)) / 100);
                SQLFacturaRefresh("SELECT IDFactura, CONCAT(Name,' ',Width) AS NameW FROM dbo.Factura WHERE Width >=" + Math.Min(sizeX*0.97, sizeY*0.97).ToString().Replace(',', '.') + " ORDER BY NumberPP");
                if (FacturaCombo.Items.Count == 0)
                SQLFacturaRefresh("SELECT IDFactura, CONCAT(Name,' ',Width) AS NameW FROM dbo.Factura ORDER BY NumberPP");


                //цвет
                if (api.Doc.D71.IsVariableNameValid("color" + index)) api.Doc.D71.AddVariable("color" + index, -1, "color");
                var = api.Doc.D71.Variable[false, "color" + index];
                SQLFacturaColorRefresh();

                if (var.Note == "color" + index) api.Doc.Variable.UpdateNote("color", ColorCombo.Text, index);

                //вырез
                if (api.Doc.D71.IsVariableNameValid("cut" + index)) api.Doc.D71.AddVariable("cut" + index, 0, "вырез");
                var = api.Doc.D71.Variable[false, "cut" + index];
                CutBox.Text = var.Value.ToString();

                //Комментарий 1
                if (api.Doc.D71.IsVariableNameValid("Comment1")) api.Doc.D71.AddVariable("Comment1", 0, "Comment1");
                var = api.Doc.D71.Variable[false, "Comment1"];
                paramBox_1.Text = var.Note;

                //Комментарий 2
                if (api.Doc.D71.IsVariableNameValid("Comment2")) api.Doc.D71.AddVariable("Comment2", 0, "Comment2");
                var = api.Doc.D71.Variable[false, "Comment2"];
                paramBox_2.Text = var.Note;





            }
        }

        private void SQLFacturaRefresh(string sqlcommand)
        {
            api.someFlag = false;
            FacturaCombo.ItemsSource = null;
            FacturaCombo.Items.Clear();

            DataSet ds = sqltool.ComboDataSet(sqlcommand, "Factura", KompasCeilingMini.Properties.Settings.Default.EasyCeilingConnectionString);

            FacturaCombo.DisplayMemberPath = "NameW";
            FacturaCombo.SelectedValuePath = "IDFactura";
            FacturaCombo.ItemsSource = ds.Tables["Factura"].DefaultView;

            FacturaCombo.SelectedValue = api.Doc.Variable.Give("factura", facturaListCheck.SelectedItem.ToString());

            api.someFlag = true;
            if (ColorCombo.SelectedIndex < -1)
                FacturaCombo.SelectedIndex = 0;
        }

        private void SQLFacturaColorRefresh()
        {
            if (FacturaCombo.SelectedIndex > -1)
            {


                ColorCombo.ItemsSource = null;
                api.someFlag = false;
                ColorCombo.Items.Clear();

                DataSet ds = sqltool.ComboDataSet("SELECT IDFacCol, Name FROM dbo.FactruaColor WHERE IDFactura=" + FacturaCombo.SelectedValue + " ORDER BY NumberPP", "FactruaColor", KompasCeilingMini.Properties.Settings.Default.EasyCeilingConnectionString);


                ColorCombo.DisplayMemberPath = "Name";
                ColorCombo.SelectedValuePath = "IDFacCol";
                ColorCombo.ItemsSource = ds.Tables["FactruaColor"].DefaultView;


                ColorCombo.SelectedValue = api.Doc.Variable.Give("color", facturaListCheck.SelectedItem.ToString());

                api.someFlag = true;

                if (ColorCombo.SelectedIndex < 0)
                    ColorCombo.SelectedIndex = 0;
            }

        }

        #endregion

        //Отправляет на образмеривание

        private void NewFileVoid()
        {
            if (api != null)
            {
                if (!api.CreateDoc())
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

            if (KompasCeilingMini.Properties.Settings.Default.WorkFolder == null || KompasCeilingMini.Properties.Settings.Default.WorkFolder == "")
            {
                System.Windows.Forms.FolderBrowserDialog WorkFolderSlct = new System.Windows.Forms.FolderBrowserDialog();

                if (WorkFolderSlct.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    KompasCeilingMini.Properties.Settings.Default.WorkFolder = WorkFolderSlct.SelectedPath;
                    KompasCeilingMini.Properties.Settings.Default.Save();
                }
            }

            openFileDialog.InitialDirectory = KompasCeilingMini.Properties.Settings.Default.WorkFolder;
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
                            api.Doc.D5 = (ksDocument2D)KmpsAppl.KompasAPI.Document2D();
                            if (api.Doc.D5 != null)
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

        #region CalcAll
        private void GetContour(string index, bool Usadka = false, bool cursor = true, reference refContour = 0)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                #region получение контура
                double x = 0, y = 0;
                RequestInfo info = (RequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);


                //Ищем или находим макрообъект по индексу потолка
                IMacroObject macroObject = api.Doc.Macro.FindCeilingMacro(index);
                if (macroObject == null) macroObject = api.Doc.Macro.MakeCeilingMacro(index);
                else if (cursor)
                {
                    api.Doc.Macro.RemoveCeilingMacro(index);
                    macroObject = api.Doc.Macro.MakeCeilingMacro(index);
                }

                if (refContour == 0)
                {
                    if (cursor)
                    {
                        api.Doc.D5.ksCursor(info, ref x, ref y, 0);
                        int refEncloseContours = api.Doc.D5.ksMakeEncloseContours(0, x, y);
                        if (refEncloseContours != 0)
                            api.Doc.Macro.AddCeilingMacro(refEncloseContours, index);  //Добавляем ksMakeEncloseContours
                        else
                            KmpsAppl.KompasAPI.ksMessage("Не найден замкнутый контур");

                        refContour = api.Doc.Macro.GiveRefFromMacro(macroObject.Reference, "F");
                    }
                    else
                    {
                        refContour = api.Doc.Macro.GiveRefFromMacro(macroObject.Reference, "F");
                    }
                }
                else
                {
                    api.Doc.Macro.RemoveCeilingMacro(index);
                    macroObject = api.Doc.Macro.MakeCeilingMacro(index);
                    api.Doc.Macro.AddCeilingMacro(refContour, index);
                    refContour = api.Doc.Macro.GiveRefFromMacro(macroObject.Reference, "F");
                }
                #endregion
                if (refContour != 0)
                    CalcAll(macroObject, Usadka);  //Считает все
                UpdateForm(false);
            }
            else MessageBox.Show(this, "Объект не захвачен", "Сообщение");

            void CalcAll(IMacroObject macroObject1, bool usadka)
            {
                if (usadka)
                {
                    api.Doc.Variable.Update("PerimetrU", 0, index);
                    if (facturaListCheck.SelectedIndex == 0) api.Doc.Variable.Update("Angle", 0, string.Empty);
                }
                else
                {
                    api.Doc.Variable.Update("Perimetr", 0, index);
                    api.Doc.Variable.Update("LineP", 0, index);
                    api.Doc.Variable.Update("CurveP", 0, index);
                    api.Doc.Variable.Update("Shov", 0, index);
                    api.Doc.Variable.Update("cut", 0, index);
                    if (facturaListCheck.SelectedIndex == 0) api.Doc.Variable.Update("Angle", 0, string.Empty);
                }

                ksInertiaParam inParam = (ksInertiaParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_InertiaParam);
                ksIterator Iterator1 = (ksIterator)KmpsAppl.KompasAPI.GetIterator();
                Iterator1.ksCreateIterator(ldefin2d.ALL_OBJ, macroObject1.Reference);
                double SqareMain = 0;
                int position = facturaListCheck.Items.IndexOf(index);
                int MainRef = 0;
                int Count = 0;
                //Перебираем кривые в поиске самой большой.
                reference refContour1 = Iterator1.ksMoveIterator("F");
                while (refContour1 != 0)
                {
                    api.Mat.ksCalcInertiaProperties(refContour1, inParam, 3);
                    if (SqareMain < inParam.F)
                    {
                        SqareMain = inParam.F; //Выбираем самую большую
                        MainRef = refContour1;
                    }
                    Count++;
                    refContour1 = Iterator1.ksMoveIterator("N");
                }
                //
                //Получить габарит mainRef
                //
                MakeGabarit(MainRef, (bool)NoMashtab.IsChecked);

                //
                //Начинаем перебор контуров со всем что есть
                //
                //Заходим в первый контур


                refContour1 = Iterator1.ksMoveIterator("F");
                while (refContour1 != 0)
                {

                    double lineTemp = 0, curveTemp = 0, shovTemp = 0, cut = 0, angleTemp = api.Doc.Variable.Give("Angle", string.Empty);
                    IContour contour1 = api.Doc.Macro.GiveContour(refContour1);
                    if (refContour1 != MainRef)
                    {
                        api.Mat.ksCalcInertiaProperties(refContour1, inParam, 3);   //Если он не главный, то вычитаем его площадь из главного.
                        SqareMain -= inParam.F;                                 //Вычитаем из главной площади
                        perriContour(contour1, ref cut, ref cut, ref angleTemp);               //Считаем предполагаем вырез
                    }
                    else perriContour(contour1, ref lineTemp, ref curveTemp, ref angleTemp);
                    //Функция расчета периметра
                    void perriContour(IContour contour, ref double Line, ref double Curve, ref double Angle)
                    {
                        if (contour != null)
                        {
                            for (int i = 0; i < contour.Count; i++)
                            {
                                IContourSegment pDrawObj = (IContourSegment)contour.Segment[i];
                                // Получить тип объекта
                                // В зависимости от типа вывести сообщение для данного типа объектов
                                try
                                {
                                    IContourLineSegment contourLineSegment = (IContourLineSegment)pDrawObj;
                                    Line += contourLineSegment.Length;
                                    Angle += 1;
                                }
                                catch
                                {
                                    ICurve2D contourLineSegment = (ICurve2D)pDrawObj.Curve2D;
                                    Curve += contourLineSegment.Length;
                                    if (i > 0)
                                        if (contour.Segment[i - 1].Type == KompasAPIObjectTypeEnum.ksObjectContourLineSegment)
                                            Angle += 1;
                                }
                            }
                        }
                    }


                    api.ProgressBar.Start(0, facturaListCheck.Items.Count, "Считаем контур ", true);

                    for (int i = 0; i < facturaListCheck.Items.Count; i++)           //Стартуем проход по потолкам
                    {

                        api.ProgressBar.SetProgress(i, "Считаем контур " + i, false);
                        if (i == position)
                            continue;                            //Если это тот же потолок то выходим.
                        else
                        {
                            IMacroObject macroObject2 = api.Doc.Macro.FindCeilingMacro(facturaListCheck.Items[i].ToString());
                            if (macroObject2 != null)
                            {
                                ksIterator Iterator2 = (ksIterator)KmpsAppl.KompasAPI.GetIterator();
                                Iterator2.ksCreateIterator(ldefin2d.ALL_OBJ, macroObject2.Reference);

                                //Заходим во второй контур
                                reference refContour2 = Iterator2.ksMoveIterator("F");
                                while (refContour2 != 0)
                                {
                                    IContour contour2 = api.Doc.Macro.GiveContour(refContour2);
                                    ksDynamicArray arrayCurve = (ksDynamicArray)KmpsAppl.KompasAPI.GetDynamicArray(ldefin2d.POINT_ARR);
                                    if (api.Mat.ksIntersectCurvCurv(contour1.Reference, contour2.Reference, arrayCurve) == 1)
                                    {
                                        angleTemp -= 1;
                                        int step = (int)arrayCurve.ksGetArrayCount();
                                        ksMathPointParam[] points = new ksMathPointParam[step];

                                        for (int j = 0; j < step; j++)
                                        {
                                            points[j] = (ksMathPointParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_MathPointParam);
                                            arrayCurve.ksGetArrayItem(j, points[j]);
                                        }

                                        double lenthTemp = 0;

                                        //Сравниваем сегменты контура с другим контуром (по сегментно)
                                        for (int j = 0; j < contour1.Count; j++)
                                        {
                                            IContourSegment segment = (IContourSegment)contour1.Segment[j];

                                            for (int k = 0; k < contour2.Count; k++)
                                            {
                                                IContourSegment segment2 = (IContourSegment)contour2.Segment[k];

                                                //Получаем крайние точки пересечения
                                                double[] intersecArr = segment.Curve2D.Intersect(segment2.Curve2D);
                                                if (intersecArr != null)
                                                    if (intersecArr.Length > 3)
                                                    {
                                                        //Узнаем длинну
                                                        double lenthTemp2 = segment.Curve2D.GetDistancePointPoint(intersecArr[0], intersecArr[1], intersecArr[2], intersecArr[3]);
                                                        lenthTemp += lenthTemp2;

                                                        //Если она больше, то нет смысла сравнивать дальше.
                                                        if (lenthTemp2 >= segment.Curve2D.Length)
                                                            break;
                                                    }
                                            }
                                        }


                                        shovTemp += lenthTemp;

                                        arrayCurve.ksDeleteArray();
                                        if (i < position)
                                        {
                                            if (Math.Round(curveTemp - shovTemp, 2) < 0) { lineTemp += curveTemp - shovTemp; curveTemp = 0; }
                                            else curveTemp -= shovTemp;
                                            shovTemp = 0;
                                        }
                                    }
                                    refContour2 = Iterator2.ksMoveIterator("N"); //Двигаем итератор 2

                                    arrayCurve.ksDeleteArray();
                                }
                                Iterator2.ksDeleteIterator(); //Удаляем итератор 2
                            }
                        }
                    }
                    //получили шов и вычитаем его из предыдущих параметров периметра или длинны выреза.

                    shovTemp = Math.Round(shovTemp, 2);
                    curveTemp = Math.Round(curveTemp, 2);
                    lineTemp = Math.Round(lineTemp, 2);

                    if (refContour1 != MainRef) cut -= shovTemp;
                    else
                    {
                        if (curveTemp - shovTemp <= 0) { lineTemp += curveTemp - shovTemp; curveTemp = 0; }
                        else curveTemp -= shovTemp;
                    }

                    if (usadka)
                    {
                        api.Doc.Variable.Add("PerimetrU", Math.Round((lineTemp + curveTemp) / 100, 2), index);
                        api.Doc.Variable.Add("Angle", angleTemp, string.Empty);
                    }
                    else
                    {
                        api.Doc.Variable.Add("Perimetr", Math.Round((lineTemp + curveTemp) / 100, 2), index);
                        api.Doc.Variable.Add("LineP", Math.Round(lineTemp / 100, 2), index);
                        api.Doc.Variable.Add("CurveP", Math.Round(curveTemp / 100, 2), index);
                        api.Doc.Variable.Add("Shov", Math.Round(shovTemp / 100, 2), index);
                        api.Doc.Variable.Add("cut", Math.Round(cut / 100, 2), index);
                        api.Doc.Variable.Add("Angle", angleTemp, string.Empty);
                    }

                    refContour1 = Iterator1.ksMoveIterator("N"); //Двигаем итератор 1
                }
                Iterator1.ksDeleteIterator(); //Удаляем итератор 1 после полного перебора
                if (usadka) api.Doc.Variable.Update("SqareU", Math.Round(SqareMain / 10000, 2), index);
                else api.Doc.Variable.Update("Sqare", Math.Round(SqareMain / 10000, 2), index);

                api.ProgressBar.Stop("Закончили", true);

                void MakeGabarit(reference objRef, bool garpun)
                {
                    ksRectangleParam recPar = (ksRectangleParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectangleParam);
                    ksRectParam spcGabarit = (ksRectParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RectParam);
                    if (api.Doc.D5.ksGetObjGabaritRect(objRef, spcGabarit) == 1)
                    {
                        ksMathPointParam mathBop = spcGabarit.GetpBot();
                        ksMathPointParam mathTop = spcGabarit.GetpTop();
                        double x = mathBop.x;
                        double y = mathBop.y;
                        double dx = mathTop.x;
                        double dy = mathTop.y;
                        double sizeX = Math.Round(Math.Abs(x - dx), 2);
                        double sizeY = Math.Round(Math.Abs(y - dy), 2);

                        if (FacturaCombo.SelectedIndex == -1) FacturaCombo.SelectedIndex = 0;

                        double width = double.Parse(SqlReturnValue("SELECT TOP 1 Width FROM dbo.Factura WHERE IDFactura=" + api.Doc.Variable.Give("factura", index), "Width"));

                        api.Doc.Variable.Update("Xcrd", x, index);
                        api.Doc.Variable.Update("Ycrd", y, index);

                        if (!usadka)
                        {
                            api.Doc.Variable.Update("XGabarit", sizeX, index);
                            api.Doc.Variable.Update("YGabarit", sizeY, index);
                        }

                        if (usadka || !garpun)
                        {
                            IVariable7 lenth = api.Doc.D71.Variable[false, "lenth" + index];

                            recPar.Init();
                            recPar.x = x;
                            recPar.y = y;
                            double Dopusk = (double)usadkaDopusk.Value / 100 + 1;

                            if ((sizeX / width) <= 1 * Dopusk && (sizeX / width) > (sizeY / width) || (sizeY > width * Dopusk))
                            {
                                lenth.Value = sizeY;
                                if (y < dy) recPar.height = sizeY;
                                else recPar.height = -sizeY;
                                if (x < dx) recPar.width = width;
                                else recPar.width = -width;
                            }
                            else
                            {
                                lenth.Value = sizeX;
                                if (y < dy) recPar.height = width;
                                else recPar.height = -width;
                                if (x < dx) recPar.width = sizeX;
                                else recPar.width = -sizeX;
                            }

                            api.Doc.D71.UpdateVariables();
                        }
                        // api.Doc.D5.ksRectangle(recPar, 0);

                        //SQLFacturaRefresh("SELECT IDFactura, CONCAT(Name,' ',Width) AS NameW FROM dbo.Factura WHERE Width >=" + Math.Min(sizeX, sizeY).ToString().Replace(',','.') + " ORDER BY NumberPP");
                    }
                }
            }

        }
        #endregion


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
            {
                if (MessageBox.Show("Новый месяц!\n Сменим папку?", "Внимание", MessageBoxButton.OKCancel, MessageBoxImage.Question) == MessageBoxResult.OK)
                {
                    System.Windows.Forms.FolderBrowserDialog WorkFolderSlct = new System.Windows.Forms.FolderBrowserDialog();

                    if (WorkFolderSlct.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        KompasCeilingMini.Properties.Settings.Default.WorkFolder = WorkFolderSlct.SelectedPath;
                        KompasCeilingMini.Properties.Settings.Default.lastDate = date.Month;
                        KompasCeilingMini.Properties.Settings.Default.Save();
                    }
                }
            }
            targetApi();
        }


        private void FacturaListCheck_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (facturaListCheck.Items.Count > 0)
                if (KmpsAppl.KompasAPI != null)
                {
                    if (facturaListCheck.SelectedIndex < 0) facturaListCheck.SelectedIndex = 0;

                    IVariable7 var = api.Doc.D71.Variable[false, "Number"];
                    var.Note = GiveNumberStroke();
                    api.Doc.D71.UpdateVariables();
                    UpdateForm(false);

                    //Возвращает строку всех фактуры
                    string GiveNumberStroke()
                    {
                        string number = facturaListCheck.Items[0].ToString();
                        for (int i = 1; i < facturaListCheck.Items.Count; i++) number += ";" + facturaListCheck.Items[i].ToString();
                        return number;
                    }
                }
        }

        private void KsContrForm_Closed(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void FacturaCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if ((api.someFlag) && (KmpsAppl.KompasAPI != null) && (FacturaCombo.SelectedIndex > -1))
            {
                api.Doc.Variable.Update("factura", double.Parse(((DataRowView)FacturaCombo.SelectedItem).Row["IDFactura"].ToString()), facturaListCheck.SelectedItem.ToString());
                api.Doc.Variable.UpdateNote("factura", ((DataRowView)FacturaCombo.SelectedItem).Row["NameW"].ToString(), facturaListCheck.SelectedItem.ToString());
                SQLFacturaColorRefresh();
                if (ColorCombo.SelectedIndex > -1)
                {
                    api.Doc.Variable.Update("color", double.Parse(((DataRowView)ColorCombo.SelectedItem).Row["IDFacCol"].ToString()), facturaListCheck.SelectedItem.ToString());
                    api.Doc.Variable.UpdateNote("color", ((DataRowView)ColorCombo.SelectedItem).Row["Name"].ToString(), facturaListCheck.SelectedItem.ToString());
                }
            }
        }


        private void PlusFacturaBtn_Copy_Click(object sender, RoutedEventArgs e)
        {
            if (facturaListCheck.Items.Count > 1)
            {
                api.Doc.Variable.Clear(facturaListCheck.SelectedItem.ToString());
                api.Doc.Macro.RemoveCeilingMacro(facturaListCheck.SelectedItem.ToString());
                facturaListCheck.Items.Remove(facturaListCheck.SelectedItem);
            }
        }

        private void PlusFacturaBtn_Click(object sender, RoutedEventArgs e)
        {
            if (api.Doc.D71 != null)
            {
                facturaListCheck.Items.Add(facturaListCheck.Items.Count + 1);

                {
                    IVariable7 var = api.Doc.D71.Variable[false, "Number"];
                    var.Note = GiveNumberStroke();
                    api.Doc.D71.UpdateVariables();
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


        private void MashtabChek_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.RightButton == System.Windows.Input.MouseButtonState.Pressed)
                mashtabChek.IsChecked = !mashtabChek.IsChecked;
            else if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
            {

            }
        }

        private void SizeCheks_Click(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.SetSize = (bool)SizeCheks.IsChecked;
            KompasCeilingMini.Properties.Settings.Default.Save();
        }

        private void CalcAllBtn_Click(object sender, RoutedEventArgs e)
        {
            if (KmpsAppl.KompasAPI != null)
            {
                if (facturaListCheck.SelectedIndex > 0) ReCalcBtn.Background = Brushes.Yellow;
                GetContour(facturaListCheck.SelectedItem.ToString(), (bool)mashtabChek.IsChecked);
                KmpsAppl.Appl.StopCurrentProcess();
            }
            else
            {
                MessageBox.Show(this, "Объект не захвачен", "Сообщение");
            }
        }

        //выбрать рабочую папку
        private void WorkFolder_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog WorkFolderSlct = new System.Windows.Forms.FolderBrowserDialog();
            if (WorkFolderSlct.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                KompasCeilingMini.Properties.Settings.Default.WorkFolder = WorkFolderSlct.SelectedPath;
                KompasCeilingMini.Properties.Settings.Default.Save();
            }
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

            if ((SqareUBox.Text == string.Empty) || (SqareUBox.Text == "0"))
            {
                MessageBox.Show("Не проставленна площадь по усадке (красным)");
                SqareUBox.Background = Brushes.Pink;
                return;
            }

            if (api != null)
            {
                ItemCollection IC = facturaListCheck.Items;
                string name = string.Empty;

                //Если номера не совпадают, то предлагаем сохранить с новым
                if (NumberUpDn.Value != api.Doc.Variable.Give("Number", string.Empty))
                {
                    if (MessageBox.Show("Сохраняем с новым номером " + NumberUpDn.Text + "?", "Внимание", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                    {
                        KompasCeilingMini.Properties.Settings.Default.ZkzLastNumber = (int)NumberUpDn.Value;
                        KompasCeilingMini.Properties.Settings.Default.Save();
                        name = api.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, facturaListCheck.Items, KompasCeilingMini.Properties.Settings.Default.sufix, string.Empty);
                    }

                    else return;
                }
                //Если имя пустое а мы сохраняем компас согласно настройкам
                else if ((api.Doc.D7.Name != string.Empty) && ((bool)kmpsCheckBox.IsChecked == true))
                {
                    //Если имя равно генирруемому
                    if (api.Doc.D7.Name.Substring(0, api.Doc.D7.Name.Length - 4) == api.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, IC, KompasCeilingMini.Properties.Settings.Default.sufix, string.Empty))
                        //То используем как есть
                        name = api.Doc.D7.Name.Substring(0, api.Doc.D7.Name.Length - 4);
                    else
                    {
                        //Нет — делаем новое.
                        oldName = api.Doc.D7.PathName;
                        name = api.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, IC, KompasCeilingMini.Properties.Settings.Default.sufix, string.Empty);
                    }
                }

                else
                {
                    //Если текущий номер уже сохранян как использованный
                    if ((int)NumberUpDn.Value <= KompasCeilingMini.Properties.Settings.Default.ZkzLastNumber)
                    {
                        if (MessageBox.Show("Вы уже использовали номер " + NumberUpDn.Text.Split('/', '-')[0] + ", или меньше\nДалее?", "Внимание", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                        {
                            KompasCeilingMini.Properties.Settings.Default.ZkzLastNumber = (int)NumberUpDn.Value;
                            KompasCeilingMini.Properties.Settings.Default.Save();
                            name = api.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, IC, KompasCeilingMini.Properties.Settings.Default.sufix, string.Empty);
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
                        name = api.Doc.GiveMeText(KompasCeilingMini.Properties.Settings.Default.NameFileText, IC, KompasCeilingMini.Properties.Settings.Default.sufix, string.Empty);
                    }
                }
                string path = (KompasCeilingMini.Properties.Settings.Default.WorkFolder + "\\" + name);
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
                    api.ProgressBar.Start(0, 3, "Сохранение", true);
                    ksRasterFormatParam rasParam = api.Doc.D5.RasterFormatParam();
                    //jpeg парам

                    if (rasParam != null)
                    {
                        rasParam.format = 2;
                        rasParam.extResolution = 30;
                    }
                    //frw Компас
                    if ((api.Doc != null) && (menu == false))
                    {
                        api.ProgressBar.SetProgress(1, "Пишем компас", true);

                        ksDocumentParam documentParam = (ksDocumentParam)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_DocumentParam);

                        api.Doc.D5.ksGetObjParam(api.Doc.D5.reference, documentParam, ldefin2d.ALLPARAM);
                        documentParam.author = Environment.UserName;
                        documentParam.comment = "KMCm";
                        api.Doc.D5.ksSetObjParam(api.Doc.D5.reference, documentParam, ldefin2d.ALLPARAM);


                        if (kmpsCheckBox.IsChecked == true) api.Doc.D5.ksSaveDocument(path + ".frw");

                        SaveStatLabel.Content = "Сохранили";
                        if (JpgCheckBox.IsChecked == true) api.Doc.D5.SaveAsToRasterFormat(KompasCeilingMini.Properties.Settings.Default.WorkFolder + "\\jpg\\" + name + ".jpg", rasParam);

                    }


                    if (oldName != string.Empty)
                    {
                        if (File.Exists(oldName)) File.Delete(oldName);
                        if (File.Exists(oldName + ".bak")) File.Delete(oldName + ".bak");
                        if (File.Exists(oldName.Replace(".frw", ".ini"))) File.Delete(oldName.Replace(".frw", ".ini"));
                    }

                    UpdateForm(false);

                    api.ProgressBar.Stop("Сохранили", true);
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
                ReCalcBtn.Background = Brushes.YellowGreen;
                for (int i = 0; i < facturaListCheck.Items.Count; i++)
                {
                    GetContour(facturaListCheck.Items[i].ToString(), (bool)mashtabChek.IsChecked, false);
                }
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

        private void KmpsCheckBox_Click(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.Save();
        }

        private void JpgCheckBox_Click(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.Save();
        }

        private void AngUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (api != null)
                api.Doc.Variable.Update("Angle", (double)AngUpDn.Value, string.Empty);
        }

        private void SaveBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveMe();
        }

        private void HeadBtn_Click(object sender, RoutedEventArgs e)
        {
            if (api != null)
            {
                double x = 0, y = 0;
                double width = api.Doc.Variable.Give("XGabarit", "1");
                ItemCollection IC = facturaListCheck.Items;
                ksRequestInfo info = (ksRequestInfo)KmpsAppl.KompasAPI.GetParamStruct((short)StructType2DEnum.ko_RequestInfo);

                int j = api.Doc.D5.ksCursor(info, ref x, ref y, 0);

                api.Doc.CreateText(KompasCeilingMini.Properties.Settings.Default.HeadText, x, y, width, facturaListCheck.Items, KompasCeilingMini.Properties.Settings.Default.sufix, (bool)AutoSizeTextChek.IsChecked, false);

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

            openFileDialog.InitialDirectory = KompasCeilingMini.Properties.Settings.Default.WorkFolder;
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
                            if (point[i] != "") api.Doc.D5.ksPoint(double.Parse(point[i].Split(',')[0].Replace('.', ',')) / 10, double.Parse(point[i].Split(',')[1].Replace('.', ',')) / 10, 1);
                        }
                    }
                    else
                    {
                        for (int i = 1; i < point.Length; i++)
                        {
                            if (point[i] != "")
                            {
                                double X1, Y1, X2, Y2;
                                X1 = Math.Round(double.Parse(point[i].Split(',')[0].Replace('.', ',')) / 10, 4);
                                Y1 = -Math.Round(double.Parse(point[i].Split(',')[1].Replace('.', ',')) / 10, 4);
                                X2 = Math.Round(double.Parse(point[i - 1].Split(',')[0].Replace('.', ',')) / 10, 4);
                                Y2 = -Math.Round(double.Parse(point[i - 1].Split(',')[1].Replace('.', ',')) / 10, 4);

                                api.Doc.D5.ksLineSeg(X1, Y1, X2, Y2, 2);
                                //api.Doc.D5.ksEndObj();
                            }
                            else
                            {
                                double X1, Y1, X2, Y2;
                                X1 = Math.Round(double.Parse(point[point.Length - 2].Split(',')[0].Replace('.', ',')) / 10, 4);
                                Y1 = -Math.Round(double.Parse(point[point.Length - 2].Split(',')[1].Replace('.', ',')) / 10, 4);
                                X2 = Math.Round(double.Parse(point[0].Split(',')[0].Replace('.', ',')) / 10, 4);
                                Y2 = -Math.Round(double.Parse(point[0].Split(',')[1].Replace('.', ',')) / 10, 4);
                                api.Doc.D5.ksLineSeg(X1, Y1, X2, Y2, 2);
                                //api.Doc.D5.ksEndObj();
                            }
                        }
                    }
                }
                else MessageBox.Show("Что-то пошло не так.");

            }


        }

        private void ColorCombo_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if ((api.someFlag) && (ColorCombo.SelectedIndex > -1))
            {
                api.Doc.Variable.Update("color", double.Parse(((DataRowView)ColorCombo.SelectedItem).Row["IDFacCol"].ToString()), facturaListCheck.SelectedItem.ToString());
                api.Doc.Variable.UpdateNote("color", ((DataRowView)ColorCombo.SelectedItem).Row["Name"].ToString(), facturaListCheck.SelectedItem.ToString());
            }
        }

        private void HeadSetting_Click(object sender, RoutedEventArgs e)
        {
            if (api != null)
            {
                EditPanel.HeadEditForm headEdit = new EditPanel.HeadEditForm(api.Doc, facturaListCheck.Items);
                headEdit.Show();
            }
        }

        private void WidthUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (api != null)
            {
                double size = api.Doc.Variable.Give("XGabarit", facturaListCheck.SelectedItem.ToString());
                if ((double)WidthUpDn.Value / size < 1) XUpDn.Value = ((1 - (double)WidthUpDn.Value / size) * 100);
            }
        }

        private void HeightUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (api != null)
            {
                double size = api.Doc.Variable.Give("YGabarit", facturaListCheck.SelectedItem.ToString());
                if ((double)HeightUpDn.Value / size < 1) YUpDn.Value = ((1 - (double)HeightUpDn.Value / size) * 100);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://money.yandex.ru/to/410011060113741");
        }

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            KompasCeilingMini.Properties.Settings.Default.sufix = sufixBox.Text;
            api.Doc.Variable.UpdateNote("Suffix", sufixBox.Text, string.Empty);
            KompasCeilingMini.Properties.Settings.Default.Save();
        }

        private void NameSetting_Click(object sender, RoutedEventArgs e)
        {
            if (api.Doc != null)
            {
                EditPanel.NameEditForm nameEdit = new EditPanel.NameEditForm(api.Doc, facturaListCheck.Items);
                nameEdit.Show();
            }
        }

        private void menuConnect_Click(object sender, RoutedEventArgs e)
        {
            targetApi();
        }

        private void paramBox_1_LostFocus(object sender, RoutedEventArgs e)
        {
            if (api.Doc != null)
                api.Doc.Variable.UpdateNote("Comment1", paramBox_1.Text, string.Empty);
        }

        private void paramBox_2_LostFocus(object sender, RoutedEventArgs e)
        {
            if (api.Doc != null)
                api.Doc.Variable.UpdateNote("Comment2", paramBox_2.Text, string.Empty);
        }

        private void NumberUpDn_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (api != null)
                api.Doc.Variable.Update("Number", (double)NumberUpDn.Value, string.Empty);
        }

        private void MakeOrd_Click(object sender, RoutedEventArgs e)
        {
            if (api != null)
                api.Doc.ST.Coordinate(facturaListCheck.SelectedItem.ToString(), KompasCeilingMini.Properties.Settings.Default.EasyCeilingConnectionString, KompasCeilingMini.Properties.Settings.Default.CoordDopusk);
        }

        private void mashtabChek_Click(object sender, RoutedEventArgs e)
        {
            if (api != null)
            {
                mashtabChek.IsChecked = api.Doc.Mashtab(!(bool)mashtabChek.IsChecked);
                //Пересчитываем контуры
                for (int i = 0; i < facturaListCheck.Items.Count; i++)
                    GetContour(facturaListCheck.Items[i].ToString(), (bool)mashtabChek.IsChecked, false);
            }
            else
                MessageBox.Show(this, "Объект не захвачен", "Сообщение");
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            observer = new BluetoothObserver();
            observer.PropertyChanged += ConnectChange;

            observer.Start("Laser Distance Meter", "0000ffb0-0000-1000-8000-00805f9b34fb", "0000ffb2-0000-1000-8000-00805f9b34fb");

            void ConnectChange(object sender, EventArgs e)
            {
                Dispatcher.Invoke((Action)delegate
                {
                    bool _connectstat;

                    observer.UpdateConnectStat().ContinueWith(r =>
                    {
                        _connectstat = (r.Result);
                        ConnectBtBtn.Content = _connectstat ? "Соеденили" : "Соедение";
                    },
                    TaskScheduler.FromCurrentSynchronizationContext());
                });

            }
        }


        void SetSizeToSelectDimention(double size, int DimRef)
        {
            if (api.Doc != null)
            {

                //если один объект
                object obj = (object)KmpsAppl.KompasAPI.TransferReference(DimRef, api.Doc.D5.reference);
                IDrawingObject pObj = (IDrawingObject)obj;
                SetSizeToLineDim(pObj);


                void SetSizeToLineDim(object dimension)
                {
                    IDrawingObject1 Dim1 = (IDrawingObject1)dimension;
                    Array arrayConstrait = (Array)Dim1.Constraints;


                    foreach (IParametriticConstraint constraint in arrayConstrait)
                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                            if (size != 0)
                                api.Doc.Variable.Update(constraint.Variable, size * 100, string.Empty); //вместо 100 компенсатор
                }
            }
        }

        private void MeasurelistBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {

        }

        private void CreateRoomBtn_Click(object sender, RoutedEventArgs e)
        {
            if (api.Doc != null)
            {
                api.someFlag = false;
                api.Doc.ST.CreateRoom((double)CreateRoomUpDn.Value, 200);
                api.someFlag = true;
            }
        }


        private void StartMeasureBtn_Click(object sender, RoutedEventArgs e)
        {
            DimentionlistBox.DataContext = null;
            DimentionlistBox.DisplayMemberPath = "Name";
            DimentionlistBox.SelectedValuePath = "Reference";
            DimentionlistBox.ItemsSource = api.Doc.ST.GetLineDimentionList();
            //DimentionlistBox.SelectedIndex = 0;

            observer = new BluetoothObserver();
            observer.PropertyChanged += ConnectChange;

            observer.Start("Laser Distance Meter", "0000ffb0-0000-1000-8000-00805f9b34fb", "0000ffb2-0000-1000-8000-00805f9b34fb");

            void ConnectChange(object sender, EventArgs e)
            {
                Dispatcher.Invoke((Action)delegate
                {
                    object Dim = (object)KmpsAppl.KompasAPI.TransferReference((int)DimentionlistBox.SelectedValue, api.Doc.D5.reference);
                    IDrawingObject1 Dim1 = (IDrawingObject1)Dim;
                    Array arrayConstrait = (Array)Dim1.Constraints;


                    foreach (IParametriticConstraint constraint in arrayConstrait)
                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                        {
                                observer.AsyncLastDimmenetion().ContinueWith(r =>
                                {
                                    api.Doc.Variable.Update(constraint.Variable, double.Parse(r.Result) * 100, string.Empty); //вместо 100 компенсатор
                                            DimentionlistBox.SelectedIndex = DimentionlistBox.SelectedIndex + 1 > DimentionlistBox.Items.Count - 1 ? 0 : DimentionlistBox.SelectedIndex + 1;
                                },
                                TaskScheduler.FromCurrentSynchronizationContext());
                        }
                });

            }
        }

        private void DimentionlistBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DimentionlistBox.SelectedIndex > -1)
            {
                api.Doc.GetChooseContainer().UnchooseAll();
                api.Doc.D5.ksLightObj((int)DimentionlistBox.SelectedValue, 1);

                object Dim = (object)KmpsAppl.KompasAPI.TransferReference((int)DimentionlistBox.SelectedValue, api.Doc.D5.reference);
                if (Dim != null)
                {
                    IDrawingObject1 Dim1 = (IDrawingObject1)Dim;
                    Array arrayConstrait = (Array)Dim1.Constraints;

                    if (arrayConstrait != null)
                        foreach (IParametriticConstraint constraint in arrayConstrait)
                            if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                                DimVariableBox.Text = api.Doc.Variable.Give(constraint.Variable, string.Empty).ToString(); //вместо 100 компенсатор
                    DimVariableBox.SelectAll();
                    api.Doc.GetSelectContainer().UnselectAll();
                }
                else MessageBox.Show("Не найден размер");
            }
        }

        private void SplitLineBtn_Click(object sender, RoutedEventArgs e)
        {
            if (api.Doc.ST != null)
            {
                api.someFlag = false;
                api.Doc.ST.SplitLine((double)SplitLineUpDn.Value, (bool)DellBaseDimChek.IsChecked);
                api.someFlag = true;
            }
        }



        private void DimVariableBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                if (DimentionlistBox.SelectedValue != null)
                {
                    object Dim = (object)KmpsAppl.KompasAPI.TransferReference((int)DimentionlistBox.SelectedValue, api.Doc.D5.reference);
                    IDrawingObject1 Dim1 = (IDrawingObject1)Dim;
                    Array arrayConstrait = (Array)Dim1.Constraints;

                    foreach (IParametriticConstraint constraint in arrayConstrait)
                        if (constraint.ConstraintType == ksConstraintTypeEnum.ksCDimWithVariable)
                        {
                            api.Doc.Variable.Update(constraint.Variable, double.Parse(DimVariableBox.Text.Replace('.', ',')), string.Empty);
                            DimentionlistBox.SelectedIndex = DimentionlistBox.SelectedIndex + 1 > DimentionlistBox.Items.Count - 1 ? 0 : DimentionlistBox.SelectedIndex + 1;
                        }
                }
            }
        }

        private void DimVariableBox_GotFocus(object sender, RoutedEventArgs e)
        {
            DimVariableBox.SelectAll();
        }

        private void FixDimBtn_Click(object sender, RoutedEventArgs e)
        {
            api.Doc.ST.SetVariableToDim(false, ksDimensionTextBracketsEnum.ksDimSquareBrackets);
        }

        private void UnfixDimBtn_Click(object sender, RoutedEventArgs e)
        {
            api.Doc.ST.SetVariableToDim(true, ksDimensionTextBracketsEnum.ksDimBrackets);
        }

        private void VariableDimBtn_Click(object sender, RoutedEventArgs e)
        {
            api.Doc.ST.SetVariableToDim(false, ksDimensionTextBracketsEnum.ksDimBracketsOff);
        }

        private void ToSelectDimBtn_Click(object sender, RoutedEventArgs e)
        {
            api.Doc.ST.SetVariableToDim(false, ksDimensionTextBracketsEnum.ksDimBracketsOff, double.Parse(DimVariableBox.Text.Replace('.',',')));
        }

        private void KsContrForm_Deactivated(object sender, EventArgs e)
        {
            if (this.Focusable == true)
            {
                this.Opacity = 0.4;
            }
        }

        private void KsContrForm_Activated(object sender, EventArgs e)
        {
            this.Opacity = 1;
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void HeadGrid_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            clicado = true;
            this.lm = e.GetPosition(this);
            this.lm.Y = Convert.ToInt16(this.Top) + this.lm.Y;
            this.lm.X = Convert.ToInt16(this.Left) + this.lm.X;
        }

        private void HeadGrid_MouseUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            clicado = false;
        }

        private void HeadGrid_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (clicado)
            {
                Point MousePosition = e.GetPosition(this);
                Point MousePositionAbs = new Point();
                MousePositionAbs.X = Convert.ToInt16(this.Left) + MousePosition.X;
                MousePositionAbs.Y = Convert.ToInt16(this.Top) + MousePosition.Y;
                this.Left = this.Left + (MousePositionAbs.X - this.lm.X);
                this.Top = this.Top + (MousePositionAbs.Y - this.lm.Y);
                this.lm = MousePositionAbs;
            }
        }

        private void HeadGrid_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            clicado = false;
        }

        private void ToPointBtn_Click(object sender, RoutedEventArgs e)
        {
            api.Doc.ST.LineToPoint();
        }
    }


}
