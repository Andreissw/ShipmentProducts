using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ShipmentProducts
{
    class GS : BaseClass
    {

        string PathCSV = @"C:\CSVFile\Credits.csv";
        string PathExcel = @"C:\CSVFile\Credits.xlsx";
        int count;
        string QuerySN;
        string QueryAggregat;
        string UpdateSerial;

        public override List<string> GetListClients() //Выцгружаем список заказчиков
        {
            using (var fas = new FASEntities())
            {
                return new List<string>() { "ВЛВ", "" };
            }
        }

        public override void RemoveShipped(int id)
        {
            LoadGrid.SelectString($@"use fas update  FAS_SerialNumbers
                                    set ShipedStatus = null
                                    where ShipedStatus = {id}");

            using (var fas = new FASEntities())
            {
               var sh= fas.FAS_Shipped_Table.Where(c => c.ID == id).FirstOrDefault();
               fas.FAS_Shipped_Table.Remove(sh);
               fas.SaveChanges();
            }
        }

        public override List<string> GetModelList() //Выгружаем модели
        {
            using (var fas = new FASEntities())
            {
                var list = fas.FAS_Models.Where(c => fas.FAS_GS_LOTs.Where(v => v.IsActive == true).Select(b => b.ModelID).ToList().Contains(c.ModelID)).Select(c => c.ModelName).ToList();
                list.Add("");
                return list;
            }
        }

        public override void GetShippedTable(DataGridView Grid)
        {
            using (var fas = new FASEntities())
            {
                Grid.DataSource = from s in fas.FAS_Shipped_Table
                                  join a in fas.FAS_Shipped_Status_TB on (int)s.Status equals a.ID
                                  join l in fas.FAS_GS_LOTs on s.LOTID equals l.LOTID
                                  where s.Status == StatusShippID
                                  select new { Лот = l.FULL_LOT_Code, Кол_во = s.count, Дата = s.Date, Статус = a.StatusName };
            }
        }

        string GetModelName()
        {
            using (var fas = new FASEntities())
            {
                return fas.FAS_Models.Where(c => c.ModelID == ModelID).Select(c => c.ModelName).FirstOrDefault();
            }
        }

        string GetFullLotName()
        {
            using (var fas = new FASEntities())
            {
                return fas.FAS_GS_LOTs.Where(c => c.LOTID == LotID).Select(c => c.FULL_LOT_Code).FirstOrDefault();
            }
        }

        string GetSpec()
        {
            using (var fas = new FASEntities())
            {
                return fas.FAS_GS_LOTs.Where(c => c.LOTID == LotID).Select(c => c.Specification).FirstOrDefault();
            }
        }

        public override void GetCustomerID(string name)
        {
            
        }

        public override void GetTable(int countValue)
        {
            this.count = countValue;

            if (!Directory.Exists(@"C:\CSVFile"))
                Directory.CreateDirectory(@"C:\CSVFile");

            Parralel();
           
        }

        async void Parralel()
        {               
            GetQuery();
            await Task.Run(() =>
            { 
                if (!GetCSV()) return;
                GetGrid(); 
                if (!ExcelMethod()) return;

                ConfimForm confim = new ConfimForm();
                var Result = confim.ShowDialog();

                if (Result == DialogResult.OK)
                {
                    var ID = AddShipped();
                    LoadGrid.SelectString($@" use fas update FAS_SerialNumbers
                                 set ShipedStatus = {ID}
                                 where SerialNumber in ( select TOP({count}) t.SerialNumber from( 
                                 select s.SerialNumber ,ROW_NUMBER() over(order by format(ManufDate , 'dd.MM.yyyy'), pac.BoxNum) num
                                 from FAS_SerialNumbers s
                                 left join FAS_Start st on s.SerialNumber = st.SerialNumber 
                                 left join FAS_PackingGS pac on s.SerialNumber =pac.SerialNumber
                                 where s.LOTID in (select g.LOTID from FAS_GS_LOTs g where g.LOTID = {LotID}) and s.ShipedStatus is null and s.IsPacked = 1 and pac.FullBox = 1) as t order by num 
                                 )");                    
                }
            });          
        }

        int AddShipped()
        {
            using (var fas = new FASEntities())
            {
                var ship = new FAS_Shipped_Table()
                {
                    LOTID = LotID,
                    count = count,
                    Date = DateTime.UtcNow.AddHours(2),
                    Status = 1,                    
                };
                fas.FAS_Shipped_Table.Add(ship);
                fas.SaveChanges();

                return fas.FAS_Shipped_Table.OrderByDescending(c => c.ID).Select(c => c.ID).FirstOrDefault();
            }
        }

        void GetGrid()
        {
            LoadGrid.Loadgrid((DataGridView)control.Controls.Find("Grid", true).FirstOrDefault(), QueryAggregat);
            control.Invoke((System.Action)(() => { control.Controls.Find("LoadLB", true).FirstOrDefault().Text = "Агрегирование данных"; }));
        }

        bool GetCSV()
        {
            var file = new StreamWriter(PathCSV, false, Encoding.GetEncoding(1251));
            //string con = @"Data Source=WSG150170\SQLEXPRESS; Initial Catalog= FAS; integrated security=True;";
            string con = @"Data Source=traceability\flat; Initial Catalog= FAS; user id=volodin;password=volodin;";
           
            using (var connection = new SqlConnection(con))
            {
                SqlCommand command = new SqlCommand();
                command = connection.CreateCommand();
                command.CommandType = CommandType.Text;
                command.CommandText = QuerySN;
                connection.Open();
                SqlDataReader sqlRead = command.ExecuteReader();
                
                if (!sqlRead.HasRows)                 
                {
                    MessageBox.Show("По запросу данных не обнаруженно!","Ошибка",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    file.Close();
                    return false;
                }

                file.WriteLine("ModelName;FullSTBSN;SerialNumber;CASID;SmartCardID;MAC;ManufDate;Liter;PalletNum;BoxNum;UnitNum;FullLotCode;Specification");
                count =  0;
                while (sqlRead.Read())
                {
                    count += 1;
                    file.WriteLine($@"{sqlRead.GetValue(0)}; {sqlRead.GetString(1)};{sqlRead.GetValue(2)};{sqlRead.GetValue(3)};{sqlRead.GetValue(4)};{sqlRead.GetValue(5)};{sqlRead.GetValue(6)};{sqlRead.GetValue(7)};{sqlRead.GetInt16(8)};{sqlRead.GetInt16(9)};{sqlRead.GetInt16(10)};{sqlRead.GetString(11)};{sqlRead.GetString(12)}");
                    control.Invoke((System.Action)(() => { control.Controls.Find("LoadLB",true).FirstOrDefault().Text = "Загруженно записей - " + count; }));                    
                }
            }
            file.Close();
            return true;

        }

        bool ExcelMethod()
        {
            control.Invoke((System.Action)(() => { control.Controls.Find("LoadLB", true).FirstOrDefault().Text = "Формирование Excel Файла"; }));           
            var grid = (DataGridView)control.Controls.Find("Grid", true).FirstOrDefault();
            Excel.Application x1 = new Excel.Application() { DisplayAlerts = false }; //Инициализации приложения Excel
            Excel.Workbook wb = x1.Workbooks.Open(PathCSV); //Создание новой книги через открытие CSV файла, где лежать данные с базы 
            try
            {
                #region Создание Excel страницы из CSV файла 1 ЛИСТ
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1); //Копируем данные с CSV в Excel                
                ws.Name = "Sn Для Алтухова"; //Имя листа
                Excel.Range used = ws.UsedRange; //Хватаем диапозон данных в формате CSV
                used.TextToColumns( //Форматируем столбец, разделяем его на несколько столбоцом по разделителю ";"
                    DataType: Excel.XlTextParsingType.xlDelimited,
                    TextQualifier: Excel.XlTextQualifier.xlTextQualifierDoubleQuote,
                    ConsecutiveDelimiter: false,
                    Tab: false,
                    Semicolon: true,
                    Comma: false,
                    Space: false,
                    Other: false,
                    FieldInfo: new object[,] { {1, Excel.XlColumnDataType.xlGeneralFormat}, {2, Excel.XlColumnDataType.xlTextFormat}//Формат столбцов
                    ,{3, Excel.XlColumnDataType.xlGeneralFormat},{4, Excel.XlColumnDataType.xlTextFormat},{5, Excel.XlColumnDataType.xlTextFormat}
                    });
                used.EntireColumn.AutoFit(); //Автоподор размера строки
                var column = used.get_Range("C1", Type.Missing); //Выделаем столбец C
                
                column.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);//Вставляю столбец после столбца FuLLSTBSN 
                used.Range[$"C2:C{count + 1}"].Value = @"=RIGHT(B2,23)"; //Устанавливаю формулу, обрезаю слева один символ
                used.Range[$"C2:C{count + 1}"].Copy(); //Копирую
                used.Range[$"C2:C{count + 1}"].PasteSpecial(Excel.XlPasteType.xlPasteValues, Operation: Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks: false, Transpose: false); //Вставляю как текст, чтобы формула исчезла
                used.Range[$"C2:C{count + 1}"].TextToColumns(FieldInfo: new object[,] { { 1, Excel.XlColumnDataType.xlTextFormat } }); //Устанавливаю текстовый формат на столбец
                used.get_Range("B1", Type.Missing).EntireColumn.Delete();//Удаляю столбец FULLSTBSN с пробелом
                used.Range["B1"].Value = "FullSTBSN"; //Называю по новой столбец
                #endregion

                #region Создание отчёта для заказчика ЛИСТ 2
                control.Invoke((System.Action)(() => { control.Controls.Find("LoadLB", true).FirstOrDefault().Text = "Формирование отчёта для заказчиков"; }));
                ws.Copy(After: wb.Worksheets[wb.Worksheets.Count]); //Дублируем новый лист
                var WS2 = (Excel.Worksheet)wb.Worksheets.get_Item(2); //Инициализируем новую переменную от нового листа
                WS2.Name = "Для заказчика"; //Называем лист своим именем 
                WS2.get_Range("D:F").EntireColumn.Delete();//Выделяем столбцы и удаляем
                WS2.get_Range("E:H").EntireColumn.Delete();//Выделяем столбцы и удаляем

                var col = WS2.get_Range("A1", Type.Missing); //Выделаем столбец A
                col.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);  //Вставляем новый столбец перед столбцом A         
                #region Имена столбцов
                WS2.Range["A1"].Value = "№ п/п";
                WS2.Range["B1"].Value = "Модель";
                WS2.Range["C1"].Value = "Серийный номер (полностью)";
                WS2.Range["E1"].Value = "Дата выпуска";
                WS2.Range["D1"].Value = "";
                WS2.Range["G1"].Value = "Спецификация";
                WS2.Range["F1"].Value = "Номер заказа";
                WS2.Range["H1"].Value = "Кол-во, шт";
                #endregion

                WS2.Range["D:D"].Group();
                WS2.Range["A1:H1"].Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;          
                WS2.Range["A1:H1"].Interior.Color = 49407;
                WS2.Range["D1"].Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;

                #region Форматрование листа 2 Для заказчика
                var rowcolumn = WS2.Range["1:1"];
                rowcolumn.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rowcolumn.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                rowcolumn.WrapText = true;
                rowcolumn.Orientation = 0;
                rowcolumn.AddIndent = false;
                rowcolumn.IndentLevel = 0;
                rowcolumn.ShrinkToFit = false;
                rowcolumn.MergeCells = false;
                rowcolumn.RowHeight = 30.75;
                rowcolumn.EntireColumn.AutoFit();
                WS2.Range["A1:H1"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                #endregion

                #region Нумерация строк
                int row = count + 1;
                var arr = new int[row, 1];
                var arr2 = new int[row, 1];            

                Parallel.For(1, row, i =>
                 {
                     arr[i - 1, 0] = i;
                     arr2[i - 1, 0] = 1;
                 });


                Excel.Range range = WS2.Range[$"A2:A{count+1}"];
                range.Value = arr;
                Excel.Range range2 = WS2.Range[$"H2:H{count+1}"];
                range2.Value = arr2;
                #endregion

                #endregion

                #region Форматирование агрегатных данных Лист 3
                Excel.Worksheet sheet = wb.Worksheets.Add(); //Добавляем новый лист
                sheet.Name = "Агрегатные данные для заказчика";
                sheet.Range["A1:I100"].Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                x1.ActiveWindow.View = Excel.XlWindowView.xlPageBreakPreview; //Тип просмотра листа
                sheet.Cells[8, 3].Value = "№ п/п"; sheet.Cells[8, 5].Value = "Диапозон серийных номеров"; sheet.Cells[8, 7].Value = "Количество";
                sheet.Cells[9, 3].Value = "с"; sheet.Cells[9, 4].Value = "по";
                sheet.Cells[9, 5].Value = "с"; sheet.Cells[9, 6].Value = "по";
                sheet.Cells[9, 7].Value = "шт";

                sheet.get_Range("C8", "D8").Merge();               
                sheet.get_Range("E8", "G8").ColumnWidth = 15;
                sheet.get_Range("C8", "D8").ColumnWidth = 8;
                sheet.get_Range("E8", "F8").Merge();
                sheet.get_Range("E8", "G8").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.get_Range("E8", "G8").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                for (int i = 0; i < grid.RowCount; i++)              
                    for (int k = 0; k < grid.ColumnCount; k++)                    
                        sheet.Cells[10 + i, 3 + k].Value = grid[k, i].Value;                       

                sheet.Range[$"C{8}:G{grid.RowCount + 8}"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Range[$"F{grid.RowCount + 9}"].Value = "Итого";
                sheet.Range[$"A{grid.RowCount + 9 + 4}"].Value = @"\_______________________|___Алтухов А.С.___/";
                sheet.Range[$"A{grid.RowCount + 9 + 5}"].Value = @"Начальник центрального склада";
                sheet.Range[$"A1"].Value = $"Отчёт по отгрузке Цифровой приемник {GetModelName()}";
                sheet.Range[$"A1"].Font.Size = 16;
                sheet.Range[$"A1:G2"].MergeCells = true;
                sheet.Range[$"A1:G2"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.Range[$"A1:G2"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                sheet.Range[$"A5"].Value = @"Заказ: ";
                sheet.Range[$"B5"].Value = $"{GetFullLotName()}";
                sheet.Range[$"A6"].Value = @"Спец.";
                sheet.Range[$"B6"].Value = $"{GetSpec()}";
                sheet.Range[$"A7"].Value = @"Номер а/м:";
                sheet.Range[$"G{grid.RowCount + 9}"].Value = $"=SUM(G8:G{grid.RowCount + 8})"; //Подсчёт суммы
                sheet.Range[$"C8:G{grid.RowCount + 8}"].Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                sheet.Range[$"C8:G9"].Interior.TintAndShade = -0.249977111117893; //Цвет ячеек                
                sheet.Range[$"A:G"].Select();
                x1.ActiveSheet.PageSetup.PrintArea = "A:G";
                sheet.Range[$"I10"].Value = "К отгрузке";
                sheet.Range[$"I11"].Value = "Литер";
                sheet.Range[$"I12"].Value = "Паллет";
                sheet.Range[$"J12"].Value = "с";
                sheet.Range[$"L12"].Value = "по";
                sheet.Range[$"L13"].Value = "по";
                sheet.Range[$"J13"].Value = "с";
                sheet.Range[$"I13"].Value = "Коробка";
                sheet.Range[$"I14:L14"].Merge();
                sheet.Range[$"I14"].Value = "Количество";
                sheet.Range[$"I15"].Value = "Итого";
                sheet.Range[$"I10:M15"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Range[$"I15:L15"].Merge();
                sheet.Range[$"J11:M11"].Merge();
                sheet.Range[$"I10:M10"].Merge();
                sheet.Range[$"I10:M10"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.Range[$"I1"].ColumnWidth = 8;
                sheet.Range[$"J1"].ColumnWidth = 3.28;
                sheet.Range[$"K1"].ColumnWidth = 7;
                sheet.Range[$"L1"].ColumnWidth = 3.28;
                sheet.Range[$"M1"].ColumnWidth = 9.86;

                //sheet.Range[$"C:I"].AutoFit();
                #endregion

                #region Создание 4 листа TricolorReport
                ws.Copy(After: wb.Worksheets[wb.Worksheets.Count]); //Дублируем новый лист
                var WS3 = (Excel.Worksheet)wb.Worksheets.get_Item(4); //Инициализируем новую переменную от нового листа
                WS3.Name = "Report Tricolor";
                WS3.Columns["A:A"].Insert();
                WS3.Columns["A:A"].Insert();
                WS3.Range["A1"].Value = "Manufacturer";
                WS3.Range["B1"].Value = "Operator";

                WS3.Range[$"A2:A{count+1}"].Value = "DTVS";
                WS3.Range[$"B2:B{count+1}"].Value = "TRICOLOR";
                WS3.Columns["E:E"].Delete();
                WS3.Columns["G:G"].Insert();
                WS3.Range["G1"].Value = "Lockdown status (OK / NOK)";
                WS3.Range[$"G2:G{count+1}"].Value = "OK";
                WS3.Columns["J:O"].Delete();
                WS3.Columns["A:O"].EntireColumn.AutoFit();
                WS3.Columns["A:O"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                WS3.Range[$"A1:I{count+1}"].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                WS3.Rows["1:1"].Insert();
                WS3.Rows["2:2"].Insert();
                WS3.Range["A1"].Value = "Version: 1.0";
                WS3.Range["A1"].Select();

                #endregion

                x1.ActiveWindow.Zoom = 110;
                wb.SaveAs(PathExcel, 51); //Сохраняем файл Excel в формате xlsx
                
                wb.Close();
                x1.Quit();
                //x1 = null;
                //wb = null;
                //ws = null;
                //WS2 = null;
                //WS3 = null;
                //GC.Collect();
                Process.Start(PathExcel);

                //for (int i = 0; i < grid.RowCount; i++)
                //    for (int k = 0; k < grid.

                control.Invoke((System.Action)(() => { control.Controls.Find("LoadLB", true).FirstOrDefault().Text = "Готово!"; }));
                return true;
            }
            catch (Exception e)
            {
                wb.Close();
                x1.Quit();
                MessageBox.Show("Ошибка при формировании Excel Файла \n" + e.ToString());
                control.Invoke((System.Action)(() => { control.Controls.Find("LoadLB", true).FirstOrDefault().Text = "Ошибка при формировании Excel Файла "; }));
                return false;
            }
        }


        public override List<string> GetLots()
        {
            using (var fas = new FASEntities())
            {
                var list = fas.FAS_GS_LOTs.Where(c => c.ModelID == ModelID & c.IsActive == true).Select(c => c.LOTCode + " | " + c.FULL_LOT_Code).ToList();
                   list.Add("");
                   return list;
            }
        }

        public override void GetLotID(string NameLot) //Смотрим LOTID
        {
            using (var fas = new FASEntities())
            {
                LotID =  fas.FAS_GS_LOTs.Where(c => c.LOTCode + " | " + c.FULL_LOT_Code == NameLot ).Select(c => c.LOTID).FirstOrDefault();
            }
        }

        void GetQuery()
        {
            #region Код с базой FAS
            //QuerySN = $@"use fas
            //                   SELECT 
            //                   ModelName,FullSTBSN ,p.[SerialNumber],CASID,SmartCardID	,MAC ,format(ManufDate , 'dd.MM.yyyy') ManufDate
            //                   ,([LiterName] + convert ( nvarchar(2),p.LiterIndex)) Liter ,[PalletNum] ,[BoxNum] ,[UnitNum], FULL_LOT_Code,  Specification
            //                   FROM [FAS].[dbo].[FAS_PackingGS]  p
            //                   left join FAS_Start  st on st.SerialNumber = p.SerialNumber 
            //                   left join FAS_Upload U on U.SerialNumber = p.SerialNumber 
            //                   left join FAS_GS_LOTs Lt On Lt.LOTID =p.LOTID
            //                   left join FAS_Models M On M.ModelID = Lt.ModelID
            //                   left join FAS_Liter As L on L.ID = P.LiterID

            //                   where p.SerialNumber in 
            //                   (
            //                select TOP({count}) t.SerialNumber from( 
            //                select s.SerialNumber ,ROW_NUMBER() over(order by format(ManufDate , 'dd.MM.yyyy'), pac.BoxNum) num
            //                from FAS_SerialNumbers s
            //                left join FAS_Start st on s.SerialNumber = st.SerialNumber 
            //                   left join FAS_PackingGS pac on s.SerialNumber =pac.SerialNumber
            //                where s.LOTID in (select g.LOTID from FAS_GS_LOTs g where g.LOTID = {LotID}) and s.ShipedStatus is null and s.IsPacked = 1 and pac.FullBox = 1) as t order by num 

            //                   )
            //                   order by LiterID, p.LiterIndex , BoxNum,UnitNum";

            //QueryAggregat = $@"select   min(num1) as по, max(num1) as шт, min(ta.SerialNumber) 'с Диапазон серийных номеров', max(ta.SerialNumber) 'по Диапазон серийных номеров', max(num1) - min(num1) + 1 'Шт'
            //                   from(
            //                   SELECT p.[SerialNumber], p.[SerialNumber] - ROW_NUMBER() over (order by p.[SerialNumber]) as num, ROW_NUMBER() over (order by p.[SerialNumber]) as num1

            //                   FROM [FAS].[dbo].[FAS_PackingGS]  p
            //                   left join FAS_Start  st on st.SerialNumber = p.SerialNumber 
            //                   left join FAS_Upload U on U.SerialNumber = p.SerialNumber 
            //                   left join FAS_GS_LOTs Lt On Lt.LOTID =p.LOTID
            //                   left join FAS_Models M On M.ModelID = Lt.ModelID
            //                   left join FAS_Liter As L on L.ID = P.LiterID
            //                   where p.SerialNumber in 
            //                   (
            //                select  TOP({count}) t.SerialNumber from( 
            //                select s.SerialNumber ,ROW_NUMBER() over(order by format(ManufDate , 'dd.MM.yyyy'), pac.BoxNum) num
            //                from FAS_SerialNumbers s
            //                left join FAS_Start st on s.SerialNumber = st.SerialNumber 
            //                   left join FAS_PackingGS pac on s.SerialNumber =pac.SerialNumber
            //                where s.LOTID in (select g.LOTID from FAS_GS_LOTs g where g.LOTID = {LotID}) and s.ShipedStatus is null and s.IsPacked = 1 and pac.FullBox = 1) as t order by num 
            //                )) TA group by num order by по";
            #endregion

            QuerySN = $@"use fas
                               SELECT 
                               ModelName,FullSTBSN ,p.[SerialNumber],CASID,SmartCardID	,MAC ,format(ManufDate , 'dd.MM.yyyy') ManufDate
                               ,([LiterName] + convert ( nvarchar(2),p.LiterIndex)) Liter ,[PalletNum] ,[BoxNum] ,[UnitNum], FULL_LOT_Code,  Specification
                               FROM [FAS].[dbo].[FAS_PackingGS]  p
                               left join FAS_Start  st on st.SerialNumber = p.SerialNumber 
                               left join FAS_Upload U on U.SerialNumber = p.SerialNumber 
                               left join FAS_GS_LOTs Lt On Lt.LOTID =p.LOTID
                               left join FAS_Models M On M.ModelID = Lt.ModelID
                               left join FAS_Liter As L on L.ID = P.LiterID

                               where p.SerialNumber in 
                               (
                            select TOP({count}) t.SerialNumber from( 
                            select s.SerialNumber ,ROW_NUMBER() over(order by format(ManufDate , 'dd.MM.yyyy'), pac.BoxNum) num
                            from FAS_SerialNumbers s
                            left join FAS_Start st on s.SerialNumber = st.SerialNumber 
                               left join FAS_PackingGS pac on s.SerialNumber =pac.SerialNumber
                            where s.LOTID in (select g.LOTID from FAS_GS_LOTs g where g.LOTID = {LotID}) and s.ShipedStatus is null and s.IsPacked = 1 and pac.FullBox = 1) as t order by num 

                               )
                               order by LiterID, p.LiterIndex , BoxNum,UnitNum";

            QueryAggregat = $@"select   min(num1) as по, max(num1) as шт, min(ta.SerialNumber) 'с Диапазон серийных номеров', max(ta.SerialNumber) 'по Диапазон серийных номеров', max(num1) - min(num1) + 1 'Шт'
                               from(
                               SELECT p.[SerialNumber], p.[SerialNumber] - ROW_NUMBER() over (order by p.[SerialNumber]) as num, ROW_NUMBER() over (order by p.[SerialNumber]) as num1

                               FROM [FAS].[dbo].[FAS_PackingGS]  p
                               left join FAS_Start  st on st.SerialNumber = p.SerialNumber 
                               left join FAS_Upload U on U.SerialNumber = p.SerialNumber 
                               left join FAS_GS_LOTs Lt On Lt.LOTID =p.LOTID
                               left join FAS_Models M On M.ModelID = Lt.ModelID
                               left join FAS_Liter As L on L.ID = P.LiterID
                               where p.SerialNumber in 
                               (
                            select  TOP({count}) t.SerialNumber from( 
                            select s.SerialNumber ,ROW_NUMBER() over(order by format(ManufDate , 'dd.MM.yyyy'), pac.BoxNum) num
                            from FAS_SerialNumbers s
                            left join FAS_Start st on s.SerialNumber = st.SerialNumber 
                               left join FAS_PackingGS pac on s.SerialNumber =pac.SerialNumber
                            where s.LOTID in (select g.LOTID from FAS_GS_LOTs g where g.LOTID = {LotID}) and s.ShipedStatus is null and s.IsPacked = 1 and pac.FullBox = 1) as t order by num 
                            )) TA group by num order by по";

        }

    }
}
