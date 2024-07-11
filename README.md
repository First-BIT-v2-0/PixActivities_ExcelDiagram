# PixActivities_ExcelDiagram
Кастомная активность. Активность предназначена для создания столбчатой диаграммы по указанному диапазону таблицы Excel. 

Код для вызова через C#

string numberSheet;
string pathFile;

// Объявление приложения
Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
appExcel.Visible = true;

//Добавить рабочую книгу
Microsoft.Office.Interop.Excel.Workbook workBook = appExcel.Workbooks.Open(pathFile);

//Получить первый лист документа (счет начинается с 1)
Microsoft.Office.Interop.Excel.Worksheet worksheet
    = (Microsoft.Office.Interop.Excel.Worksheet)appExcel.Worksheets[numberSheet];

//Текущее количество столбцов и строк таблицы
int countCol = worksheet.UsedRange.Columns.Count;
int countRow = worksheet.UsedRange.Rows.Count;

//Создание диаграммы
ChartObjects xlCharts = (ChartObjects)worksheet.ChartObjects(Type.Missing);
ChartObject myChart = (ChartObject)xlCharts.Add(150, 0, 550, 250);
Chart chart = myChart.Chart;
Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection
    = (Microsoft.Office.Interop.Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);
Series series = seriesCollection.NewSeries();
series.XValues = worksheet.get_Range(firstColStart, firstColEnd);
series.Values = worksheet.get_Range(secondColStart, secondColEnd);
chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie;

//Закрыть книгу с сохранением
workBook.Save();
workBook.Close();

// Закрыть приложение
appExcel.Quit();