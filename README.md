# PixActivities_ExcelDiagram
Кастомная активность. Активность предназначена для создания столбчатой диаграммы по указанному диапазону таблицы Excel. 

Код для вызова через C#
```
int numberSheet; 
string pathFile;
string firstCell;
string secondCell;
string secondColStart;
string secondColEnd;


Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
appExcel.Visible = true;

//Добавить рабочую книгу
Microsoft.Office.Interop.Excel.Workbook workBook = appExcel.Workbooks.Open(pathFile);

//Получить первый лист документа (счет начинается с 1)
Microsoft.Office.Interop.Excel.Worksheet worksheet
    = (Microsoft.Office.Interop.Excel.Worksheet)appExcel.Worksheets[numberSheet];

//Создание диаграммы
Microsoft.Office.Interop.Excel.ChartObjects xlCharts = (Microsoft.Office.Interop.Excel.ChartObjects)worksheet.ChartObjects(Type.Missing);
Microsoft.Office.Interop.Excel.ChartObject myChart = (Microsoft.Office.Interop.Excel.ChartObject)xlCharts.Add(250, 0, 450, 250);
Microsoft.Office.Interop.Excel.Chart chart = myChart.Chart;
Microsoft.Office.Interop.Excel.SeriesCollection seriesCollection
    = (Microsoft.Office.Interop.Excel.SeriesCollection)chart.SeriesCollection(Type.Missing);
Microsoft.Office.Interop.Excel.Series series = seriesCollection.NewSeries();
series.Values = worksheet.get_Range(firstCell, secondCell);
chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered;

//Закрыть книгу с сохранением
workBook.Save();
workBook.Close();

// Закрыть приложение
appExcel.Quit();
```
