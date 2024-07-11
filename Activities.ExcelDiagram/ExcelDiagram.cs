using System;
using Activities.ExcelDiagram.Properties;
using BR.Core;
using BR.Core.Attributes;
using Microsoft.Office.Interop.Excel;

namespace Activities.ExcelDiagram
{
    [LocalizableScreenName("ExcelDiagram_ScreenName", typeof(Resources))] // Имя активности, отображаемое в списке активностей и в заголовке шага
    [LocalizablePath("PathActivities", typeof(Resources))] // Путь к активности в панели "Активности"
    [LocalizableDescription("Activities_Description", typeof(Resources))] // описание активности
    
    [Image(typeof(ExcelDiagram), "Activities.ExcelDiagram.colorPic.png")] //Иконка активности

    public class ExcelDiagram : Activity
    {
        [LocalizableScreenName("PathFile_ScreenName", typeof(Resources))]
        [LocalizableDescription("PathFile_Description", typeof(Resources))]
        [IsRequired]
        [IsFilePathChooser]
        public string pathFile { get; set; }

        private int defValue = 1;
        [LocalizableScreenName("NumberSheet_ScreenName", typeof(Resources))]
        [LocalizableDescription("NumberSheet_Description", typeof(Resources))]
        [IsRequired]

        public int numberSheet
        {
            get { return defValue; }
            set { defValue = value; }
        }

        [LocalizableScreenName("Cell11_ScreenName", typeof(Resources))]
        [LocalizableDescription("Cell11_Description", typeof(Resources))]
        [IsRequired]
        public string firstColStart { get; set; }

        [LocalizableScreenName("Cell12_ScreenName", typeof(Resources))]
        [LocalizableDescription("Cell12_Description", typeof(Resources))]
        [IsRequired]
        public string firstColEnd { get; set; }

        [LocalizableScreenName("Cell21_ScreenName", typeof(Resources))]
        [LocalizableDescription("Cell21_Description", typeof(Resources))]
        [IsRequired]
        public string secondColStart { get; set; }

        [LocalizableScreenName("Cell22_ScreenName", typeof(Resources))]
        [LocalizableDescription("Cell22_Description", typeof(Resources))]
        [IsRequired]
        public string secondColEnd { get; set; }

        public override void Execute(int? optionID)
        {
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
        }

    }
}
