using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using Inventor;


namespace InventorCOM
{
    class Tester
    {
        // перед использованием программы необходимо создать в инвенторе документ и открыть лист, на котором ты собираешься рисовать графики 
        static void Main(string[] args)
        {   
            // обертка над API инвентора
            InventorManager test = new InventorManager();
            // получаем открытый документ
            DrawingDocument doc = (DrawingDocument)test.App.ActiveDocument;
            // получаем открытую страницу
            Sheet sheet = doc.ActiveSheet;
            
            // этот график откомментирую ниже, так как он сложнее. в нем я использовал практически все возможности программы
            PlotTemperatureProfile(
                    @"C:\Users\Artem\Desktop\cooling_t_complex.csv", "Температуры", 32, sheet
            );
            
            PlotEfficiencyProfile(
                    @"C:\Users\Artem\Desktop\cooling_t_efficiency.csv", "Эффективность", 8, sheet
            );
        }

        static void PlotTemperatureProfile(String dataPath, String plotName, float yPos, Sheet sheet) {
            // получить графопостроитель. вызов sheet.Sketches.Add() создает на листе новый эскиз. хранить его в отдельной переменной
            // особого смысла нет, так напрямую с ним работа не ведется.
            InventorPlotter plotter = new InventorPlotter(sheet.Sketches.Add());
            // загрузить данные: передается строка с абсолютным путем до данных в формате csv с запятой в качестве разделителя между числам
            // шапки в файле с данными быть не должно
            plotter.ImportData(dataPath);
            // задается минимальное и максимальное значение y , на которое будет распространяться график
            // есть аналогичная функция SetXLim
            plotter.SetYLim(500, 1450);
            // задается положение левого нижнего угла графика относительно левого нижнего угла листа в сантиметрах
            plotter.LocatePlot(5, yPos);
            // задается размер прямоугольника, в который будет вписан график. сначала ширина, затем высота
            plotter.SetPlotSize(35, 20);

            // функция AddColor добавляет цвет в список цветов. цвета будут использованы в порядке добавления при построении графиков
            // если цвета закончились, графики будут построены черным цветом
            plotter.AddColor(255, 0, 0);
            plotter.AddColor(0, 255, 0);
            plotter.AddColor(0, 0, 255);

            // функция строит график с заданной толщиной линии. строятся сразу все графики
            plotter.PlotPieceWise(lineWeight: 0.05f);
            // строятся осевые линии
            plotter.PlotAxisLines();
            // задается имя эскиза, на котором строится график. если эскиз с таким именем уже существует, программа упадет
            plotter.SetPlotName(plotName);
            // построить линии сетки, перпендикулярные оси x. первым аргументом передается начальное значение, на котором будет построена сетка, вторым - 
            // шаг сетки (не помню, почему почему у меня в качестве начального значения стоит 6)
            plotter.PlotXGrid(6, 10f);
            // сделать подписи к графику по оси x. первые 2 аргумента те же, что и в предыдущей функции, затем идет размер шрифта и смещение в направлении от оси
            // список остальных аргументов можешь посмотреть в определении функции
            plotter.PlotXTicks(6, 10f, fontSize: 0.6f, transverseOffset: 0.3f);
            // то же самое, что и для x
            plotter.PlotYGrid(0, 100f);
            // то же самое, что и для x. longwiseCorrection - смещение подписи вдоль оси ввверх
            plotter.PlotYTicks(0, 100f, fontSize: 0.6f, transverseOffset: 1f, longwiseCorrection:0.3f);
        }

        static void PlotEfficiencyProfile(String dataPath, String plotName, float yPos, Sheet sheet)
        {
            InventorPlotter plotter = new InventorPlotter(sheet.Sketches.Add());
            plotter.ImportData(dataPath);
            plotter.SetYLim(0.0f, 1);
            plotter.LocatePlot(5, yPos);
            plotter.SetPlotSize(35, 20);

            plotter.AddColor(0, 0, 255);

            plotter.PlotPieceWise(lineWeight: 0.05f);
            plotter.PlotAxisLines();
            plotter.SetPlotName(plotName);
            plotter.PlotXGrid(6, 10f);
            plotter.PlotXTicks(6, 10f, fontSize: 0.6f, transverseOffset: 0.3f);
            plotter.PlotYGrid(0, 0.1f);
            plotter.PlotYTicks(0, 0.1f, fontSize: 0.6f, transverseOffset: 1f, longwiseCorrection: 0.3f, format:":0.0");
        }
    }
}
