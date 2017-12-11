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
        static void Main(string[] args)
        {
            InventorManager test = new InventorManager();
            DrawingDocument doc = (DrawingDocument)test.App.ActiveDocument;
            Sheet sheet = doc.ActiveSheet;
            DrawingSketch sketch = sheet.Sketches.Add();
            InventorPlotter plotter = new InventorPlotter(sketch);
            plotter.ImportData(@"C:\Users\Artem\Desktop\test.csv");
            plotter.LocatePlot(10, 10);
            plotter.AddColor(255, 0, 0);
            plotter.PlotPieceWise();
            plotter.PlotAxisLines();
            plotter.SetPlotName("Пробный график");
            plotter.PlotXGrid(0, 0.5f);
            plotter.PlotXTicks(0, 0.5f, scale:2);
            plotter.PlotYGrid(0, 5f);
            plotter.PlotYTicks(0, 5f, scale:2);
            plotter.PlotYTicks(0, 5f, scale: 4, transverseOffset: 1);
            plotter.LabelXAxis("\u03C7 label");
            plotter.LabelYAxis("\u03B3 label", position: "top", transverseCorrection: 2);

            /*
            plotter.DrawXArrow(traverseOffset: 3);
            plotter.DrawYArrow(traverseOffset: 2);

            /*
            plotter.ImportData(@"C:\Users\Клюквина Татьяна\Desktop\test plotter\Массивы\еуые.csv");
            plotter.LocatePlot(5, 5);
            plotter.PlotPieceWise();
            plotter.PlotAxisLines(yOffsetTop:1);
            plotter.SetPlotName("Пробный график");
            plotter.PlotXGrid(0, 0.5f);
            plotter.PlotYGrid(0, 0.5f);
            plotter.PlotXTicks(0, 1f);
            plotter.PlotYTicks(0, 1f);
            plotter.LabelXAxis("\u039B");
            plotter.LabelYAxis("\u03A3", position: "top");
            plotter.DrawXArrow(traverseOffset: 3);
            plotter.DrawYArrow(traverseOffset: 2);
            */
        }
    }
}
