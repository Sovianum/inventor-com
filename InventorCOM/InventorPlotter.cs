using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using Inventor;

namespace InventorCOM
{
    class InventorPlotter
    {
        public float[] XArray
        {
            get { return xArray; }
            set { xArray = value; }
        }
        public List<float[]> YArrays
        {
            get { return yArrays; }
            set { yArrays = value; }
        }
        private DrawingSketch sketch;
        private Application app;
        private TransientGeometry transGeom;
        private float[] xArray;
        private List<float[]> yArrays;
        private float initX;
        private float initY;
        private float length;
        private float height;
        private Point2d leftBottom;
        private Point2d leftTop;
        private Point2d rightBottom;
        private Point2d rightTop;
        private TickConverterDelegate xConverter;
        private TickConverterDelegate yConverter;

        private List<Color> colors;

        private float xMin, xMax, yMin, yMax;
        private float yScale;
        private float xScale;

        public InventorPlotter(DrawingSketch sketch) {
            this.sketch = sketch;
            this.app = (Application)sketch.Application;
            this.transGeom = this.app.TransientGeometry;
            this.yArrays = new List<float[]>();
            this.initX = 0;
            this.initY = 0;
            this.length = 9;
            this.height = 6;
            this.xConverter = x => (x).ToString();
            this.yConverter = y => (y).ToString();
            colors = new List<Color>();
        }

        public void ImportData(string path) {
            StreamReader sr = new StreamReader(path);
            List<string[]> stringData = new List<string[]>();

            while (!sr.EndOfStream) {
                stringData.Add(sr.ReadLine().Split(','));
            }

            int colCount = stringData[0].GetLength(0);
            int rowCount = stringData.Count;

            for (int j = 0; j != colCount; ++j) {
                float[] floatVect = new float[rowCount];
                for (int i = 0; i != rowCount; ++i) {
                    floatVect[i] = System.Single.Parse(stringData[i][j], System.Globalization.CultureInfo.InvariantCulture);
                }

                if (j == 0)
                {
                    this.xArray = new float[floatVect.Length];
                    this.xArray = floatVect;
                }
                else {
                    this.yArrays.Add(floatVect);
                }
            }

            UpdateInternals();
        }

        //импорт данных по заданному пути, указанный во вкладке Tester
        public void SetPlotName(string name)
        {
            //устанавливает название эскиза назаданное (указывается в tester)
            this.sketch.Name = name;
        }

        public void PlotPieceWise(float lineWeight = 0.025f) {
            //строит все графики, которые есть в массиве. можно указать толщину линии, сейчас стоит тонкая. f ставить обязательно!!!!
            Transaction plotTransaction = this.app.TransactionManager.StartTransaction(this.app.ActiveDocument, "Plotting");
            for (int i = 0; i != yArrays.Count; i++) {
                this.PlotPieceWise1Profile(this.xArray, yArrays[i], GetColor(i), lineWeight);
            }
            plotTransaction.End();
        }

        public void LocatePlot(float x, float y) {
            //местоположение графика
            this.initX = x;
            this.initY = y;
        }

        public void SetPlotSize(float length, float height) {
            this.length = length;
            this.height = height;
            UpdateInternals();
        }

        public void PlotAxisLines(float xOffsetLeft = 0, float xOffsetRight = 0, float yOffsetBottom = 0, float yOffsetTop = 0, bool isRectangle = true, float lineWeight = 0.01f)
        {
            //построение обоих осей
            //xOffsetLeft - величина отступа от графика влево. по умолчанию стоит 0 (указаны через равно)
            //xOffsetRight -                            вправо
            //yOffsetBottom -                           вниз
            //yOffsetTop -                              вверх
            //isRectangle - сетка углом или прямоугольником. По умолчанию стоит прямоугольник
            //lineWeight - толщина линии осей

            xMin += this.initX - xOffsetLeft;
            xMax += this.initX + xOffsetRight;

            yMin += this.initY - yOffsetBottom;
            yMax += this.initY + yOffsetTop;

            this.leftBottom = transGeom.CreatePoint2d(ToPlotCoordX(xMin), ToPlotCoordY(yMin));
            this.rightBottom = transGeom.CreatePoint2d(ToPlotCoordX(xMax), ToPlotCoordY(yMin));
            this.leftTop = transGeom.CreatePoint2d(ToPlotCoordX(xMin), ToPlotCoordY(yMax));
            this.rightTop = transGeom.CreatePoint2d(ToPlotCoordX(xMax), ToPlotCoordY(yMax));

            List<SketchLine> axisLines = new List<SketchLine>();

            Transaction plotTransaction = this.app.TransactionManager.StartTransaction(this.app.ActiveDocument, "Plotting");
            sketch.Edit();
            axisLines.Add(sketch.SketchLines.AddByTwoPoints(this.leftBottom, this.rightBottom));
            axisLines.Add(sketch.SketchLines.AddByTwoPoints(this.leftBottom, this.leftTop));

            if (isRectangle)
            {
                axisLines.Add(sketch.SketchLines.AddByTwoPoints(this.leftTop, this.rightTop));
                axisLines.Add(sketch.SketchLines.AddByTwoPoints(this.rightBottom, this.rightTop));
            }
            sketch.ExitEdit();

            foreach (SketchLine line in axisLines) {
                line.LineWeight = lineWeight;
            }
            plotTransaction.End();
        }

        public void PlotXGrid(float originValue, float step, float lineWeight = 0.01f) {
            //строит сетку перпендикулярно оси х
            //originValue - начальное значение по оси x от которого начинает строится сетка
            //step - шаг сетки в см чертежа 
            Transaction plotTransaction = this.app.TransactionManager.StartTransaction(this.app.ActiveDocument, "Plotting");
            this.sketch.Edit();

            List<float> gridOffsets = new List<float>();
            List<SketchLine> gridLines = new List<SketchLine>();
            
            for (float value = originValue; value * xScale <= length; value += step) {
                gridOffsets.Add(value * xScale);
            }

            foreach (float offset in gridOffsets) {
                Point2d bottomPoint = transGeom.CreatePoint2d(offset + initX, this.leftBottom.Y);
                Point2d topPoint = transGeom.CreatePoint2d(offset + initX, this.leftTop.Y);
                gridLines.Add(sketch.SketchLines.AddByTwoPoints(bottomPoint, topPoint));
            }

            foreach (SketchLine line in gridLines) {
                line.LineWeight = lineWeight;
            }

            this.sketch.ExitEdit();
            plotTransaction.End();

        }

        public void PlotYGrid(float originValue, float step, float lineWeight = 0.01f)
        {
            Transaction plotTransaction = this.app.TransactionManager.StartTransaction(this.app.ActiveDocument, "Plotting");
            this.sketch.Edit();

            List<float> gridOffsets = new List<float>();
            List<SketchLine> gridLines = new List<SketchLine>();

            for (float value = originValue; value * yScale < height; value += step)
            {
                gridOffsets.Add(value * yScale);
            }

            foreach (float offset in gridOffsets)
            {
                Point2d bottomPoint = transGeom.CreatePoint2d(this.leftBottom.X, offset + initY);
                Point2d topPoint = transGeom.CreatePoint2d(this.rightBottom.X, offset + initY);
                gridLines.Add(sketch.SketchLines.AddByTwoPoints(bottomPoint, topPoint));
            }

            foreach (SketchLine line in gridLines)
            {
                line.LineWeight = lineWeight;
            }

            this.sketch.ExitEdit();
            plotTransaction.End();

        }

        public void PlotXTicks(float originValue, float step, float transverseOffset = 0.1f, float fontSize = 0.3f, float longwiseCorrection = 0f, float scale = 1) {
            //построение подписей к оси х
            //transverseOffset - смещение перпендикулярно оси х вниз
            //fontSize - размер шрифта
            //longwiseCorrection - величина, на которую можно сместить текст вдоль оси х
            List<float> longwiseOffsets = new List<float>();
            List<TextBox> labels = new List<TextBox>();
            TickConverterDelegate localConverter = x => string.Format("<StyleOverride FontSize='{0}'>{1}</StyleOverride>", fontSize, this.xConverter(x / xScale * scale));

            this.sketch.Edit();

            for (float value = originValue; value * xScale < length; value += step)
            {
                longwiseOffsets.Add(value * xScale);
            }

            foreach (float longwiseOffset in longwiseOffsets) {
                Point2d textPos = transGeom.CreatePoint2d(longwiseOffset + initX - longwiseCorrection, initY - transverseOffset);
                TextBox label = this.sketch.TextBoxes.AddFitted(textPos, localConverter(longwiseOffset));
                label.FormattedText = localConverter(longwiseOffset);
                label.Style.VerticalJustification = VerticalTextAlignmentEnum.kAlignTextUpper;
                label.Style.HorizontalJustification = HorizontalTextAlignmentEnum.kAlignTextCenter;
            }
            this.sketch.ExitEdit();
        }

        public void PlotYTicks(float originValue, float step, float transverseOffset = 0.3f, float fontSize = 0.3f, float longwiseCorrection = 0.15f, float scale = 1)
        {
            List<float> longwiseOffsets = new List<float>();
            List<TextBox> labels = new List<TextBox>();
            TickConverterDelegate localConverter = y => string.Format("<StyleOverride FontSize='{0}'>{1}</StyleOverride>", fontSize, this.yConverter(y / yScale * scale));


            this.sketch.Edit();
            for (float value = originValue; value * yScale < height; value += step)
            {
                longwiseOffsets.Add(value * yScale);
            }

            foreach (float longwiseOffset in longwiseOffsets)
            {
                Point2d textPos = transGeom.CreatePoint2d(initX - transverseOffset, longwiseOffset + initY + longwiseCorrection);
                TextBox label = this.sketch.TextBoxes.AddFitted(textPos, localConverter(longwiseOffset));
                label.FormattedText = localConverter(longwiseOffset);
            }
            this.sketch.ExitEdit();
        }

        public void LabelXAxis(string labelText, string position = "center", float longwiseCorrection = 1, float transverseCorrection = 1, float labelAngle = 0, float fontSize = 0.3f) {
            //название оси х
            //position позиция подписи. можно поставить right
            //labelAngle - угол поворота текста (по умолчанию=0)
            this.sketch.Edit();
            Point2d labelPosition = transGeom.CreatePoint2d(0, 0);
            if (position == "center") {
                float labelX = (float)(this.leftBottom.X + this.rightBottom.X) / 2 - longwiseCorrection;
                float labelY = (float)this.leftBottom.Y - transverseCorrection;
                labelPosition = transGeom.CreatePoint2d(labelX, labelY);
            }
            else if (position == "right") {
                float labelX = (float)this.rightBottom.X - longwiseCorrection;
                float labelY = (float)this.leftBottom.Y - transverseCorrection;
                labelPosition = transGeom.CreatePoint2d(labelX, labelY);
            }
            else {
                System.Diagnostics.Debug.Assert(false, "Неверно указано положение осевой подписи");
            }

            TextBox axisLabel = this.sketch.TextBoxes.AddFitted(labelPosition, labelText);

            axisLabel.FormattedText = string.Format("<StyleOverride FontSize='{0}'>{1}</StyleOverride>", fontSize, axisLabel.Text);
            axisLabel.Rotation = labelAngle;

            this.sketch.ExitEdit();
        }

        public void LabelYAxis(string labelText, string position = "center", float longwiseCorrection = 1, float transverseCorrection = 1, float labelAngle = (float)Math.PI / 2, float fontSize = 0.3f)
        {
            this.sketch.Edit();
            Point2d labelPosition = transGeom.CreatePoint2d(0, 0);
            if (position == "center")
            {
                float labelX = (float)this.leftBottom.X - transverseCorrection;
                float labelY = (float)(this.leftBottom.Y + this.leftTop.Y) / 2 - longwiseCorrection;
                labelPosition = transGeom.CreatePoint2d(labelX, labelY);
            }
            else if (position == "top")
            {
                float labelX = (float)this.leftBottom.X - transverseCorrection;
                float labelY = (float)this.leftTop.Y - longwiseCorrection;
                labelPosition = transGeom.CreatePoint2d(labelX, labelY);
            }
            else {
                System.Diagnostics.Debug.Assert(false, "Неверно указано положение осевой подписи");
            }

            TextBox axisLabel = this.sketch.TextBoxes.AddFitted(labelPosition, labelText);

            axisLabel.Rotation = labelAngle;
            axisLabel.FormattedText = string.Format("<StyleOverride FontSize='{0}'>{1}</StyleOverride>", fontSize, axisLabel.Text);

            this.sketch.ExitEdit();
        }

        public void DrawArrow(float x, float y, float angle, float arrowLength, float length = 0.5f, float arrowAngle = 20, float lineWeight = 0.01f) {
            //рисует стрелку. менять ничего не надо и вызывать ее тоже не надо
            angle = (float)(angle * Math.PI / 180);

            this.sketch.Edit();
            Point2d endPoint = this.transGeom.CreatePoint2d(x, y);
            Point2d startPoint = this.transGeom.CreatePoint2d(endPoint.X - arrowLength * Math.Cos(angle), endPoint.Y - arrowLength * Math.Sin(angle));
            SketchLine arrowLine = this.sketch.SketchLines.AddByTwoPoints(endPoint, startPoint);
            arrowLine.LineWeight = lineWeight;
            this.sketch.ExitEdit();

            this.DrawEndArrow(endPoint, angle, length, arrowAngle);
        }

        public void DrawXArrow(string position = "center", float arrowLength = 3, float longwiseOffset = 0, float traverseOffset = 0) {
            //строит стрелку параллельно оси х
            //arrowLength - длина стрелки
            //longwiseOffset - продольное смещение назад, отрицательное смещение - вперед (всегда против положительного направления оси)
            //traverseOffset - поперечное смещение. положительное - влево, отрицательное - вправо 
            float x = 0;
            float y = 0;
            if (position == "center")
            {
                x = (float)((this.leftBottom.X + this.rightBottom.X) / 2 - longwiseOffset);
                y = (float)(this.leftBottom.Y - traverseOffset);
            }
            else if (position == "right")
            {
                x = (float)(this.rightBottom.X - longwiseOffset);
                y = (float)(this.leftBottom.Y - traverseOffset);
            }
            else {
                System.Diagnostics.Debug.Assert(false, "Неправильно задано положение стрелки");
            }
            this.DrawArrow(x, y, 0, arrowLength);
        }

        public void DrawYArrow(string position = "center", float arrowLength = 3, float longwiseOffset = 0, float traverseOffset = 0)
        {
            float x = 0;
            float y = 0;
            if (position == "center")
            {
                x = (float)(this.leftBottom.X - traverseOffset);
                y = (float)((this.leftBottom.Y + this.leftTop.Y) / 2 - longwiseOffset);
            }
            else if (position == "top")
            {
                x = (float)(this.leftBottom.X - traverseOffset);
                y = (float)(this.leftTop.Y - longwiseOffset);
            }
            else
            {
                System.Diagnostics.Debug.Assert(false, "Неправильно задано положение стрелки");
            }
            this.DrawArrow(x, y, 90, arrowLength);
        }

        public void AddColor(byte r, byte g, byte b) {
            colors.Add(app.TransientObjects.CreateColor(r, g, b));
        }

        private void DrawEndArrow(Point2d endPoint, float angle, float length = 0.5f, float arrowAngle = 20) {
            arrowAngle = (float)(arrowAngle / 180 * Math.PI);
            float arrowHalfWidth = (float)(length * Math.Tan(arrowAngle / 2));

            this.sketch.Edit();
            Point2d ortogonOriginPoint = this.transGeom.CreatePoint2d(endPoint.X - length * Math.Cos(angle), endPoint.Y - length * Math.Sin(angle));
            Point2d topOriginPoint = this.transGeom.CreatePoint2d(ortogonOriginPoint.X - arrowHalfWidth * Math.Sin(angle),
                ortogonOriginPoint.Y + arrowHalfWidth * Math.Cos(angle));
            Point2d bottomOriginPoint = this.transGeom.CreatePoint2d(ortogonOriginPoint.X + arrowHalfWidth * Math.Sin(angle),
                ortogonOriginPoint.Y - arrowHalfWidth * Math.Cos(angle));

            SketchLine line1 = this.sketch.SketchLines.AddByTwoPoints(bottomOriginPoint, topOriginPoint);
            SketchLine line2 = this.sketch.SketchLines.AddByTwoPoints(bottomOriginPoint, endPoint);
            SketchLine line3 = this.sketch.SketchLines.AddByTwoPoints(topOriginPoint, endPoint);

            this.sketch.GeometricConstraints.AddCoincident((SketchEntity)line1.StartSketchPoint, (SketchEntity)line2);
            this.sketch.GeometricConstraints.AddCoincident((SketchEntity)line1, (SketchEntity)line2.StartSketchPoint);

            this.sketch.GeometricConstraints.AddCoincident((SketchEntity)line2.EndSketchPoint, (SketchEntity)line3);
            this.sketch.GeometricConstraints.AddCoincident((SketchEntity)line2, (SketchEntity)line3.EndSketchPoint);

            this.sketch.GeometricConstraints.AddCoincident((SketchEntity)line1.EndSketchPoint, (SketchEntity)line3);
            this.sketch.GeometricConstraints.AddCoincident((SketchEntity)line1, (SketchEntity)line3.StartSketchPoint);

            ObjectCollection arrowEdges = this.app.TransientObjects.CreateObjectCollection();
            arrowEdges.Add(line1);
            arrowEdges.Add(line2);
            arrowEdges.Add(line3);

            Profile arrowContour = this.sketch.Profiles.AddForSolid(ProfilePathSegments: arrowEdges);
            this.sketch.SketchFillRegions.Add(arrowContour);

            this.sketch.ExitEdit();

        }

        private delegate string TickConverterDelegate(float value);

        private void PlotPieceWise1Profile(float[] x, float[] y, Color color, float lineWeight = 1)
        {
            sketch.Edit();
            int cntX = x.GetLength(0);
            int cntY = y.GetLength(0);
            System.Diagnostics.Debug.Assert(cntX == cntY, "Размеры массивов не равны.");

            int cnt = cntX;
            Point2d[] pointArray = new Point2d[cnt];

            for (int i = 0; i != cnt; ++i)
            {
                pointArray[i] = transGeom.CreatePoint2d(ToPlotCoordX(x[i]), ToPlotCoordY(y[i]));
            }

            List<SketchLine> plotLines = new List<SketchLine>();
            for (int i = 0; i != pointArray.GetLength(0) - 1; ++i)
            {
                plotLines.Add(sketch.SketchLines.AddByTwoPoints(pointArray[i], pointArray[i + 1]));
            }

            foreach (SketchLine line in plotLines) {
                line.LineWeight = lineWeight;
                line.OverrideColor = color;
            }
            sketch.ExitEdit();
        }

        private float ToRealCoordY(float yPlot)
        {
            return (yPlot - initY) / yScale + yMin;
        }

        private float ToRealCoordX(float xPlot)
        {
            return (xPlot - initX) / xScale + xMin;
        }

        private float ToPlotCoordY(float yReal) {
            return (yReal - yMin) * yScale + initY;
        }

        private float ToPlotCoordX(float xReal) {
            return (xReal - xMin) * xScale + initX;
        }

        private void UpdateInternals() {
            xMin = xArray.Min(); ;
            xMax = xArray.Max();

            yMin = this.yArrays[0].Min();
            yMax = this.yArrays[0].Max();

            foreach (float[] vect in this.yArrays)
            {
                yMax = vect.Max() > yMax ? vect.Max() : yMax;
                yMin = vect.Min() < yMin ? vect.Min() : yMin;
            }

            SetXScale();
            SetYScale();
        }

        private void SetYScale() {
            yScale = height / (yMax - yMin);
        }

        private void SetXScale() {
            xScale = length / (xMax - xMin);
        }

        private Color GetColor(int curveNum) {
            if (curveNum < colors.Count) {
                return colors[curveNum];
            }
            return app.TransientObjects.CreateColor(0, 0, 0);
        }
        
    }
}
