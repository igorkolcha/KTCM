namespace KTCM
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            this.SetStyle(ControlStyles.ResizeRedraw, true);

            this.DoubleBuffered = true;

            this.BackColor = Color.Green;
        }

        #region Рисование формы OnPaint Point2D

        // Определите область рисования
        private System.Drawing.Rectangle PlotArea;

        // Единица определяется в мировой системе координат (логические координаты графика)
        private float xMin = 0f;
        private float xMax = 10f;
        private float yMin = 0f;
        private float yMax = 10f;

        // Координаты начала графика (отступ слева и сверху)
        private int x = 275;
        private int y = 25;
        // Отступы справа и снизу
        private int marginRight = 20;
        private int marginBottom = 40;

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e); // Обязательно вызываем базовый метод
            Graphics g = e.Graphics;

            // 1. Динамический расчет размеров
            // ClientSize.Width - это текущая ширина внутренней области окна
            // Вычитаем x (отступ слева) и marginRight (отступ справа)
            int currentWidth = this.ClientSize.Width - x - marginRight;
            int currentHeight = this.ClientSize.Height - y - marginBottom;

            // Защита от ошибок, если окно свернули или сделали слишком маленьким
            if (currentWidth <= 0 || currentHeight <= 0) return;

            // 2. Обновляем PlotArea новыми размерами
            PlotArea = new Rectangle(x, y, currentWidth, currentHeight);

            // 3. Рисуем рамку
            g.DrawRectangle(Pens.Black, PlotArea);

            // 4. Рисуем линии (они автоматически масштабируются, так как Point2D использует PlotArea)
            using (Pen aPen = new Pen(Color.White, 1))
            {
                // Горизонтальная линия (Y=5)
                g.DrawLine(aPen, Point2D(new PointF(0.5f, 5)), Point2D(new PointF(9.5f, 5)));

                // Вертикальная линия (X=5)
                g.DrawLine(aPen, Point2D(new PointF(5, 0.5f)), Point2D(new PointF(5, 9.5f)));
            }

            // Примечание: g.Dispose() вызывать НЕЛЬЗЯ, так как Graphics предоставлен системой через PaintEventArgs
        }

        private PointF Point2D(PointF ptf)
        {
            PointF aPoint = new PointF();
            // Формула преобразования координат
            // Если PlotArea.Width и Height меняются, то и результат этой формулы изменится автоматически
            aPoint.X = PlotArea.X + (ptf.X - xMin) * PlotArea.Width / (xMax - xMin);

            // Обратите внимание: координата Y в WinForms растет вниз, поэтому вычитаем из Bottom
            aPoint.Y = PlotArea.Bottom - (ptf.Y - yMin) * PlotArea.Height / (yMax - yMin);

            return aPoint;
        }
        #endregion
    }
}
