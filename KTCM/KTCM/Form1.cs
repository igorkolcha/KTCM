using System.Data;
using System.Data.SQLite;

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

        #region drawing on the form OnPaint Point2D

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
        private int marginBottom = 25;

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


        private void Form1_Load(object sender, EventArgs e)
        {
            ConnectionDataBase.Connection(dataGridView1, "SELECT фамилия FROM шн");

            /*dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //dataGridView1.Columns["фамилия"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // Обновить при изменении размера формы
            this.Resize += (s, args) =>
            {
                dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            };*/
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // Проверяем, что клик по реальной ячейке
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;

            var cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            var value = cell.Value?.ToString() ?? string.Empty;

            groupBox1.Visible = true;

            textBox_GroupBox1.Visible = false;
            label_GroupBox1_Text.Text = "удалить фамилию";
            label_GroupBox1.Text = value;
            label_GroupBox1.Visible = true;

            bool hasText = !string.IsNullOrWhiteSpace(value);

            button_GroupBox1_Delete.Visible = hasText;
            button_GroupBox1_Exit.Visible = hasText;
            button_GroupBox1_Save.Visible = !hasText;
            textBox_GroupBox1.Visible = !hasText;
        }

        private void button_GroupBox1_Exit_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
        }

        private void textBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            groupBox1.Visible = true;

            textBox_GroupBox1.Visible = true;
            label_GroupBox1.Visible = false;
            label_GroupBox1_Text.Text = "введите фамилию";

            button_GroupBox1_Delete.Visible = false;
            button_GroupBox1_Exit.Visible = true;
            button_GroupBox1_Save.Visible = true;
            textBox_GroupBox1.Visible = true;
        }

        private void button_GroupBox1_Delete_Click(object sender, EventArgs e)
        {
            ConnectionDataBase.DeleteEmployee(dataGridView1, label_GroupBox1);

            ConnectionDataBase.Connection(dataGridView1, "SELECT фамилия FROM шн");
        }

        private void button_GroupBox1_Save_Click(object sender, EventArgs e)
        {
            ConnectionDataBase.AddEmployee(textBox_GroupBox1);
            textBox_GroupBox1.Text = "";

            ConnectionDataBase.Connection(dataGridView1, "SELECT фамилия FROM шн");
        }

        #region toolStripMenuItem
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '1'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '2'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '3'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '4'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '5'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '6'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '7'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }
        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '8'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }
        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '9'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '10'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '11'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            DataGridToExcel.DataGridDataTime(dataGridView3, "SELECT станции, фамилия, дата, начало, конец, ктсм FROM ктсм WHERE месяц = '12'");
            DataGridToExcel.DataGridViewToExcel(dataGridView3);
        }
        #endregion


        private void button5_Click(object sender, EventArgs e)
        {
            ConnectionDataBase.BeginWork(button5, dateTimePicker1, dataGridView1);
        }
    }
}
