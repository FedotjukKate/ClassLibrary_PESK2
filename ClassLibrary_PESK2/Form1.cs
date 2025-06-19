using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Eplan.EplApi.DataModel.Graphics.Color;
namespace ClassLibrary_PESK2
{
    public partial class Form1 : Form
    {
        private const string LabelPrefix = "dynamicLabel_"; // Префикс для имен динамических Label
        private int Label6Y = 265;

        public Form1()
        {
            InitializeComponent();
        }
        private void comboBoxI_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            int i = Convert.ToInt32(comboBox5.SelectedItem);
            int X = Convert.ToInt32(label5.Location.X);
            int Y = Convert.ToInt32(label5.Location.Y);

            ClearDynamicLabels();

            for (int j = 1; j <= i; j++)
            {
                Label newLabel = new Label();
                // Устанавливаем текст Label
                newLabel.Text = "Выбрано: " + i;
                // Устанавливаем положение Label 
                newLabel.Location = new System.Drawing.Point(X, Y + 45); // Смещение от верхнего левого угла формы
                // Устанавливаем автоматический размер Label
                newLabel.AutoSize = true;
                newLabel.Name = LabelPrefix + Convert.ToString(i);
                // Добавляем Label на форму
                this.Controls.Add(newLabel);

                //Двигаем нижний Label
                label6.Top = label6.Top + 45;
            }
        }

        // Метод для очистки динамически созданных Label
        private void ClearDynamicLabels()
        {
            for (int i = this.Controls.Count - 1; i >= 0; i--)
            {
                Control control = this.Controls[i];
                if (control is Label && control.Name.StartsWith(LabelPrefix))
                {
                    this.Controls.Remove(control);
                    control.Dispose();
                }
            }
            label6.Top = Label6Y;
        }
    }
}
