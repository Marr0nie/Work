using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using System.Reflection;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;

namespace zarplata1.Forms
{
    /// <summary>
    /// Логика взаимодействия для WindowSelectCodesN.xaml
    /// </summary>
    public partial class WindowSelectCodesN
    {
        public ObservableCollection<ChargeType> CTN;
        public string checkedCodes = "";
        public bool fromTC = false;

        public WindowSelectCodesN(ObservableCollection<ChargeType> codes, string s,bool b)
        {
            InitializeComponent();

            if (codes == null) // если список еще не существует
            {
                this.CTN = App.ZarplataDB.GetChargeTypeN(); // список имеющихся кодов загружается из базы
                chbSelectAll.IsChecked = true; // устанавливается галочка "Отметить все"
            }
            else // если передан непустой список
            {
                this.CTN = codes;
            }

            string[] split = s.Split(',');
            foreach (ChargeType code in this.CTN)
                code.SelectedForTabeluch = false;

            foreach (ChargeType code in this.CTN)
            {
                for (int i = 0; i < split.Length; i++)
                {
                    if (code.Code == split[i].Trim()) code.SelectedForTabeluch = true;
                }
            }

            dgCodes.ItemsSource = this.CTN; // таблица заполняется этим списком
            fromTC = b;
        }

        private void chbSelectAll_Checked(object sender, RoutedEventArgs e)
        {
            // если в другом CheckBox стоит галочка, то снять галочку, т.к. возможен только один вариант
            if (chbUnselectAll.IsChecked == true)
                chbUnselectAll.IsChecked = false;

            // все подразделения помечаются галочками
            foreach (ChargeType code in CTN)
                code.SelectedForTabeluch = true;
        }

        private void chbUnselectAll_Checked(object sender, RoutedEventArgs e)
        {
            // если в другом CheckBox стоит галочка, то снять галочку, т.к. возможен только один вариант
            if (chbSelectAll.IsChecked == true)
                chbSelectAll.IsChecked = false;

            // у всех подразделений убираются галочки
            foreach (ChargeType code in CTN)
                code.SelectedForTabeluch = false;
        }

        public void WritecbFill()
        {
            string Filename;
            Filename = Environment.CurrentDirectory + "\\..\\..\\..\\zarplata1\\********\\";
            StreamWriter stw = new StreamWriter(Filename + ".txt", false, System.Text.Encoding.GetEncoding(1251));
            stw.WriteLine(checkedCodes);
            stw.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (fromTC)
            {
                if (CTN.Count == CTN.Where(w => w.SelectedForTabeluch).Count()) // если выбраны все коды
                {
                    checkedCodes = "Все"; // в ТекстБоксе выбранных кодов будет отображаться слово "Все"
                }
                else if (CTN.Where(w => w.SelectedForTabeluch).Count() == 0)
                {
                    checkedCodes = "-"; // в ТекстБоксе выбранных кодов будет отображаться "-"
                }
                else // если выбраны не все коды
                {
                    foreach (ChargeType code in CTN.Where(w => w.SelectedForTabeluch == true))
                    {
                        // в ТекстБоксе выбранных кодов будет отображаться строка,
                        // составленная из кодов выбранных кодов через запятую
                        checkedCodes += (code.Code + ", ");
                    }
                    checkedCodes = checkedCodes.Substring(0, checkedCodes.Length - 1);
                }
                App.MainWindowApp.codes = CTN;
                WritecbFill();
            }   
            this.Close(); // закрытие окна
        }

        private void dgCodes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // меняется значение флажка
            (dgCodes.SelectedItem as ChargeType).SelectedForTabeluch =
                !(dgCodes.SelectedItem as ChargeType).SelectedForTabeluch;

            // убираются галочки "Отметить все" и "Снять отметки у всех"
            chbSelectAll.IsChecked = false;
            chbUnselectAll.IsChecked = false;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            if (!(fromTC))
            {
                if (CTN.Count == CTN.Where(w => w.SelectedForTabeluch).Count()) // если выбраны все коды
                {
                    checkedCodes = "Все"; // в ТекстБоксе выбранных кодов будет отображаться слово "Все"
                }
                else if (CTN.Where(w => w.SelectedForTabeluch).Count() == 0)
                {
                    checkedCodes = "-"; // в ТекстБоксе выбранных кодов будет отображаться "-"
                }
                else // если выбраны не все коды
                {
                    foreach (ChargeType code in CTN.Where(w => w.SelectedForTabeluch == true))
                    {
                        // в ТекстБоксе выбранных кодов будет отображаться строка,
                        // составленная из кодов выбранных кодов через запятую
                        checkedCodes += (code.Code + ", ");
                    }
                    checkedCodes = checkedCodes.Substring(0, checkedCodes.Length - 1);
                }
                App.MainWindowApp.codes = CTN;
            }
        }
    }
}
