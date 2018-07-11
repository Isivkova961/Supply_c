using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;


namespace Supply
{
    public partial class fSupply : Form
    {
        List<TextBox> tbNumberList = new List<TextBox>();
        List<RichTextBox> rtbProductList = new List<RichTextBox>();
        List<TextBox> tbQuantityList = new List<TextBox>();
        List<TextBox> tbEdList = new List<TextBox>();
        List<TextBox> tbPriceList = new List<TextBox>();
        List<TextBox> tbAmountList = new List<TextBox>();
        string[] NumEd = new string[] {"один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять", "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", 
            "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"};
        string[] NumEd1 = new string[] { "одна", "две"};
        string[] NumDec = new string[] { "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто" };
        string[] NumSot = new string[] { "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятсот" };
        string[] XLion0 = new string[] { "тысяч", "миллионов", "миллиардов" };
        string[] XLion1 = new string[] { "тысяча", "миллион", "миллиард" };
        string[] XLion2 = new string[] { "тысячи", "миллиона", "миллиарда" };
        string[] Rub = new string[] { "рубль", "рубля", "рублей" };
        string[] Mes = new string[] { "январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь" };
        string sSummaItog;
        string PathName;

        public fSupply()
        {
            InitializeComponent();
            rbUrL.Checked = true;
        }

        private void bClearField_Click(object sender, EventArgs e)
        {
            NewData();
        }

        private void NewData()
        {
            //Очистка полей Данных договора и поставки
            tbNumDoc.Text = "";
            tbShipment.Text = "";
            tbDelivery.Text = "";

            //Очистка полей Данных покупателя Юр.лица
            tbNameWork.Text = "";
            tbShortName.Text = "";
            tbVLice.Text = "";
            tbBase.Text = "";
            tbUrAdres.Text = "";
            tbINN.Text = "";
            tbKPP.Text = "";
            tbRS.Text = "";
            tbKorS.Text = "";
            tbBIK.Text = "";
            tbNameBank.Text = "";
            tbTelefonUr.Text = "";
            tbEmail.Text = "";

            //Очистка полей Данных покупателя Физ. лица
            tbFIO.Text = "";
            tbDocument.Text = "";
            tbSeria.Text = "";
            tbNumber.Text = "";
            tbKemIssu.Text = "";
            tbAdres.Text = "";
            tbTelefon.Text = "";   
      
            //Очистка полей Спецификации
            tbNumber1.Text = "";
            rtbProduct1.Text = "";
            tbQuantity1.Text = "";
            tbEd1.Text = "";
            tbPrice1.Text = "";
            tbAmount1.Text = "";
            tbItogo.Text = "";
            tbNDS.Text = "";

            //Удаляем динамически созданные компоненты
            for (int i = tbNumberList.Count - 1; i >= 0; i--)
            {
                tpSpecif.Controls.Remove(tbNumberList[i] as TextBox);
                tbNumberList.Remove(tbNumberList[i]);

                tpSpecif.Controls.Remove(rtbProductList[i] as RichTextBox);
                rtbProductList.Remove(rtbProductList[i]);

                tpSpecif.Controls.Remove(tbQuantityList[i] as TextBox);
                tbQuantityList.Remove(tbQuantityList[i]);

                tpSpecif.Controls.Remove(tbEdList[i] as TextBox);
                tbEdList.Remove(tbEdList[i]);

                tpSpecif.Controls.Remove(tbPriceList[i] as TextBox);
                tbPriceList.Remove(tbPriceList[i]);

                tpSpecif.Controls.Remove(tbAmountList[i] as TextBox);
                tbAmountList.Remove(tbAmountList[i]);
            }

            //Очистить списки компонент

            


        }

        private void rbUrL_CheckedChanged(object sender, EventArgs e)
        {
            lNDS.Visible = rbUrL.Checked;
            tbNDS.Visible = rbUrL.Checked;

            ChangeState();
            PathName = "Договор ЮЛ.docx";
        }

       //Процедура замена в документе данных закладок на текст из программы
        internal void BookMarkReplaceNative(Word.Bookmark bookmark, string newText)
        {
           object rng = bookmark.Range;
           string bookmarkName = bookmark.Name;

           bookmark.Range.Text = newText;

         //  Word.Document document = this.Application.ActiveDocument;
         //  document.Bookmarks.Add(bookmarkName, ref rng);
           
        }

        //Работа с отчетом в Word
        private void ReportWord()
        {
            object rng;
            string bookmarkName;                        

            Word._Application application;
            Word._Document document;

            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;

            //Создаем объект приложения Word
            application = new Word.Application();
            //Создаем путь к файлу
            Object templatePathObj = Directory.GetCurrentDirectory() +"\\" + PathName;
            //Обработчик ошибок
            try
            {
                //Создаем документ на основе шаблона               
                document = application.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);                            
            }
            catch (Exception error)
            {
                //document.Close(ref falseObj, ref missingObj, ref missingObj);
                application.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                application = null;
                throw error;
            }
            //Проверяем сколько нужно вставить строк в таблицу спецификации
            for (int i = 1; i <= tbNumberList.Count; i++)
            {
                document.Tables[1].Rows.Add(document.Tables[1].Rows[i + 2]);
            }
                //Заменяем закладки в документе
            for (int i = document.Bookmarks.Count; i >= 1; i--)
            {
                bookmarkName = document.Bookmarks[i].Name;
                rng = document.Bookmarks[i].Range;

                //Закладки по вкладке Данные договора и поставки
                if (CompareBm(bookmarkName, "Номердоговора") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbNumDoc.Text;
                }
                if (CompareBm(bookmarkName, "Дата") == 0)
                {
                    document.Bookmarks[i].Range.Text = dtpDataDoc.Text;
                }
                if (CompareBm(bookmarkName, "Дней") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbShipment.Text;
                }
                if (CompareBm(bookmarkName, "Доставкакуда") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbDelivery.Text;
                }
                if (CompareBm(bookmarkName, "Где") == 0)
                {
                    string s = tbDelivery.Text;
                    if (s.IndexOf(" ") > -1)
                    {
                        string s1 = s.Substring(0, s.IndexOf(" ") - 1);
                        document.Bookmarks[i].Range.Text = s1;
                    }
                    else
                    {
                        document.Bookmarks[i].Range.Text = s;
                    }
                }
                //Закладки по вкладке Данные покупателя Юр.лицо
                if (CompareBm(bookmarkName, "Наименование") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbNameWork.Text;
                }
                if (CompareBm(bookmarkName, "КорНаименование") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbShortName.Text;
                }
                if (CompareBm(bookmarkName, "Лице") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbVLice.Text;
                }
                if (CompareBm(bookmarkName, "Основание") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbBase.Text;
                }
                if (CompareBm(bookmarkName, "Юрадрес") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbUrAdres.Text;
                }
                if (CompareBm(bookmarkName, "ИНН") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbINN.Text;
                }
                if (CompareBm(bookmarkName, "КПП") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbKPP.Text;
                }
                if (CompareBm(bookmarkName, "Расчетныйсчет") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbRS.Text;
                }
                if (CompareBm(bookmarkName, "Коррсчет") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbKorS.Text;
                }
                if (CompareBm(bookmarkName, "БИК") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbBIK.Text;
                }
                if (CompareBm(bookmarkName, "Реквизитыбанка") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbNameBank.Text;
                }
                if (CompareBm(bookmarkName, "Телефон") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbTelefonUr.Text;
                }
                if (CompareBm(bookmarkName, "Емайл") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbEmail.Text;
                }
                //Закладки по вкладке данные покупателя Физ.лицо
                if (CompareBm(bookmarkName, "ФИО") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbFIO.Text;
                }
                if (CompareBm(bookmarkName, "Документ") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbDocument.Text;
                }
                if (CompareBm(bookmarkName, "Серия") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbSeria.Text;
                }
                if (CompareBm(bookmarkName, "НомерПас") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbNumber.Text;
                }
                if (CompareBm(bookmarkName, "Датавыдачи") == 0)
                {
                    document.Bookmarks[i].Range.Text = dtpIssu.Text;
                }
                if (CompareBm(bookmarkName, "Кем") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbKemIssu.Text;
                }
                if (CompareBm(bookmarkName, "Адрес") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbAdres.Text;
                }
                if (CompareBm(bookmarkName, "ТелефонФиз") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbTelefon.Text;
                }
                //Закладки по вкладке данные покупателя ИП
                if (CompareBm(bookmarkName, "Имяпокупателя") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbFIOIP.Text;
                }
                
                //Закладки по вкладке Спецификация
                if (CompareBm(bookmarkName, "Номер") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbNumber1.Text;
                    for (int j = 0; j <= tbNumberList.Count - 1; j++)
                    {
                        document.Tables[1].Cell(3 + j, 1).Range.Text = tbNumberList[j].Text;
                    }
                }
                if (CompareBm(bookmarkName, "Товар") == 0)
                {
                    document.Bookmarks[i].Range.Text = rtbProduct1.Text;
                    for (int j = 0; j <= rtbProductList.Count - 1; j++)
                    {
                        document.Tables[1].Cell(3 + j, 2).Range.Text = rtbProductList[j].Text;
                    }
                }
                if (CompareBm(bookmarkName, "Колво") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbQuantity1.Text;
                    for (int j = 0; j <= tbQuantityList.Count - 1; j++)
                    {
                        document.Tables[1].Cell(3 + j, 3).Range.Text = tbQuantityList[j].Text;
                    }
                }
                if (CompareBm(bookmarkName, "Ед") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbEd1.Text;
                    for (int j = 0; j <= tbEdList.Count - 1; j++)
                    {
                        document.Tables[1].Cell(3 + j, 4).Range.Text = tbEdList[j].Text;
                    }
                }
                if (CompareBm(bookmarkName, "Цена") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbPrice1.Text;
                    for (int j = 0; j <= tbPriceList.Count - 1; j++)
                    {
                        document.Tables[1].Cell(3 + j, 5).Range.Text = tbPriceList[j].Text;
                    }
                }
                if (CompareBm(bookmarkName, "Сумма") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbAmount1.Text;
                    for (int j = 0; j <= tbAmountList.Count - 1; j++)
                    {
                        document.Tables[1].Cell(3 + j, 6).Range.Text = tbAmountList[j].Text;
                    }
                }
                if (CompareBm(bookmarkName, "ИТОГО1") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbItogo.Text;
                }
                if (CompareBm(bookmarkName, "ИТОГО2") == 0)
                {
                    string s = tbItogo.Text;
                    int index = s.IndexOf(",");
                    string itog = s.Substring(0, index);
                    
                    itog = itog + " рублей " + s.Substring(index + 1, 2) + " коп.";

                    document.Bookmarks[i].Range.Text = itog;
                }
                if (CompareBm(bookmarkName, "ИТОГО3") == 0)
                {
                    string s = tbItogo.Text;
                    int index = s.IndexOf(",");
                    
                    string s1 = sSummaItog + " " + s.Substring(index + 1, 2) + " коп.";
                    document.Bookmarks[i].Range.Text = s1.Substring(0, 1).ToUpper()+ s1.Remove(0, 1);
                }
                if (CompareBm(bookmarkName, "НДС") == 0)
                {
                    document.Bookmarks[i].Range.Text = tbNDS.Text;
                }
                



            }
            application.Visible = true;
        }

        //Функция удаления с имен закладок все что стоит до _
        static int CompareBm(string ABmName, string AName)
        {
            int i = ABmName.IndexOf("_");
            string ABmName1 = ABmName;
            int ABmLength = ABmName.Length;

            if (i > 0)
            {
                ABmName1 = ABmName.Substring(0, i);              
            }

            int result = String.Compare(ABmName1, AName);            
            return result;
        }

        private void bFormDoc_Click(object sender, EventArgs e)
        {
            ReportWord();
        }

        private void bAdd_Click(object sender, EventArgs e)
        {
            if (tbNumberList.Count < 6)
            {
                TextBox tbNumber2 = new TextBox();
                tbNumber2.Location = new Point(23, 45 + 29 * (tbNumberList.Count + 1));
                tbNumber2.Size = new Size(41, 23);
                tbNumberList.Add(tbNumber2);
                tpSpecif.Controls.Add(tbNumber2);

                RichTextBox rtbProduct2 = new RichTextBox();
                rtbProduct2.Location = new Point(70, 45 + 29 * (rtbProductList.Count + 1));
                rtbProduct2.Size = new Size(376, 23);
                rtbProductList.Add(rtbProduct2);
                tpSpecif.Controls.Add(rtbProduct2);

                TextBox tbQuantity2 = new TextBox();
                tbQuantity2.Location = new Point(452, 45 + 29 * (tbQuantityList.Count + 1));
                tbQuantity2.Size = new Size(52, 23);
                tbQuantityList.Add(tbQuantity2);
                tpSpecif.Controls.Add(tbQuantity2);

                TextBox tbEd2 = new TextBox();
                tbEd2.Location = new Point(510, 45 + 29 * (tbEdList.Count + 1));
                tbEd2.Size = new Size(41, 23);
                tbEdList.Add(tbEd2);
                tpSpecif.Controls.Add(tbEd2);

                TextBox tbPrice2 = new TextBox();
                tbPrice2.Location = new Point(557, 45 + 29 * (tbPriceList.Count + 1));
                tbPrice2.Size = new Size(66, 23);
                tbPriceList.Add(tbPrice2);
                tpSpecif.Controls.Add(tbPrice2);

                TextBox tbAmount2 = new TextBox();
                tbAmount2.Location = new Point(629, 45 + 29 * (tbAmountList.Count + 1));
                tbAmount2.Size = new Size(83, 23);
                tbAmountList.Add(tbAmount2);
                tpSpecif.Controls.Add(tbAmount2);
            }
            else
            {
                MessageBox.Show("Больше добавлять нельзя!","Сообщение", MessageBoxButtons.OK);  
            }
        }

        private void bCountUp_Click(object sender, EventArgs e)
        {
            double summa = 0;
            string sAmount;
            //Расчет сумм
            if (tbQuantity1.Text != "" && tbPrice1.Text != "")
            {
                if (tbPrice1.Text.IndexOf(",") > 0)
                {
                    if ((tbPrice1.Text.Length - tbPrice1.Text.IndexOf(",")) <= 2)
                    {
                        tbPrice1.Text = tbPrice1.Text + "0";
                    }
                }
                else
                {
                    tbPrice1.Text = tbPrice1.Text + ",00";
                }

                sAmount = Convert.ToString(Convert.ToInt32(tbQuantity1.Text) * Convert.ToDouble(tbPrice1.Text));

                if (sAmount.IndexOf(",") > 0)
                {
                    if ((sAmount.Length - sAmount.IndexOf(",")) > 2)
                    {
                        tbAmount1.Text = sAmount;
                    }
                    else
                    {
                        tbAmount1.Text = sAmount + "0";
                    }
                }
                else
                {
                    tbAmount1.Text = sAmount + ",00";
                }

                summa = summa + Convert.ToDouble(tbAmount1.Text);
            }
            else
            {
                if (tbQuantity1.Text == "" && tbPrice1.Text == "")
                {
                    tbAmount1.Text = "0,00";
                }
                if (tbQuantity1.Text == "")
                {
                    tbQuantity1.Text = "0,00";
                }
                if (tbPrice1.Text == "")
                {
                    tbPrice1.Text = "0,00";
                }
                
            }

            for (int i = 0; i < tbNumberList.Count; i++)
            {
                if (tbQuantityList[i].Text != "" && tbPriceList[i].Text != "")
                {
                    if (tbPriceList[i].Text.IndexOf(",") > 0)
                    {
                        if ((tbPriceList[i].Text.Length - tbPriceList[i].Text.IndexOf(",")) <= 2)
                        {
                            tbPriceList[i].Text = tbPriceList[i].Text + "0";
                        }
                    }
                    else
                    {
                        tbPriceList[i].Text = tbPriceList[i].Text + ",00";
                    }

                    sAmount = Convert.ToString(Convert.ToInt32(tbQuantityList[i].Text) * Convert.ToDouble(tbPriceList[i].Text));

                    if (sAmount.IndexOf(",") > 0)
                    {
                        if ((sAmount.Length - sAmount.IndexOf(",")) > 2)
                        {
                            tbAmountList[i].Text = sAmount;
                        }
                        else
                        {
                            tbAmountList[i].Text = sAmount + "0";
                        }
                    }
                    else
                    {
                        tbAmountList[i].Text = sAmount + ",00";
                    }

                    summa = summa + Convert.ToDouble(tbAmountList[i].Text);

                }
                else
                {
                    if (tbQuantityList[i].Text == "" && tbPriceList[i].Text == "")
                    {
                        tbAmountList[i].Text = "0,00";
                    }
                    if (tbQuantityList[i].Text == "")
                    {
                        tbQuantityList[i].Text = "0,00";
                    }
                    if (tbPriceList[i].Text == "")
                    {
                        tbPriceList[i].Text = "0,00";
                    }
                    
                }
            }

            string sSumma = Convert.ToString(summa);
            int k = sSumma.IndexOf(",");
            string Itog;
                if (k > 0)
                {
                    Itog = sSumma.Substring(0, k);
                    if ((sSumma.Length - k) > 2)
                    {
                        tbItogo.Text = Convert.ToString(summa);
                    }
                    else
                    {
                        tbItogo.Text = sSumma + "0";
                    }
                }
                else
                {
                    Itog = sSumma;
                    tbItogo.Text = sSumma + ",00";
                }

            sSummaItog = ConvertNumToString(Convert.ToInt32(Itog));
        }

        private string ConvertNumToString(int Number)
        {
            string rub1, rub2;
            int i, k, index, number;
            i = - 1;
            k = 0;
            rub1 = "";
            rub2 = "";

            while (Number > 0)
            {
                k++;

                number = Number % 1000;
                Number = Number / 1000;

                i++;

                if ((number / 100) > 0)
                {
                    index = number / 100;
                    number = number % 100;
                    rub1 = rub1 + NumSot[index - 1] + " ";
                }

                if (number >= 20)
                {
                    if ((number / 10) > 0)
                    {
                        index = number / 10;
                        number = number % 10;

                        if (number > 0)
                        {
                            if (number < 3 && i == 1)
                            {
                                rub1 = rub1 + NumDec[index - 2] + " " + NumEd1[number - 1] + " ";
                            }
                            else
                            {
                                rub1 = rub1 + NumDec[index - 2] + " " + NumEd[number - 1] + " ";
                            }
                        }
                        else
                        {
                            rub1 = rub1 + NumDec[index - 2] + " ";
                        }
                    }
                }
                else
                {
                    if (number < 20 && number > 2)
                    {
                        rub1 = rub1 + NumEd[number - 1];
                    }
                    else
                    {
                        if (number < 3 && number > 0 && i == 1)
                        {
                            rub1 = rub1 + NumEd1[number - 1] + " ";
                        }
                        else
                        {
                            if (number < 3 && number > 0 && i != 1)
                            {
                                rub1 = rub1 + NumEd[number - 1] + " ";
                            }
                        }
                    }
                }

                if (k > 1)
                {
                    if (i > 0)
                    {
                        if (rub1 != "")
                        {
                            if (number == 1)
                            {
                                rub1 = rub1 + XLion1[i - 1] + " ";
                            }
                            if (number > 1 && number < 5)
                            {
                                rub1 = rub1 + XLion2[i - 1] + " ";
                            }
                            if (number >= 5 || number == 0)
                            {
                                rub1 = rub1 + XLion0[i - 1] + " ";

                            }
                        }
                    }
                }

                if (k == 1)
                {

                    if (number == 1)
                    {
                        rub2 = Rub[0];
                    }
                    else
                    {
                        if (number > 1 && number < 5)
                        {
                            rub2 = Rub[1];
                        }
                        else
                        {
                            if (number >= 5 || number == 0)
                            {
                                rub2 = Rub[2];
                            }
                        }
                    }
                }

                rub2 = rub1 + rub2;
                rub1 = "";
                
                
            }

            return rub2;



        }

        private void rbFizL_CheckedChanged(object sender, EventArgs e)
        {
            ChangeState();
            PathName = "Договор ФЛ.docx";
        }

        private void ChangeState()
        {
            //Если выбрали ЮЛ
            if (rbUrL.Checked == true)
            {
                tpDataFiz.Parent = null;
                tbDataIP.Parent = null;
                tpDataUr.Parent = tcSupply;

                //Чтобы вкладка спецификация выходила после вкладки Данных о покупателе
                tpSpecif.Parent = null;
                tpSpecif.Parent = tcSupply;

            }
            else
            {
                //Если выбрали ФЛ
                if (rbFizL.Checked == true)
                {
                    tpDataFiz.Parent = tcSupply;
                    tbDataIP.Parent = null;
                    tpDataUr.Parent = null;

                    //Чтобы вкладка спецификация выходила после вкладки Данных о покупателе
                    tpSpecif.Parent = null;
                    tpSpecif.Parent = tcSupply;
                }
                //Если выбрали ИП
                else
                {
                    tpDataFiz.Parent = null;
                    tbDataIP.Parent = tcSupply;
                    tpDataUr.Parent = null;

                    //Чтобы вкладка спецификация выходила после вкладки Данных о покупателе
                    tpSpecif.Parent = null;
                    tpSpecif.Parent = tcSupply;
                }

            }
        }

        private void rbIP_CheckedChanged(object sender, EventArgs e)
        {
            ChangeState();
            PathName = "Договор ИП.docx";
        }
    }
}
