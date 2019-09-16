using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HighSpeedRailBooking
{
    public partial class Form1 : Form
    {
        IWebDriver driver;
        int CheckCount = 0;
        int Success = 0;

        public Form1()
        {
            InitializeComponent();
            Initial();
            initCombobox();
        }

        private class Station
        {
            public string st_Name { get; set; }
            public string st_Value { get; set; }
        }
        private bool IsElementPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        private void Initial()
        {
            //隱藏command
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            ChromeOptions options = new ChromeOptions();

            if (ConfigurationManager.AppSettings["headless"] == "1")
            {
                options.AddArguments("headless"); //隱藏瀏覽器 以下設定要加不能元素會找步道幹
                options.AddArguments("disable-gpu"); 
                options.AddArguments("window-size=1200,1100");
            }
           
            // options.AddArguments("--start-maximized");
            //停用開發人員模式防止部分網站檔腳本
            options.AddExcludedArgument("enable-automation");
            options.AddAdditionalCapability("useAutomationExtension", false);

            if (ConfigurationManager.AppSettings["HideCommandPromptWindow"] == "1")
            {
                driver = new ChromeDriver(service,options);
            }
            else
            {
                driver = new ChromeDriver(options);
            }

                

            driver.Navigate().GoToUrl("https://irs.thsrc.com.tw/IMINT/");

            //找不到元素預設10秒
           // driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            //點選同意cooki
            driver.FindElement(By.XPath("//*[@id=\"btn-confirm\"]")).Click();

            Verify();
        }

        private void Verify()
        {
            //截圖驗證碼                                        
            IWebElement element = driver.FindElement(By.XPath("//*[@id=\"BookingS1Form_homeCaptcha_passCode\"]"));
            //把整頁截圖並轉成位元
            Byte[] byteArray = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;
            Bitmap screenshot = new Bitmap(new System.IO.MemoryStream(byteArray));
            //切成正方形
            Rectangle croppedImage = new Rectangle(element.Location.X, element.Location.Y, element.Size.Width, element.Size.Height);
            Rectangle oDstRect = new Rectangle(new Point(0, 0), croppedImage.Size);
      
            System.Drawing.Image oDstImage = new System.Drawing.Bitmap(screenshot, croppedImage.Size);
            System.Drawing.Graphics oDC = System.Drawing.Graphics.FromImage(oDstImage);
            oDC.DrawImage(screenshot, oDstRect, croppedImage, GraphicsUnit.Pixel);
            //screenshot = screenshot.Clone(croppedImage, screenshot.PixelFormat); 影響效能改 Graphic
            oDstImage.Save("screenshot.png");
            oDstImage.Dispose();
            screenshot.Dispose();
            FileStream fs = new FileStream(Application.StartupPath + "/screenshot.png", FileMode.Open);
            pictureBox1.Image = System.Drawing.Image.FromStream(fs);
            fs.Close();
            fs.Dispose();
        }

        private void initCombobox()
        {
            List<Station> lis_DataList = new List<Station>()
            {
                new Station
                {
                    st_Name="請選擇"
                },
                 new Station
                {
                    st_Name="南港",
                    st_Value="1"
                } ,
                 new Station
                {
                    st_Name="台北",
                    st_Value="2"
                } ,
                 new Station
                {
                    st_Name="板橋",
                    st_Value="3"
                } ,
                 new Station
                {
                    st_Name="桃園",
                    st_Value="4"
                },
                new Station
                {
                    st_Name="新竹",
                    st_Value="5"
                },
                new Station
                {
                    st_Name="苗栗",
                    st_Value="6"
                },
                 new Station
                {
                    st_Name="台中",
                    st_Value="7"
                },
                 new Station
                {
                    st_Name="彰化",
                    st_Value="8"
                },
                 new Station
                {
                    st_Name="雲林",
                    st_Value="9"
                },
                 new Station
                {
                    st_Name="嘉義",
                    st_Value="10"
                },
                 new Station
                {
                    st_Name="台南",
                    st_Value="11"
                },
                 new Station
                {
                    st_Name="左營",
                    st_Value="12"
                }
            };





            comboBoxStart.DataSource = lis_DataList;
            comboBoxStart.DisplayMember = "st_Name";
            comboBoxStart.ValueMember = "st_Value";

            comboBoxEnd.BindingContext = new BindingContext();
            comboBoxEnd.DataSource = lis_DataList;
            comboBoxEnd.DisplayMember = "st_Name";
            comboBoxEnd.ValueMember = "st_Value";

            comboBox3.SelectedIndex = 1;
            comboBox4.SelectedIndex = 0;
            comboBox5.SelectedIndex = 0;
            comboBox6.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;

            radioButton1.Checked = true;
            radioButton3.Checked = true;
        }

        private void buttonReset_Click(object sender, EventArgs e)
        {
            driver.FindElement(By.XPath("//*[@id=\"BookingS1Form_homeCaptcha_reCodeLink\"]")).Click();
            Thread.Sleep(500);
            Verify();
        }


        private void buttonBook_Click(object sender, EventArgs e)
        {

            if (Success == 1)
            {
                MessageBox.Show("訂票成功,如要重新訂票,請重開程式!!");
                return;
            }

            //選擇起訖車廂
            if (comboBoxStart.SelectedIndex > 0)
            {
                var select = driver.FindElement(By.XPath("//*[@id=\"content\"]/tbody/tr[1]/td[2]/span/select"));
                var selectElement = new SelectElement(select);
                selectElement.SelectByValue(comboBoxStart.SelectedValue.ToString());
            }

            if (comboBoxEnd.SelectedIndex > 0)
            {
                var select = driver.FindElement(By.XPath("//*[@id=\"content\"]/tbody/tr[1]/td[2]/select"));
                var selectElement = new SelectElement(select);
                selectElement.SelectByValue(comboBoxEnd.SelectedValue.ToString());
            }



            //選擇車廂

            if (radioButton1.Checked && CheckCount==1)
            {
                CheckCount = 0;
                driver.FindElement(By.Id("trainCon:trainRadioGroup_0")).Click();
            }
            if (radioButton2.Checked && CheckCount == 0)
            {
                CheckCount = 1;
                driver.FindElement(By.Id("trainCon:trainRadioGroup_1")).Click();
            }

            //選擇車廂

            //座位喜好
            if (radioButton3.Checked)
            {
                driver.FindElement(By.XPath("//*[@id=\"seatRadio0\"]")).Click();
            }
            if (radioButton4.Checked)
            {
                driver.FindElement(By.XPath("//*[@id=\"seatRadio1\"]")).Click();
            }
            if (radioButton5.Checked)
            {
                driver.FindElement(By.XPath("//*[@id=\"seatRadio2\"]")).Click();
            }

            //時間
            driver.FindElement(By.XPath("//*[@id=\"toTimeInputField\"]")).Clear();
            driver.FindElement(By.XPath("//*[@id=\"toTimeInputField\"]")).SendKeys(dateTimePicker1.Value.ToString("yyyy/MM/dd"));

            //車次號碼
            driver.FindElement(By.XPath("//*[@id=\"bookingMethod_1\"]")).Click();
            driver.FindElement(By.XPath("//*[@id=\"toTrainID\"]/input")).Clear();
            driver.FindElement(By.XPath("//*[@id=\"toTrainID\"]/input")).SendKeys(textBox1.Text);

            //票數
            //全票
            var select1 = driver.FindElement(By.XPath("//*[@id=\"content\"]/tbody/tr[6]/td[2]/span/span[1]/span/select"));
            var selectElement1 = new SelectElement(select1);
            selectElement1.SelectByText(comboBox3.SelectedItem.ToString());
            //孩童票(6-11歲)
            var select2 = driver.FindElement(By.XPath("//*[@id=\"content\"]/tbody/tr[6]/td[2]/span/span[2]/span/select"));
            var selectElement2 = new SelectElement(select2);
            selectElement2.SelectByText(comboBox4.SelectedItem.ToString());
            //愛心票
            var select3 = driver.FindElement(By.XPath("//*[@id=\"content\"]/tbody/tr[6]/td[2]/span/span[3]/span/select"));
            var selectElement3 = new SelectElement(select3);
            selectElement3.SelectByText(comboBox5.SelectedItem.ToString());
            //敬老票(65歲以上)
            var select4 = driver.FindElement(By.XPath("//*[@id=\"content\"]/tbody/tr[6]/td[2]/span/span[4]/span/select"));
            var selectElement4 = new SelectElement(select4);
            selectElement4.SelectByText(comboBox6.SelectedItem.ToString());
            //大學生優惠票
            var select5 = driver.FindElement(By.XPath("//*[@id=\"content\"]/tbody/tr[6]/td[2]/span/span[5]/span/select"));
            var selectElement5 = new SelectElement(select5);
            selectElement5.SelectByText(comboBox7.SelectedItem.ToString());

            //驗證碼
            driver.FindElement(By.XPath("//*[@id=\"action\"]/tbody/tr/td/span[3]/input")).Clear();
            driver.FindElement(By.XPath("//*[@id=\"action\"]/tbody/tr/td/span[3]/input")).SendKeys(textBoxverify.Text);

            //開始查詢
            driver.FindElement(By.XPath("//*[@id=\"SubmitButton\"]")).Click();

            //顯示有誤的條件
            if (IsElementPresent(By.XPath("//*[@id=\"error\"]/span/ul")))
            {
                var text = driver.FindElement(By.XPath("//*[@id=\"error\"]/span/ul"));
                labelResult.Text = text.Text.Replace(",", ",/n/r");
                Success = 0;
            }

            //顯示選擇的車票價錢
            if (IsElementPresent(By.XPath("//*[@id=\"content\"]/span[1]/table/tbody/tr[3]"))) 
            {
                var result = driver.FindElement(By.XPath("//*[@id=\"content\"]/span[1]/table/tbody/tr[3]"));
                labelResult.Text = result.Text.Replace(",",",/n/r");
                Success = 1;
            }

            if (Success == 1)
            {
                driver.Quit();
            }



        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {   
            if(Success == 0)
                driver.Quit();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            comboBox7.SelectedIndex = 0;
            comboBox7.Enabled = false;
            comboBox7.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox7.SelectedIndex = 0;
            comboBox7.Enabled = true;
            comboBox7.DropDownStyle = ComboBoxStyle.DropDown;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex > 1 || comboBox4.SelectedIndex > 0 || comboBox5.SelectedIndex > 0 || comboBox6.SelectedIndex > 0 || comboBox7.SelectedIndex > 0)
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = false;
                radioButton4.Enabled = false;
                radioButton5.Enabled = false;
            }
            else
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = true;
                radioButton4.Enabled = true;
                radioButton5.Enabled = true;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex > 1 || comboBox4.SelectedIndex > 0 || comboBox5.SelectedIndex > 0 || comboBox6.SelectedIndex > 0 || comboBox7.SelectedIndex > 0)
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = false;
                radioButton4.Enabled = false;
                radioButton5.Enabled = false;
            }
            else
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = true;
                radioButton4.Enabled = true;
                radioButton5.Enabled = true;
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex > 1 || comboBox4.SelectedIndex > 0 || comboBox5.SelectedIndex > 0 || comboBox6.SelectedIndex > 0 || comboBox7.SelectedIndex > 0)
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = false;
                radioButton4.Enabled = false;
                radioButton5.Enabled = false;
            }
            else
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = true;
                radioButton4.Enabled = true;
                radioButton5.Enabled = true;
            }
        }
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedIndex > 1 || comboBox4.SelectedIndex > 0 || comboBox5.SelectedIndex > 0 || comboBox6.SelectedIndex > 0 || comboBox7.SelectedIndex > 0)
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = false;
                radioButton4.Enabled = false;
                radioButton5.Enabled = false;
            }
            else
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = true;
                radioButton4.Enabled = true;
                radioButton5.Enabled = true;
            }
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
           if(comboBox3.SelectedIndex > 1 || comboBox4.SelectedIndex > 0 || comboBox5.SelectedIndex > 0 || comboBox6.SelectedIndex > 0 || comboBox7.SelectedIndex > 0)
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = false;
                radioButton4.Enabled = false;
                radioButton5.Enabled = false;
            }
            else
            {
                radioButton3.Checked = true;
                radioButton3.Enabled = true;
                radioButton4.Enabled = true;
                radioButton5.Enabled = true;
            }
        }


    }
}
