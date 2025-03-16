using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OpenQA.Selenium.Support.UI;

namespace G2BwpfApp
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void FetchData_Click(object sender, RoutedEventArgs e)
        {
            // Selenium ChromeDriver 설정
            ChromeOptions options = new ChromeOptions();
            options.AddExcludedArgument("enable-logging");
            ChromeDriver driver = new ChromeDriver(options);

            try
            {
                // 날짜 선택
                string formattedStartDate = startDateTextBox.Text.Trim();
                string formattedEndDate = endDateTextBox.Text.Trim();

                // 검색 페이지 열기
                driver.Navigate().GoToUrl("https://www.g2b.go.kr:8101/ep/tbid/tbidFwd.do");

                // 검색어 입력
                string query = "PRA";
                IWebElement bidNm = driver.FindElement(By.Id("bidNm"));
                bidNm.Clear();
                bidNm.SendKeys(query);

                // 시작 날짜 입력
                IWebElement startDateInput = driver.FindElement(By.Id("fromBidDt"));
                startDateInput.Clear();
                startDateInput.SendKeys(formattedStartDate);

                // 종료 날짜 입력
                IWebElement endDateInput = driver.FindElement(By.Id("toBidDt"));
                endDateInput.Clear();
                endDateInput.SendKeys(formattedEndDate);

                // 검색 버튼 클릭
                IWebElement searchButton = driver.FindElement(By.ClassName("btn_mdl"));
                searchButton.Click();

                // 결과
                IWebElement results = driver.FindElement(By.ClassName("results"));
                IReadOnlyCollection<IWebElement> divList = results.FindElements(By.TagName("div"));

                List<string> resultsList = new List<string>();
                foreach (IWebElement div in divList)
                {
                    resultsList.Add(div.Text);
                    IReadOnlyCollection<IWebElement> aTags = div.FindElements(By.TagName("a"));
                    if (aTags.Count > 0)
                    {
                        foreach (IWebElement aTag in aTags)
                        {
                            string link = aTag.GetAttribute("href");
                            resultsList.Add(link);
                        }
                    }
                }

                List<List<string>> resultList = new List<List<string>>();
                for (int i = 0; i < resultsList.Count; i += 12)
                {
                    resultList.Add(resultsList.GetRange(i, Math.Min(12, resultsList.Count - i)));
                }

                // 검색 결과를 resultTextBox에 출력
                resultTextBox.Clear();
                foreach (var item in resultList)
                {
                    resultTextBox.AppendText(string.Join(", ", item) + Environment.NewLine);
                }

                // 엑셀로 저장
                string savePath = saveLocationTextBox.Text;
                SaveToExcel(resultList, savePath);
            }
            catch (Exception ex)
            {
                // 예외 발생 시 오류 메시지 표시
                resultTextBox.Text = $"오류 발생: {ex.Message}";
            }
            finally
            {
                // 드라이버 종료
                driver.Quit();
            }
        }

        private void SaveToExcel(List<List<string>> data, string filePath)
        {
            // 엑셀로 저장
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  // 라이선스 컨텍스트 설정
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("검색 결과");

                for (int row = 0; row < data.Count; row++)
                {
                    for (int col = 0; col < data[row].Count; col++)
                    {
                        worksheet.Cells[row + 1, col + 1].Value = data[row][col];
                    }
                }

                package.Save();
            }
        }

        private void SelectSaveLocation_Click(object sender, RoutedEventArgs e)
        {
            // 저장 경로 선택
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Excel 파일 (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                saveLocationTextBox.Text = saveFileDialog.FileName;
            }
        }

    }
}
