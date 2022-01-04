using OpenQA.Selenium.Chrome;
using RPACorreios;

class Program
{
    static void Main(string[] args)
    {
        string url = "https://buscacepinter.correios.com.br/app/endereco/index.php";

        List<string> numeroCEP = new List<string>()
        {
            "17014900" // Preencher com o valor do CEP desejado, nesse caso estou utilizando da prefeitura de Bauru
        };

        ChromeDriver driver = new ChromeDriver();
        driver.Navigate().GoToUrl(url);
        driver.Manage().Window.Maximize();

        foreach (var endereco in numeroCEP)
        {
            var pesquisaEl = driver.FindElement(OpenQA.Selenium.By.Id("endereco"));
            pesquisaEl.SendKeys(endereco);

            var btnBuscar = driver.FindElement(OpenQA.Selenium.By.XPath($"//button[@id='btn_pesquisar']"));
            btnBuscar.Click();

            Thread.Sleep(3000);

            // Atribuindo valores da tabela de pesquisa para as variáveis 

            var logradouroValue = driver.FindElement(OpenQA.Selenium.By.XPath("//*[@id='resultado-DNEC\']/tbody/tr/td[1]"));
            var bairroValue = driver.FindElement(OpenQA.Selenium.By.XPath("//*[@id='resultado-DNEC\']/tbody/tr/td[2]"));
            var localidadeValue = driver.FindElement(OpenQA.Selenium.By.XPath("//*[@id='resultado-DNEC\']/tbody/tr/td[3]"));
            var cepValue = driver.FindElement(OpenQA.Selenium.By.XPath("//*[@id='resultado-DNEC\']/tbody/tr/td[4]"));

            SendExcel excelValues = new SendExcel();

            excelValues.retornoTabela(logradouroValue.Text, bairroValue.Text, localidadeValue.Text, cepValue.Text);
        }
    }
}
