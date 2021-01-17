using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using FileHelpers;
using System.Reflection;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;


namespace SeleniumPrueba
{
    class Program
    {
        static void Main(string[] args)
        {
            string Mensaje = "";
            Procesos Data = new Procesos();
            
            //Ir a Mercado Libre
            Data.ExtraerData("MercadoLibre", out Mensaje);

            //Ir a Amazon
            Data.ExtraerData("Amazon", out Mensaje);

        }

    }

    class Procesos
    {


        #region AbrirExcel
        private FileInfo OpenWorkbook(string ruta)
        {
            DirectoryInfo outputDir = new DirectoryInfo(ruta);
            FileInfo newFile = new FileInfo(outputDir.FullName);
            return newFile;
        }
        #endregion

        #region ExtraerData
        public void ExtraerData(string Pagina, out string mensaje)
        {
            int fila = 3;
            string Producto = "";
            string Precio = "";
            string ruta = @"D:\Documents\Prueba\Lista de Productos.xlsx";
            string hoja = "Productos";
            try
            {
                //Abrir Excel
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(OpenWorkbook(ruta)))
                {
                    //Abrir Hoja
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[hoja];
                    
                    while (worksheet.Cells["A" + fila].Value != null)
                    {
                        //Extraer Valor Producto Columna A, Excel
                        Producto = worksheet.Cells["A" + fila].Value.ToString();

                        if (Pagina.Equals("Amazon"))
                        {
                            //Ingresar Pagina Amazon
                            IngresarPaginaAmazon(Producto, out Precio);
                            //Llenar resultado encontrado precio
                            worksheet.Cells["C" + fila].Value = Precio;
                            package.Save();
                            Thread.Sleep(2000);
                            
                        }
                        else
                        {
                            //Ingresar pagina Mercado Libre
                            IngresarPaginaMC(Producto, out Precio);
                            //Llenar resultado encontrado precio
                            worksheet.Cells["B" + fila].Value = Precio;
                            package.Save();
                            Thread.Sleep(2000);
                        }

                        fila++;
                    }
                }
                mensaje = "Termino Proceso Correctamente";
            }
            catch (Exception ex)
            {
                mensaje = "Ocurrio un error inesperado " + ex.Message;
            }
        }
        #endregion

        #region IngresarPaginaMercadoLibre
        public void IngresarPaginaMC(string Producto, out string Precio)
        {
            ////INGRESAR A LA PAGINA WEB DESDE FIREFOX
            IWebDriver driver = new FirefoxDriver();
            driver.Navigate().GoToUrl("https://www.mercadolibre.com.co/");
            driver.Manage().Window.Maximize(); //*

            //ESCRIBIR EN EL BUSCADOR
            //masthead-search-term
            IWebElement input = driver.FindElement(By.ClassName("nav-search-input"));
            input.SendKeys(Producto);

            //DAR CLICK EN BUSCADOR
            //search-btn
            IWebElement btnSearch = driver.FindElement(By.ClassName("nav-search-btn"));
            btnSearch.Click();

            //TRAER PRECIO PRIMER RESULTADO
            try
            {
                //if (driver.FindElement(By.ClassName("price-tag ui-search-price__part ui-search-price__original-value price-tag__disabled")).Equals(""))
                if (Validaciones.Exists(Validaciones.FindElementSafe(driver, By.ClassName("price-tag ui-search-price__part ui-search-price__original-value price-tag__disabled"))))
                {
                    IWebElement output = driver.FindElement(By.ClassName("price-tag ui-search-price__part ui-search-price__original-value price-tag__disabled"));
                    Precio = output.Text;
                }
                else
                {
                    IWebElement output = driver.FindElement(By.ClassName("ui-search-price__second-line"));
                    Precio = output.Text;
                    IWebElement PrecioCon = output.FindElement(By.ClassName("price-tag-fraction"));
                    Precio = PrecioCon.Text;
                }

            }
            catch (Exception)
            {
                IWebElement output = driver.FindElement(By.ClassName("ui-search-price__second-line"));
                Precio = output.Text;
                IWebElement PrecioCon = output.FindElement(By.ClassName("price-tag-fraction"));
                Precio = PrecioCon.Text;

            }
            driver.Close();
        }
        #endregion

        #region IngresarPaginaAmazon
        public void IngresarPaginaAmazon(string Producto, out string Precio)
        {
            ////INGRESAR A LA PAGINA WEB DESDE FIREFOX
            IWebDriver driver = new FirefoxDriver();
            driver.Navigate().GoToUrl("https://www.amazon.com/");
            driver.Manage().Window.Maximize(); //*

            //ESCRIBIR EN EL BUSCADOR
            //masthead-search-term
            IWebElement input = driver.FindElement(By.Id("twotabsearchtextbox"));
            input.SendKeys(Producto);

            //DAR CLICK EN BUSCADOR
            //search-btn
            IWebElement btnSearch = driver.FindElement(By.Id("nav-search-submit-button"));
            btnSearch.Click();

            //TRAER PRECIO PRIMER RESULTADO
            IWebElement output = driver.FindElement(By.ClassName("a-price"));
            Precio = output.Text;
            driver.Close();

        }
        #endregion

        
    }

    public static class Validaciones
    {
        public static IWebElement FindElementSafe(this IWebDriver driver, By by)
        {
            try
            {
                return driver.FindElement(by);
            }
            catch (NoSuchElementException)
            {
                return null;
            }
        }

        public static bool Exists(this IWebElement element)
        {
            if (element == null)
            {
                return false;
            }
            return true;
        }
    }

}
