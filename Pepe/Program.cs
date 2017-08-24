using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Pepe.Model;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.IO;
namespace Pepe
{
    class Program
    {
        public static void ErrorLogging(Exception ex)
        {
            string strPath = "Log.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
            Email.Email.SendEmail("lightbot@lightsourcehr.com", "xeeshan.ah@gmail.com", "Error log", ex.Message, strPath);
        }
        public static void ErrorLogging2(Exception ex)
        {
            string strPath = "Log.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
            Email.Email.SendEmail("lightbot@lightsourcehr.com", "certs@lightsourcehr.com", "Error log", ex.Message, strPath);
        }
        public static void ErrorLogging1(Exception ex)
        {
            string strPath = "Log.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
            Email.Email.SendEmail("lightbot@lightsourcehr.com", "ba@lightsourcehr.com", "Error log", ex.Message, strPath,"data.txt");
        }
        static void Main(string[] args)
        {
            List<Record> Records = new List<Record>();
            List<Record> Importable = new List<Record>();
            List<Record> cases = new List<Record>();
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddUserProfilePreference("download.default_directory", basePath);
            chromeOptions.AddUserProfilePreference("intl.accept_languages", "nl");
            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
            wait:
            Console.WriteLine("System going to wait ...");
            while(true)
            {
                string wait = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();
                if (wait == "2340")
                {
                    break;
                }
            }
            Console.WriteLine("bot process started...");
            Records.Clear();
            Importable.Clear();
            cases.Clear();
            IWebDriver gc = new ChromeDriver(chromeOptions);
            Console.ForegroundColor = ConsoleColor.White;
            try
            {
                //...loging in 
                gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/Login");
                gc.FindElement(By.Name("LoginID")).SendKeys("lightbot");
                gc.FindElement(By.Name("Password")).SendKeys("RPAuser!");
                gc.FindElement(By.Name("Password")).SendKeys(Keys.Enter);
                //.............

            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Login int to client space failed");
            }
            //redireciting to reports one
            try
            {
                gc.Navigate().GoToUrl("https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\AI1+Ancillary+Risk+Fees");
                System.Threading.Thread.Sleep(4000);
            }
            catch
            {
                
            }
            //extracting data from report Al1 risk fee
            try
            {
                Console.WriteLine("Scrapping data :");
                //getting reports from alteranating rows
                foreach (IWebElement c in gc.FindElements(By.ClassName("AlternatingItem")))
                {
                    Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text);
                    Records.Add(r);
                }
                //getting records from normal rows
                foreach (IWebElement c in gc.FindElements(By.ClassName("ReportItem")))
                {
                    Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text);
                    Records.Add(r);
                }
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Data scrapping failed !");
                Console.ForegroundColor = ConsoleColor.White;
            }

            //..........
            //...........
            //moving to al2 report
            //redireciting to reports one
            try
            {
                gc.Navigate().GoToUrl("https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\\AI2+Ancillary+Risk+Fees");
                System.Threading.Thread.Sleep(4000);
            }
            catch
            {

            }
            //extracting data from report Al1 risk fee
            try
            {
                Console.WriteLine("Scrapping data :");
                //getting reports from alteranating rows
                foreach (IWebElement c in gc.FindElements(By.ClassName("AlternatingItem")))
                {
                    Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text);
                    Records.Add(r);
                }
                //getting records from normal rows
                foreach (IWebElement c in gc.FindElements(By.ClassName("ReportItem")))
                {
                    Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text);
                    Records.Add(r);
                }
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Data scrapping from \"Al\" failed !");
                Console.ForegroundColor = ConsoleColor.White;
            }

            Console.WriteLine("Data extracted from reports :\n");
            foreach(Record r in Records)
            {
                if(r.location==null || r.location=="")
                {
                    cases.Add(r);
                }
                else
                {
                    Importable.Add(r);
                }
                r.printData();
            }

            Console.WriteLine("\n Number of importable Records :" + Importable.Count);
            Console.WriteLine("\n Number of records for cases :" + cases.Count);
            if(cases.Count>0)
            {
                foreach (Record record in cases)
                {
                    Exception ex = new Exception(record.toString());
                    ErrorLogging2(ex);
                 
                    try
                    {
                        gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/peo/client");
                        gc.FindElement(By.Id("dropdownMenu1")).Click();
                        gc.FindElement(By.Id("All")).Click();
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ClientNumber")).SendKeys(record.clientId.ToString()); // put the client number here.
                        System.Threading.Thread.Sleep(2000);
                        gc.FindElement(By.ClassName("formSearchBtn")).Click();
                        System.Threading.Thread.Sleep(1000);
                        gc.FindElements(By.ClassName("cs-underline"))[1].Click();
                        gc.FindElement(By.XPath("//*[@id='customHeaderHtml']/div[2]/div[6]/div/div[1]/table/tbody/tr/td[1]/span")).Click();
                        gc.FindElement(By.XPath("//*[@id='lstDataform_btnAdd']")).Click();
                        gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys("R");
                        System.Threading.Thread.Sleep(1500);
                        gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys(Keys.Tab);
                        gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys("M");
                        System.Threading.Thread.Sleep(1500);
                        gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys(Keys.Enter);
                        gc.FindElement(By.XPath("//*[@id='CallerName']")).SendKeys("Lightbot");
                        gc.FindElement(By.XPath("//*[@id='EmailAddress']")).SendKeys("lightbot@lightsourcehr.com");
                        DateTime dateTime = DateTime.Now;
                        var date = dateTime.ToString("MM/dd/yyyy");
                        gc.FindElement(By.XPath("//*[@id='DueDate']")).SendKeys(date);

                        
                        gc.FindElement(By.XPath("//*[@id='Subject']")).SendKeys(record.caseNo.ToString());

                        gc.FindElement(By.XPath("//*[@id='btnSave']")).Click();
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Result : Case creation Sucessfull.");

                    }
                    catch (Exception ex1)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Result : Case creation failed..!");
                        Exception e = new Exception("Case Reporting failed !");
                        ErrorLogging(ex1);
                        ErrorLogging(e);
                    }
                }
            }
            if(Importable.Count>0)
            {
                //creating txt file for prism import
                string delimiter = "\t";
                try
                {
                    using (var writer = new System.IO.StreamWriter(basePath + "\\data.txt"))
                    {
                        foreach (Record i in Importable)
                        {
                            writer.WriteLine(i.billDate + delimiter + i.billEventCode + delimiter + i.billRates + delimiter + i.billUnits + delimiter + i.clientId +delimiter+ delimiter + i.location + delimiter + delimiter +i.comment);
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("Import file creation failed.");
                }

                //moving to prism to try importing the text file.
                top:
                gc.Navigate().GoToUrl("https://ctw.prismhr.com/ctw/dbweb.asp?dbcgm=1");
                System.Threading.Thread.Sleep(1000);
                //logging in prism
                try
                {
                    gc.FindElement(By.XPath("//*[@id='text4v1']")).SendKeys("lightbot");
                    gc.FindElement(By.XPath("//*[@id='password6v1']")).SendKeys("RPAuser1!");
                    gc.FindElement(By.XPath("//*[@id='button8v1']")).Click();
                    System.Threading.Thread.Sleep(1000);

                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Prism login failed!");
                    Console.ForegroundColor = ConsoleColor.White;
                    Exception e = new Exception("Prism Login failed !");
                    ErrorLogging(ex);
                    ErrorLogging(e);
                }

                try
                {
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).Click();
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("C");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("c");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("l");

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("i");

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("e");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("n");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("t");

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(" bill");
                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);
                    System.Threading.Thread.Sleep(1000);

                    gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Enter);
                    System.Threading.Thread.Sleep(10000);

                }
                catch (Exception ex)
                {

                    Exception e = new Exception("Client Bill Pending report opening failed !");
                    ErrorLogging(ex);
                    ErrorLogging(e);
                    goto end;
                }
                System.Threading.Thread.Sleep(1000);
                gc.FindElement(By.XPath("//*[@id='button31v2']")).Click();
                System.Threading.Thread.Sleep(1000);
                String current = gc.CurrentWindowHandle;
                foreach (String winHandle in gc.WindowHandles)
                {
                    gc.SwitchTo().Window(winHandle);
                }
                //sometimes the upload window doesn't open
                if (gc.CurrentWindowHandle != current)
                {
                    try
                    {
                        gc.FindElement(By.XPath("//*[@id='fname']")).SendKeys(basePath + "\\data.txt"); //put the full path of file here

                        gc.FindElement(By.XPath("//*[@id='submit1']")).Click();

                        System.Threading.Thread.Sleep(1000);
                        gc.FindElement(By.XPath("//*[@id='BUTTON1']")).Click();
                        System.Threading.Thread.Sleep(20000);
                        gc.SwitchTo().Window(current);
                    }
                    catch
                    {
                        try
                        {
                            gc.Close();
                        }
                        catch
                        {

                        }
                        try
                        {
                            gc.SwitchTo().Window(current);
                            gc.Close();
                        }
                        catch
                        {

                        }
                        goto end;
                    }
                    try
                    {
                        Exception s = new Exception(gc.FindElement(By.XPath("//*[@id='body_span29v2']")).Text);
                        ErrorLogging1(s);

                    }
                    catch
                    {

                    }

                    try
                    {
                        gc.FindElement(By.XPath("//*[@id='button33v2']")).Click();
                        System.Threading.Thread.Sleep(2000);
                        gc.FindElement(By.XPath("//*[@id='button32v2']")).Click();
                        System.Threading.Thread.Sleep(2000);
                        gc.SwitchTo().Alert().Accept();
                        gc.SwitchTo().Window(current);
                        gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        gc.SwitchTo().Window(current);
                        gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();

                    }
                    tryagain:
                    try
                    {

                        gc.Close();
                        gc.Dispose();
                        Console.WriteLine("Process Complete..!");

                    }
                    catch (Exception ex)
                    {
                        Exception e = new Exception("Chrome closing failed failed !");
                        ErrorLogging(ex);
                        ErrorLogging(e);
                        System.Threading.Thread.Sleep(10000);
                        goto tryagain;
                    }

                }
                else
                {
                    tryagain:
                    try
                    {
                        gc.Close();
                        gc.Dispose();
                        goto top;
                    }
                    catch (Exception ex)
                    {
                        Exception e = new Exception("Chrome closing failed failed !");
                        ErrorLogging(ex);
                        ErrorLogging(e);
                        System.Threading.Thread.Sleep(10000);
                        goto tryagain;
                    }
                }

            }//end of if condiditon
            end:
            Console.WriteLine("Process Completed..!");
            Console.WriteLine("System going to wait in 3 sec...");
            System.Threading.Thread.Sleep(3000);
            goto wait;
        }
    }
}
