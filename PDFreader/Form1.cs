using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Xml.XPath;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;
using System.Text.RegularExpressions;
using Keys = OpenQA.Selenium.Keys;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using OpenQA.Selenium.Support.UI;

namespace PDFreader
{
    public struct Paper
    {
        public string name;
        public string doi;
        public Dictionary<string, int> words;
        public int len;//load
        public float mass;
        public Paper(string _name)
        {
            name = _name;
            doi = "";
            words = null;
            mass = 0;
            len = 0;
        }
        public Paper(string _name,string _doi, int _load )
        {
            name = _name;
            doi = _doi;
            words = null;
            mass = _load;
            len = 0;
        }
        public void addMass(float _mass)
        {
            mass += _mass;
        }
    }
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            var json_path = "Papers_Stereo.json";

            //doiFromResearchGateMany(json_path);
            loadSciHubMany(json_path);
        }
        void doiFromResearchGateMany(string json_path)
        {
            var papers = loadFronJson(json_path);
            int len = papers.Length;
            int batch_size = 100;
            for(int i =1084; i<len; i+= batch_size)
            {
                ChromeOptions opt = new ChromeOptions();
                opt.PageLoadStrategy = PageLoadStrategy.None;
                var driver = new ChromeDriver(@"C:\brouser", opt);
                doiFromResearchGateMany(driver, json_path,i,batch_size);
                driver.Close();
                driver.Dispose();
                Thread.Sleep(30000);
            }
            
        }

        void loadSciHubMany(string json_path)
        {
            var papers = loadFronJson(json_path);
            int len = papers.Length;
            int batch_size = 10;
            for (int i = 39; i < len; i += batch_size)
            {
                downloadOneArticleFromScihubWait( json_path, i, batch_size);
                Thread.Sleep(30000);
            }

        }
        void PapersFromSciHubManyJson(string json_path)
        {
            var papers = loadFronJson(json_path);
            var driver = new ChromeDriver(@"C:\brouser");
            for (int i = 0; i < papers.Length; i++)
            {
                var paper = papers[i];
                var doi = paper.doi;
                if(doi.Length>2)
                {
                    Console.WriteLine(doi[0]+" "+doi);
                    if(doi[0]=='1')
                    {
                        var ret = downloadOneArticleFromScihub(driver, doi);
                        papers[i] = new Paper(paper.name,paper.doi,ret);
                        saveToJson(papers.ToList(), json_path);
                    }
                   
                }
                            
            }
            driver.Close();
            driver.Dispose();
        }
        void doiFromResearchGateMany(IWebDriver driver,string json_path,int start_ind, int batch_size)
        {
            var papers = loadFronJson(json_path);
            //loginResearchGate(driver);

            for (int i = 0; i < batch_size; i++)
            {
                var ind = start_ind + i;
                var paper = papers[ind];                
                var doi = "";
                try
                {
                    doi = learnDoiArticleFromGoogleResearchGate(driver, paper.name);
                }
                catch
                {

                }
                var doi_paper = new Paper(paper.name);
                doi_paper.doi = doi;
                papers[ind] = doi_paper;
                saveToJson(papers.ToList(), json_path);         
            }
        }
        #region translate_with_transcription
        void transcrYandexMany(string path)
        {
            string file1;
            using (StreamReader sr = new StreamReader(path))
            {
                file1 = sr.ReadToEnd();
            }
            string[] lines = file1.Split(new char[] { '\n' });
            List<string> transc = new List<string>();
            List<string> transl = new List<string>();
            List<string> words = new List<string>();
            int ind = 0;
           for (int i=0;i<20;i++)
            {
               var tr = transcrYandexOne(lines[i]);
                if(tr!=null)
                {
                    transc.Add(tr[0]);
                    transl.Add(tr[1]);
                    words.Add(lines[i]);
                }
                Console.WriteLine(ind);
                ind++;
            }
            Console.WriteLine(transl.Count);
            using (StreamWriter sw = new StreamWriter("transc.txt", false, System.Text.Encoding.UTF8))
            {
                foreach(string word in transc)
                {
                    sw.WriteLine(word);
                }
            }
            using (StreamWriter sw = new StreamWriter("transl.txt", false, System.Text.Encoding.UTF8))
            {
                foreach (string word in transl)
                {
                    sw.WriteLine(word);
                }
            }
            using (StreamWriter sw = new StreamWriter("words.txt", false, System.Text.Encoding.UTF8))
            {
                foreach (string word in words)
                {
                    sw.Write(word);
                }
            }
        }

        string[] transcrYandexOne(string text)
        {
            var translateUrl = "https://dictionary.yandex.net/dicservice.json/lookup?ui=ru&text="
                                + text + "&lang=en-ru&flags=23";
            using (var wc = new WebClient())
            {
                wc.Encoding = Encoding.UTF8;
                var resultHtml = wc.DownloadString(translateUrl);
                dynamic trsJson = JObject.Parse(resultHtml);
                try
                {
                    var trsl = trsJson.def[0].tr[0].text;
                    var trscr = trsJson.def[0].ts;
                    return new string[2] { trsl, trscr };
                    
                }
                catch
                {

                }
            }
            return null;
        }
        string[] translOneWord(string word)
        {
            IWebDriver driver = new ChromeDriver(@"C:\brouser");
            driver.Navigate().GoToUrl("https://myefe.ru/anglijskaya-transkriptsiya.html");
            IWebElement element = driver.FindElement(By.XPath("//*[@id= 'mlsw7-main-search-input' ]"));
            element.SendKeys(word);
            //element = driver.FindElement(By.XPath("//*[@class= 'ui basic primary disabled button mlsw7-search-action' ]"));
            element = driver.FindElement(By.XPath("//*[@type= 'submit' ]"));
            element.Click();
            Thread.Sleep(2000);
            element = driver.FindElement(By.XPath("//*[@class= 'ml-item' ]"));
            string trans = element.Text;
            foreach(sbyte a in trans)
            {
                Console.WriteLine(a);
            }
            return new string[2]{ "",""};
        }
        #endregion
        #region articles
        bool findBegDelim(char symb)
        {
            if(symb == '(')
            {
                return true;
            }
            return false;
        }
        bool findEndDelim(char symb)
        {
            if (symb == ')')
            {
                return true;
            }
            return false;
        }
        bool conditionDelim(string text)
        {
            int age = 0;
            var ret = Int32.TryParse(text, out age);
            if(ret)
            {
                if(age>1920 && age < 2030)
                {
                    return true;
                }
            }
            return false ;
        }
       
        string[] parseNameArtFromText_withoutKvSk(string text)
        {
            List<string> art_names = new List<string>();
            var lenDelim = 5;
            var ind = 0;
            for(int i=0; i<text.Length; i++)
            {
                if(findBegDelim(text[i]))
                {
                    if (findEndDelim(text[i + lenDelim]) && conditionDelim(text.Substring(i+1,4)))
                    {
                        //Console.WriteLine("Find :"  + text.Substring(i, 40));
                        if(text.Length>250)
                        {
                            var _name = text.Substring(i + 8, 240);
                            var name = _name.Split(new char[] { '.' })[0];
                            art_names.Add(name);
                        }
                        
                        //Console.WriteLine(name);
                        //Console.WriteLine("__________________--");
                        
                    }
                    ind++;
                }
            }
            Console.WriteLine("ART LEN: " + art_names.Count+ " ind: " + ind+ " text.Length: " + text.Length);
            return art_names.ToArray();
        }
        string[] parseNameArtFromText(string text)
        {
            int indexRef = text.LastIndexOf("References");
            Console.WriteLine(indexRef);
            text = text.Remove(0, indexRef);
            string[] references = text.Split(new char[] { '\n' });
            List<string> ref_parse = new List<string>();
            string subtext = "";
            for (int i = 0; i < references.Length - 1; i++)
            {
                const char ch = '[';
                if (references[i].Length > 3)
                {
                    subtext += references[i];
                }
                if (references[i + 1][0].CompareTo(ch) == 0)
                {
                    ref_parse.Add(subtext);
                    subtext = "";
                }
            }

            //-------------------------------------------
            List<string> art_names = new List<string>();
            foreach (string str in ref_parse)
            {
                Console.WriteLine(str);
                string[] ref_spl = str.Split(new char[] { ',' });
                List<string> list_pice = new List<string>();
                for (int i = 1; i < ref_spl.Length; i++)
                {
                    if (ref_spl[i].Length > 24)
                    {
                        list_pice.Add(ref_spl[i]);
                    }
                }
                if (list_pice.Count != 0)
                {
                    art_names.Add(list_pice[0]);
                }
            }
            int i1 = 0;
            foreach (string str in art_names)
            {
                i1++;
                Console.WriteLine(i1.ToString() + "" + str);
            }
            return art_names.ToArray();
        }
        string[] extractArticleNamePdf(string path)
        {
            string[] files = Directory.GetFiles(path);
            int i = 0;
            foreach (string filename in files)
            {
                Console.WriteLine(filename);
                string fileName = System.IO.Path.GetFileName(filename);

                fileName = fileName.Trim();
                fileName = fileName.Substring(0, fileName.Length - 4);


                files[i] = fileName;
                i++;
            }
            return files;
        }
        public void loginResearchGate(IWebDriver driver)
        {           
            driver.Navigate().GoToUrl("https://www.researchgate.net/login?");
            Thread.Sleep(2000);
            string email = "";
            string password = "";
            IWebElement element = driver.FindElement(By.XPath("//*[@id='input-login']"));
            element.SendKeys(email);
            element = driver.FindElement(By.XPath("//*[@id='input-password']"));
            element.SendKeys(password);
            element = driver.FindElement(By.XPath("//*[@data-testid = 'loginCta']"));
            element.Click();
            
        }
        public string learnDoiArticleFromResearchGate(string artName)
        {
            IWebDriver driver = new ChromeDriver(@"C:\brouser");
            string doi = learnDoiArticleFromResearchGate(driver, artName);
            return doi;
        }
        public string learnDoiArticleFromGoogleResearchGate(IWebDriver driver, string artName)
        {
            string doi = "";
            driver.Navigate().GoToUrl("https://www.google.com/");
            Thread.Sleep(700);
            IWebElement element = driver.FindElement(By.XPath("//*[@class='gLFyf gsfi']"));
            element.SendKeys(artName+" researchgate");
            element.SendKeys(Keys.Return);
            Thread.Sleep(1000);
            element = driver.FindElement(By.XPath("//*[@id='search']//div[1]//div[1]//div[1]//div[1]//div[1]//div[1]//a[1]")); 
            driver.Navigate().GoToUrl(element.GetProperty("href"));
            Thread.Sleep(3500);
            element = driver.FindElement(By.XPath("//*[@class='research-detail-header-section__metadata']//div[2]//a[1]"));

            doi = element.Text;

            //Console.WriteLine(doi);
            return doi;
        }
            public string learnDoiArticleFromResearchGate(IWebDriver driver, string artName)
        {
            string doi = "";
            driver.Navigate().GoToUrl("https://www.researchgate.net/");
            Thread.Sleep(100);
            IWebElement element = driver.FindElement(By.XPath("//*[@class='header-search__bar-button']"));
            element.Click();
            Thread.Sleep(500);
            element = driver.FindElement(By.XPath("//*[@class='search-container__form-input']"));
            element.SendKeys(artName);
            Thread.Sleep(500);
            element.SendKeys(Keys.Return);//list papers
            Thread.Sleep(1500);
            element = driver.FindElement(By.XPath("//*[@class='nova-legacy-e-text nova-legacy-e-text--size-l nova-legacy-e-text--family-sans-serif nova-legacy-e-text--spacing-none nova-legacy-e-text--color-inherit nova-legacy-e-text--clamp-3 nova-legacy-v-publication-item__title']//a[1]"));
            element.Click();
            Thread.Sleep(500);
            element = driver.FindElement(By.XPath("//*[@class='nova-legacy-e-link nova-legacy-e-link--color-inherit nova-legacy-e-link--theme-decorated']"));
            doi = element.Text;
            return doi;
        }
        public string learnDoiArticleFromPubmed(IWebDriver driver, string artName)
        {
            string doi = "";
            driver.Navigate().GoToUrl("https://pubmed.ncbi.nlm.nih.gov");
            IWebElement element = driver.FindElement(By.XPath("//*[@name= 'term' ]"));
            element.SendKeys(artName);
            element = driver.FindElement(By.XPath("//*[@class= 'search-btn' ]"));
            element.Click();
            try
            {
                IWebElement element_ref = driver.FindElement(By.XPath("//*[@class='citation-doi' ]"));
                doi = element_ref.Text;
            }
            catch (Exception e)
            {
                try
                {
                    IWebElement element_ref = driver.FindElement(By.XPath("//*[@class = 'docsum-pmid']"));
                    string pmid = element_ref.Text;
                    driver.Navigate().GoToUrl("https://pubmed.ncbi.nlm.nih.gov");
                    element = driver.FindElement(By.XPath("//*[@name= 'term' ]"));
                    element.SendKeys(pmid);
                    element = driver.FindElement(By.XPath("//*[@class= 'search-btn' ]"));
                    element.Click();
                    element_ref = driver.FindElement(By.XPath("//*[@class='citation-doi' ]"));
                    doi = element_ref.Text;
                }
                catch
                {

                }

            }
            return doi;
        }
        public string learnDoiArticleFromPubmed(string artName)
        {
            IWebDriver driver = new ChromeDriver(@"C:\brouser");
            string doi = learnDoiArticleFromPubmed(driver, artName);
            return doi;
        }






        public void downloadOneArticleFromScihubWait(string json_path,int star_ind,int bath_size)
        {
            var papers = loadFronJson(json_path);
            ChromeOptions opt = new ChromeOptions();

            opt.PageLoadStrategy = PageLoadStrategy.None;
            var driver = new ChromeDriver(@"C:\brouser", opt);
            string url = "https://sci-hub.ru";
            for (int i = 0; i < bath_size; i++)
            {
                var open = false;
                newTabcea(driver, url);
                while (open==false)
                {
                    //newTabcea(driver, url);
                    open = checkAllTabcea(driver);
                    if(open==false)
                    {
                        newTabcea(driver, url);
                    }
                    Thread.Sleep(2000);
                }
                int ind = i + star_ind;
                var paper = papers[ind];
                var doi = paper.doi;
                if (doi.Length > 2)
                {
                    Console.WriteLine(ind + " " + doi);
                    if (doi[0] == '1')
                    {
                        Console.WriteLine(driver.PageSource.Length);
                        openArticle(driver, doi);
                        Thread.Sleep(5000);
                        //var ret = clickLoad(driver, doi);
                        manualDownload();
                        Thread.Sleep(1000);
                        driver.Close();
                        //closeLastPage(driver);
                        papers[ind] = new Paper(paper.name, paper.doi, 0);
                        saveToJson(papers.ToList(), json_path);
                        
                    }

                }

            }
           // driver.Close();
            driver.Dispose();
        }
        void manualDownload()
        {
            Left_Click(new Point(833, 161));
            Thread.Sleep(1000);
            Left_Click(new Point(805, 516));
            //SendKeys.SendWait("{Enter}");
        }
        void newTabcea(IWebDriver driver,string url)
        {
            if(driver.WindowHandles.Count<50)
            {
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                ((IJavaScriptExecutor)driver).ExecuteScript("window.open();");
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                driver.Navigate().GoToUrl(url);
            }
            
        }
        bool checkAllTabcea(IWebDriver driver)
        {
            var pages = driver.WindowHandles;

            bool check_page = false;
            for (int i=0; i<pages.Count;i++)
            {
                driver.SwitchTo().Window(pages[i]);
                check_page = checkTabcea(driver);
                if (check_page)
                {
                    return check_page;
                }
            }            
            return check_page;
        }
        bool checkTabcea(IWebDriver driver)
        {

            //var check_page = ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete");
            return driver.PageSource.Length>30000 && driver.PageSource.Length < 31000;
        }
        int clickLoad(IWebDriver driver, string name)
        {
            try
            {
                IWebElement element_ref = driver.FindElement(By.XPath("//*[@id='buttons']//button[1]"));
                element_ref.Click();
                return 1;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return 0;
        }
        void openArticle(IWebDriver driver, string name)
        {
            IWebElement element = driver.FindElement(By.XPath("//*[@type= 'textbox' ]"));
            element.SendKeys(name);
            element = driver.FindElement(By.XPath("//*[@id= 'open' ]"));
            element.Click();
        }
        public int downloadOneArticleFromScihub(IWebDriver driver, string name)
        {
            int error = 0;
            driver.Navigate().GoToUrl("https://sci-hub.ru");
            IWebElement element = driver.FindElement(By.XPath("//*[@type= 'textbox' ]"));
            element.SendKeys(name);
            element = driver.FindElement(By.XPath("//*[@id= 'open' ]"));
            element.Click();
            try
            {
                IWebElement element_ref = driver.FindElement(By.XPath("//*[@id='buttons']//button[1]"));
                element_ref.Click();
            }
            catch (Exception e)
            {
                error = 1;
            }         
            return 1;
        }
        public int downloadManyArticleFromScihub(string[] name)
        {
            IWebDriver driver = new ChromeDriver(@"C:\brouser");
            foreach (string name_art in name)
            {
                int error = 0;
                driver.Navigate().GoToUrl("https://sci-hub.ru");
                IWebElement element = driver.FindElement(By.XPath("//*[@type= 'textbox' ]"));
                element.SendKeys(name_art);
                element = driver.FindElement(By.XPath("//*[@id= 'open' ]"));
                element.Click();
                try
                {
                    IWebElement element_ref = driver.FindElement(By.XPath("//*[@href= '#' ]"));
                    element_ref.Click();
                }
                catch (Exception e)
                {
                    error = 1;
                }
                if (error == 1)
                {
                    string doi = learnDoiArticleFromPubmed(driver, name_art);
                    Console.WriteLine(doi);
                    downloadOneArticleFromScihub(driver, doi);

                }

            }
            return 1;
        }
        #endregion

        #region patent
        private void button1_Click(object sender, EventArgs e)
        {
            string path_art = @"C:\Users\Dell\Desktop\patents";
            string[] keyw = { "scan", "laser", "сканир", "лазер" };
            // var papers = scanManyPdf(@"C:\Users\Dell\Desktop\учёба\асп\Кандидатская\лазерные сканеры\патенты", keyw);
            var papers = loadFronJson("Papers.json");
            for (int i = 0; i < papers.Length; i++)
            {
                if (papers[i].len != 0)
                {
                    foreach (var key in keyw)
                    {
                        papers[i].mass += (float)papers[i].words[key] / (float)papers[i].len;
                    }
                }

            }


            var sortedpapers = sortPapers(papers);
            string folder = "sort";
            for (int i = sortedpapers.Length - 1; i > sortedpapers.Length - 100; i--)
            {

                Console.WriteLine(sortedpapers[i].name + ": " + sortedpapers[i].mass);
                File.Copy(System.IO.Path.Combine(path_art, sortedpapers[i].name), System.IO.Path.Combine(path_art, folder, (sortedpapers.Length + 1 - i).ToString() + "_" + sortedpapers[i].name));
            }
            Console.WriteLine("END");
        }
        public int downloadOnePatentFromGoogle(string name)
        {
            IWebDriver driver = new ChromeDriver(@"C:\brouser");
            return downloadOnePatentFromGoogle(driver, name);
        }

        public void closeLastPage(IWebDriver driver)
        {
            var pages = driver.WindowHandles;
            var page = pages[pages.Count - 1];
            driver.SwitchTo().Window(page);
            driver.Close();
            driver.SwitchTo().Window(pages[0]);
        }
        public void clickDownload(IWebDriver driver)
        {
            var element_download = driver.FindElement(By.XPath("//*[@target= '_blank' ]"));
            element_download.Click();
            Thread.Sleep(1500);
            Left_Click(new Point(1810, 135));
            Thread.Sleep(1500);
            Left_Click(new Point(1302, 770));
            closeLastPage(driver);
        }
        public void searchOn(IWebDriver driver,string name)
        {
            driver.Navigate().GoToUrl("https://patents.google.com");
            
            Thread.Sleep(500);
            IWebElement element = driver.FindElement(By.XPath("//*[@id= 'searchInput' ]"));
            element.SendKeys(name);
            element = driver.FindElement(By.XPath("//*[@id='searchButton' ]"));
            element.Click();
            Thread.Sleep(1000);
        }
        public string[] findRefs(IWebDriver driver)
        {
            var element_ref = driver.FindElements(By.XPath("//*[@class= 'style-scope state-modifier' ]"));
            List<string> names = new List<string>();
            for (int i = 0; i < element_ref.Count; i++)
            {
                var text = element_ref[i].Text;

                if (text.Length > 3)
                {

                    if (text.Contains("RU") || text.Contains("EP") || text.Contains("US"))
                    {
                        if (text.Contains("Publication") || text.Contains("Priority"))
                        {
                            var strText = text.Split();
                            text = strText[2];
                            names.Add(text);
                            Console.WriteLine(text);
                        }
                        else
                        {
                            names.Add(text);
                            Console.WriteLine(text);
                        }
                    }
                }
            }
            return names.ToArray();
        }
        public int downloadOnePatentFromGoogle(IWebDriver driver,string name)
        {
            try
            {
                searchOn(driver,name);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            try
            {
                clickDownload(driver);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return 1;
        }

        public void downloadManyPatentFromGoogle(string name, int deapth)
        {
            IWebDriver driver = new ChromeDriver(@"C:\brouser");
            Left_Click(new Point(880, 25));
            driver.Navigate().GoToUrl("https://patents.google.com");
            downloadManyPatentFromGoogle(driver, name, deapth);
        }
        public void downloadManyPatentFromGoogle(IWebDriver driver,string name,int deapth)
        {
            driver.Navigate().GoToUrl("https://patents.google.com");
            deapth--;
            try
            {
                searchOn(driver, name);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            try
            {
                var names = findRefs(driver);
                try
                {
                    clickDownload(driver);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                foreach (string _name in names)
                {
                    try
                    {
                        if(deapth>0)
                        {
                            downloadManyPatentFromGoogle(driver, _name, deapth);
                        }
                        else
                        {
                            downloadOnePatentFromGoogle(driver, _name);
                        }
                        
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void deleteSameNames(string path)
        {
            string[] files = Directory.GetFiles(path);
            foreach (string filename in files)
            {
                string fileName = System.IO.Path.GetFileName(filename);
                if(fileName.Contains("("))
                {
                    var file = new FileInfo(filename);
                    file.Delete();
                }
            }
        }
        #endregion

        #region pdf
        public string ReadPDF(string path,int startPage = 0)
        {
            using (PdfReader reader = new PdfReader(path))
            {
                string text = "";
                StringBuilder stringBuilder = new StringBuilder();
                ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
                for (int i = startPage; i <= reader.NumberOfPages; i++)
                {                                      
                   PdfTextExtractor.GetTextFromPage(reader, i, its);
                    if (i == reader.NumberOfPages)
                    {
                        stringBuilder.Append(PdfTextExtractor.GetTextFromPage(reader, i, its));
                    }
                }
                return stringBuilder.ToString();
            }
        }
        public Paper[] sortPapers(Paper[] papers,string key)
        {
            var sortedPapers = from u in papers
                              orderby u.words[key]
                              select u;
            return sortedPapers.ToArray();
        }
        public Paper[] sortPapers(Paper[] papers)
        {
            var sortedPapers = from u in papers
                               orderby u.mass
                               select u;
            return sortedPapers.ToArray();
        }

        public List<Paper> scanManyPdf(string path, string[] keywords)
        {
            var listPaper = new List<Paper>();
            string[] files = Directory.GetFiles(path);
            foreach (string filename in files)
            {
                var text = ReadPDF(filename);
                string fileName = System.IO.Path.GetFileName(filename);
                var paper = new Paper(fileName);
              
                paper.words = new Dictionary<string, int>();
                paper.len = text.Length;
                foreach (string word in keywords)
                {
                    var reg = new Regex(@"" + word + @"(\w*)", RegexOptions.IgnoreCase);
                    var match = reg.Matches(text);
                    paper.words.Add(word, match.Count);
                }
                listPaper.Add(paper);

            }
            saveToJson(listPaper, "Papers.json");
            return listPaper;
        }
        void saveToJson(List<Paper> listPaper,string path)
        {
            JsonSerializer serializer = new JsonSerializer();
            serializer.NullValueHandling = NullValueHandling.Ignore;
            serializer.Formatting = Formatting.Indented;

            using (StreamWriter sw = new StreamWriter(path))
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                serializer.Serialize(writer, listPaper);
            }
        }
        Paper[] loadFronJson(string path)
        {
            string jsontext = "";
            using (StreamReader file = File.OpenText(path))
            {
                jsontext = file.ReadToEnd();
            }
            return JsonConvert.DeserializeObject<Paper[]>(jsontext);
        }
        string[] extractImagesFromPdf(string path_art, string path_img)
        {
            string[] files = Directory.GetFiles(path_art);
            int i = 0;
            foreach (string filename in files)
            {
                Console.WriteLine(filename);
                string namePdf = i.ToString();
                while (namePdf.Length < 6)
                {
                    namePdf = "0" + namePdf;
                }
                ExtractImagesFromPDF(filename, path_img, namePdf);

                i++;
            }
            return files;
        }
        #endregion


        #region ExtractImagesFromPDF
        public static void ExtractImagesFromPDF(string sourcePdf, string outputPath, string nameImg)
        {
            // NOTE:  This will only get the first image it finds per page.
            PdfReader pdf = new PdfReader(sourcePdf);
            //RandomAccessFileOrArray raf = new iTextSharp.text.pdf.RandomAccessFileOrArray(sourcePdf);

            try
            {
                for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++)
                {
                    PdfDictionary pg = pdf.GetPageN(pageNumber);
                    PdfDictionary res = (PdfDictionary)PdfReader.GetPdfObject(pg.Get(PdfName.RESOURCES));

                    PdfDictionary xobj = (PdfDictionary)PdfReader.GetPdfObject(res.Get(PdfName.XOBJECT));
                    if (xobj != null)
                    {
                        foreach (PdfName name in xobj.Keys)
                        {
                            PdfObject obj = xobj.Get(name);
                            if (obj.IsIndirect())
                            {
                                PdfDictionary tg = (PdfDictionary)PdfReader.GetPdfObject(obj);
                                PdfName type = (PdfName)PdfReader.GetPdfObject(tg.Get(PdfName.SUBTYPE));
                                if (PdfName.IMAGE.Equals(type))
                                {
                                    try
                                    {
                                        int XrefIndex = Convert.ToInt32(((PRIndirectReference)obj).Number.ToString(System.Globalization.CultureInfo.InvariantCulture));
                                        PdfObject pdfObj = pdf.GetPdfObject(XrefIndex);
                                        PdfStream pdfStrem = (PdfStream)pdfObj;
                                        byte[] bytes = PdfReader.GetStreamBytesRaw((PRStream)pdfStrem);
                                        if ((bytes != null))
                                        {
                                            using (System.IO.MemoryStream memStream = new System.IO.MemoryStream(bytes))
                                            {
                                                memStream.Position = 0;
                                                System.Drawing.Image img = System.Drawing.Image.FromStream(memStream);
                                                // must save the file while stream is open.
                                                if (!Directory.Exists(outputPath))
                                                    Directory.CreateDirectory(outputPath);

                                                string path = System.IO.Path.Combine(outputPath, nameImg + " " + pageNumber.ToString() + ".jpg");
                                                //Console.WriteLine(path);
                                                EncoderParameters parms = new EncoderParameters(1);
                                                parms.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, 0);
                                                // GetImageEncoder is found below this method
                                                ImageCodecInfo jpegEncoder = GetImageEncoder("JPEG");
                                                img.Save(path, jpegEncoder, parms);
                                                parms.Param[0].Dispose();
                                                parms.Dispose();
                                                jpegEncoder = null;
                                                img.Dispose();
                                                break;


                                            }
                                        }
                                    }
                                    catch
                                    {

                                    }
                                }
                                tg = null;
                                type = null;
                            }
                        }
                    }
                }
            }

            catch (Exception e)
            {
                // Console.WriteLine(e.Message);
            }
            finally
            {
                pdf.Close();
                pdf.Dispose();
            }


        }

        public static System.Drawing.Imaging.ImageCodecInfo GetImageEncoder(string imageType)
        {
            imageType = imageType.ToUpperInvariant();
            foreach (ImageCodecInfo info in ImageCodecInfo.GetImageEncoders())
            {
                if (info.FormatDescription == imageType)
                {
                    return info;
                }
            }

            return null;
        }
        #endregion

        #region mouse
        [DllImport("User32.dll")]
        static extern void mouse_event(MouseFlags dwFlags, int dx, int dy, int dwData, UIntPtr dwExtraInfo);
        [Flags]
        enum MouseFlags
        {
            Move = 0x0001, LeftDown = 0x0002, LeftUp = 0x0004, RightDown = 0x0008,
            RightUp = 0x0010, Absolute = 0x8000
        };
        
        private void Left_Click(Point pos)
        {
            int x = (int)pos.X * 65535 / 1920;
            int y = (int)pos.Y * 65535 / 1080;
            mouse_event(MouseFlags.Absolute | MouseFlags.Move, x, y, 0, UIntPtr.Zero);
            mouse_event(MouseFlags.Absolute | MouseFlags.LeftDown, x, y, 0, UIntPtr.Zero);
            mouse_event(MouseFlags.Absolute | MouseFlags.LeftUp, x, y, 0, UIntPtr.Zero);

        }
        #endregion
    }
}

