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

namespace PDFreader
{
    public struct Paper
    {
        public string name;
        public Dictionary<string, int> words;
        public int len;
        public float mass;
        public Paper(string _name)
        {
            name = _name;
            words = null;
            mass = 0;
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
             transcrYandexMany("glossary.txt");

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
        string[] parseNameArtFromText(string text)
        {
            int indexRef = text.LastIndexOf("References");
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
        public int downloadOneArticleFromScihub(IWebDriver driver, string name)
        {
            driver.Navigate().GoToUrl("https://sci-hub.do");
            Thread.Sleep(2000);
            IWebElement element = driver.FindElement(By.XPath("//*[@type= 'textbox' ]"));
            element.SendKeys(name);
            element = driver.FindElement(By.XPath("//*[@id= 'open' ]"));
            element.Click();
            Thread.Sleep(2000);
            try
            {
                IWebElement element_ref = driver.FindElement(By.XPath("//*[@href= '#' ]"));
                element_ref.Click();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return 1;
        }
        public int downloadManyArticleFromScihub(string[] name)
        {
            IWebDriver driver = new ChromeDriver(@"C:\brouser");
            foreach (string name_art in name)
            {
                int error = 0;
                driver.Navigate().GoToUrl("https://sci-hub.do");
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
            var papers = loadFronJson();
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
        public string ReadPDF(string path)
        {
            using (PdfReader reader = new PdfReader(path))
            {
                string text = "";
                ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text += PdfTextExtractor.GetTextFromPage(reader, i, its);
                    if (i == reader.NumberOfPages)
                    {
                        text = PdfTextExtractor.GetTextFromPage(reader, i, its);
                    }
                }
                return text;
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
            saveToJson(listPaper);
            return listPaper;
        }
        void saveToJson(List<Paper> listPaper)
        {
            JsonSerializer serializer = new JsonSerializer();
            serializer.NullValueHandling = NullValueHandling.Ignore;
            serializer.Formatting = Formatting.Indented;

            using (StreamWriter sw = new StreamWriter("Papers.json"))
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                serializer.Serialize(writer, listPaper);
            }
        }
        Paper[] loadFronJson()
        {
            string jsontext = "";
            using (StreamReader file = File.OpenText("Papers.json"))
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

