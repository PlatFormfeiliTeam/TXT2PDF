using iTextSharp.text;
using iTextSharp.text.pdf;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TXT2PDF
{
    class Program
    {
        IDatabase db = SeRedis.redis.GetDatabase();
        string filedir = ConfigurationManager.AppSettings["filedir"];
        string rediskey = ConfigurationManager.AppSettings["rediskey"];  

        static void Main(string[] args)
        {
            if (ConfigurationManager.AppSettings["AutoRun"].ToString().Trim() == "Y")
            {
                // txt2pdf();
                fn_share fn_share = new fn_share(); 
                string filedic = System.Environment.CurrentDirectory + @"\log\";
                if (!Directory.Exists(filedic))
                {
                    Directory.CreateDirectory(filedic);
                }
                string filename = filedic + "txt2pdf_log_" + DateTime.Now.ToString("yyyyMMddHH") + ".txt";

                Program p = new Program();

                fn_share.systemLog(filename, "----------------------------------------------------------------------------------------------------\r\n\r\n");

                int count = Convert.ToInt32(ConfigurationManager.AppSettings["count"].ToString());
                for (int i = 0; i < count; i++)
                {     
                    p.txt2pdf(fn_share, filename);
                }
            }
        }
        private void txt2pdf(fn_share fn_share, string filename)
        {
            fn_share.systemLog(filename, "================ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " start================ \r\n");

            string json = string.Empty; string sql = string.Empty;
            try
            {
                json = db.ListLeftPop(rediskey);

                fn_share.systemLog(filename, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " json：" + json + "\r\n");

                if (!string.IsNullOrEmpty(json))
                {
                    JObject jo = (JObject)JsonConvert.DeserializeObject(json);//转JSON对象
                    if (File.Exists(filedir + jo.Value<string>("FILENAME")))
                    {
                        fn_share.systemLog(filename, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  文件路径：" + filedir + jo.Value<string>("FILENAME") + "\r\n");

                        int index = jo.Value<string>("FILENAME").LastIndexOf(".");
                        string preffix = jo.Value<string>("FILENAME").Substring(0, index + 1);
                        Document doc = new Document(PageSize.A4.Rotate());
                        PdfWriter.GetInstance(doc, new FileStream(filedir + preffix + "pdf", FileMode.Create));
                        doc.Open();
                        //中文字型問題REF http://renjin.blogspot.com/2009/01/using-chinese-fonts-in-itextsharp.html                       
                        string fontPath = Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\..\Fonts\kaiu.ttf";
                        //橫式中文
                        BaseFont bfChinese = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        iTextSharp.text.Font fontChinese = new iTextSharp.text.Font(bfChinese, 8f);
                        StreamReader sr = new StreamReader(filedir + jo.Value<string>("FILENAME"));
                        string line = null;
                        while ((line = sr.ReadLine()) != null)
                        {
                            doc.Add(new Paragraph(line, fontChinese));
                        }
                        doc.Close();

                        fn_share.systemLog(filename, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  文件转换结束\r\n");

                        sql = "update list_attachment set FILENAME='" + preffix + "pdf" + "' where  FILENAME='" + jo.Value<string>("FILENAME") + "' and ENTID='" + jo.Value<string>("ENTID") + "'";
                        DBMgr.ExecuteNonQuery(sql);

                        fn_share.systemLog(filename, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  ENTID：" + jo.Value<string>("ENTID") + "   更新sql结束 \r\n");
                    }
                }
            }
            catch (Exception ex)
            {
                db.ListRightPush(rediskey, json);
                fn_share.systemLog(filename, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  异常：" + ex.Message + "\r\n");
            }
            finally
            {
                fn_share.systemLog(filename, "================ " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " end  ================ \r\n");
            }
        }
    }
}
