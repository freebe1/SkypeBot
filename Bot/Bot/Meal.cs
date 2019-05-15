using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Windows.Forms;
using HtmlAgilityPack;

namespace Bot
{
    class Meal
    {
        string result = null;
        string url = "http://m.welliv.co.kr/mobile/mealmenu_list.jsp";
        WebResponse response = null;
        StreamReader reader = null;
        String[] arr;
        private string add;

        public Meal() { }

        public String parse(string what)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";
                response = request.GetResponse();
                reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
                result = reader.ReadToEnd();
            }
            catch (Exception ex)
            {
                // handle error
                MessageBox.Show(ex.Message);
            }
            finally
            {
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(result);

                HtmlNode specificNode = doc.GetElementbyId("mainContent");
                int i = 0;
                var query = from table in doc.DocumentNode.SelectNodes("//table").Cast<HtmlNode>()
                            from row in table.SelectNodes("tr").Cast<HtmlNode>()
                            from cell in row.SelectNodes("th|td").Cast<HtmlNode>()
                            select new { Table = table.Id, CellText = cell.InnerText };
                arr = new String[8];
                foreach (var cell in query)
                {
                    if (i > 7) break;
                    arr.SetValue(cell.CellText, i);
                    i++;
                }
                i = 0;

                for (int j = 0; j < 7; j++)
                {
                    arr[j] = arr[j].Replace("&nbsp;", " ");
                }

                if (reader != null)
                    reader.Close();
                if (response != null)
                    response.Close();
            }
            if (what=="AM") return arr[5];
            if (what == "NOON") return arr[6];
            if (what == "PM") return arr[7];
            return "";
        }
        
    }
}
