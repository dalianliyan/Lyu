using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web;
using System.Net;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
namespace MarvellousWorks.PracticalPattern.Common.Tests
{
    [TestClass]
    public class WebRequestFixture
    {
        const string HomePage = "http://www.cnblogs.com/callwangxiang/";
        const string Pattern = @"((http|https|ftp):(\/\/|\\\\)((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)";
        const string Mark = "ExerciseAA";
        const int IntervalSeconds = 10;

        List<string> urls;
        int executeTimes;
        Timer timer;

        [TestInitialize]
        public void Initialize()
        {
            urls = new List<string>();
            var page = GetWebPage(HomePage);
            var coll = Regex.Matches(page, Pattern);
            foreach (var item in coll)
            {
                if (item.ToString().Contains(Mark))
                    urls.Add(item.ToString());
            }
            urls = urls.Distinct().ToList();
            timer = new Timer(InvokePage, null, IntervalSeconds * 1000, IntervalSeconds * 1000);
        }

        public string GetWebPage(string url)
        {
            var request = (HttpWebRequest)WebRequest.Create(url);
            Trace.WriteLine(url);
            request.Method = "GET";
            using (var reader = new StreamReader(((HttpWebResponse)request.GetResponse()).GetResponseStream()))
            {
                return reader.ReadToEnd();
            }            
        }

        void InvokePage(object sender)
        {
            urls.AsParallel().ForAll(x=>GetWebPage(x));
            executeTimes++;
            if (executeTimes == 20)
            {
                executeTimes = 0;
                Thread.Sleep(120000);
                timer = null;
            }

        }

        [TestMethod]
        public void TestReloadPage()
        {
            urls.ForEach(x=>Trace.WriteLine(x));
            
            while (true)
            {
                if(timer == null)
                    timer = new Timer(InvokePage, null, IntervalSeconds * 1000, IntervalSeconds * 1000);
            }
        }
    }
}
