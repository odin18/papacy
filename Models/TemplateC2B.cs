using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace papacy1.Models
{
    public class TemplateC2B
    {
        public string 單別單號 { get; set; }
        public string 日期 { get; set; }
        public string 記錄 { get; set; }
        public string 電腦編號 { get; set; }
        public string 批號 { get; set; }
        public string 紗別 { get; set; }
        public string 規格 { get; set; }
        public string 等級 { get; set; }
        public string 經緯別 { get; set; }
        public string 編號 { get; set; }
        public string 總箱數 { get; set; }
        public string 總管數 { get; set; }
        public string 總毛重 { get; set; }
        public string 總淨重 { get; set; }
        public List<Detail> 明細 { get; set; }

        public class Detail
        {
            public string No { get; set; }
            public string Num { get; set; }
            [JsonProperty("G.W.")]
            public string GW { get; set; }
            [JsonProperty("N.W.")]
            public string NW { get; set; }
        }
    }
}
