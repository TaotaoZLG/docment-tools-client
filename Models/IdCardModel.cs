using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace docment_tools_client.Models
{
    public class IdCardModel
    {
        public string 姓名 { get; set; }
        public string 性别 { get; set; }
        public string 民族 { get; set; }
        public string 出生 { get; set; }
        public string 住址 { get; set; }
        public string 公民身份号码 { get; set; }
        public string 签发机关 { get; set; }
        public string 有效期限 { get; set; }

        // 转键值对
        public Dictionary<string, string> ToKeyValue()
        {
            return new Dictionary<string, string>
            {
                { nameof(姓名), 姓名 ?? "" },
                { nameof(性别), 性别 ?? "" },
                { nameof(民族), 民族 ?? "" },
                { nameof(出生), 出生 ?? "" },
                { nameof(住址), 住址 ?? "" },
                { nameof(公民身份号码), 公民身份号码 ?? "" },
                { nameof(签发机关), 签发机关 ?? "" },
                { nameof(有效期限), 有效期限 ?? "" }
            };
        }
    }
}
