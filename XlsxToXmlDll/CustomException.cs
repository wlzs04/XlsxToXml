using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsxToXmlDll
{
    class CustomException : Exception
    {
        public string customMessage = "未添加报错信息!";

        public CustomException(string customMessage)
        {
            this.customMessage = customMessage;
        }
    }
}
