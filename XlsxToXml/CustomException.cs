using System;
using System.Collections.Generic;
using System.Text;

namespace XlsxToXml
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
