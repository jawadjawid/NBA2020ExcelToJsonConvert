using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelToJSON
{
    class ExcelFile
    {
        private string path;

        private FileInfo fileInfo;

        private ExcelPackage package;

        public ExcelFile(string path)
        {
            this.path = path;
            fileInfo = new FileInfo(path);
            package = new ExcelPackage(fileInfo);
        }

        public ExcelPackage GetPackage()
        {
            return package;
        }


    }
}
