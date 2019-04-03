using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test.Common.Model.Internal;

namespace Test.Common.Model
{
    public class FileReaderResponse
    {
        public string ProjectOrderNo { get; set; }
        public DataTable GridSource { get; set; }
        public List<FileReaderModel> FileModels { get; set; }
        public int ExceptionCode { get; set; }
        public string ExceptionMessage { get; set; }

        public override string ToString()
        {
            var result = FileModels.Aggregate("", (source, nextItem) => string.Format("{0}\n{1}", source, string.Format("{4}: {0},{2},{1} with count: {3}", nextItem.WireDef.WireTypeName, nextItem.WireDef.WireCrossSection, nextItem.WireDef.WireColor, nextItem.CountOfElements, nextItem.WireDef.ConnectionLocation)));
            return result;
        }
    }
}
