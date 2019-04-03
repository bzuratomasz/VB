using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test.Common.Model;

namespace Test.Common.Interfaces
{
    public interface IFileReader
    {
        FileReaderResponse ProcceedFiles(string path);
    }
}
