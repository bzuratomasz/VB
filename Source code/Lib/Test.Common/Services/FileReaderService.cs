using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test.Common.Interfaces;
using Test.Common.Model;
using Test.Common.Model.Internal;

namespace Test.Common.Services
{
    public class FileReaderService : IFileReader
    {
        public FileReaderResponse ProcceedFiles(string path)
        {
            var response = new FileReaderResponse();

            response.ProjectOrderNo = path.Substring(path.LastIndexOf("\\") + 1);

            response.GridSource = new DataTable();
            response.GridSource.Columns.Add("Conn.Location");
            response.GridSource.Columns.Add("Conn.Qty");

            response.FileModels = new List<FileReaderModel>();

            var bufferWireList = new List<FileReaderModel>();
            var comparer = new CustomObjectComparerService();

            try
            {
                DirectoryInfo directoryFiles = new DirectoryInfo(path);
                FileInfo[] Files = directoryFiles.GetFiles("*.txt");
                List<string> listConnLocation = new List<string>();

                foreach (FileInfo file in Files)
                {
                    using (var reader = new StreamReader(file.FullName))
                    {
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(';');

                            if (!values[1].Contains("Loc1"))
                            {
                                var leftSide = (values[1].Substring(values[1].Length - 1)).ToString();
                                var rightSide = (values[5].Substring(values[5].Length - 1)).ToString();
                                var leftRigtJoin = string.Format("{0}{1}", leftSide, rightSide);
                                listConnLocation.Add(leftRigtJoin);

                                var wireTypeName = values[10];
                                var wireColor = values[11];
                                var wireCrossSection = values[12];

                                bufferWireList.Add(new FileReaderModel()
                                {
                                    WireDef = new WireDefinition()
                                    {
                                        WireTypeName = wireTypeName,
                                        WireColor = wireColor,
                                        WireCrossSection = wireCrossSection,
                                        ConnectionLocation = leftRigtJoin
                                    }
                                });
                            }
                        }
                    }
                }

                var currBuff = listConnLocation
                      .GroupBy(l => l)
                      .Select(g => new
                      {
                          location = g.Key,
                          qty = listConnLocation.Count(s => s.Contains(g.Key))
                      })
                      .ToList();


                response.FileModels = bufferWireList
                      .GroupBy(l => l.WireDef, comparer)
                      .Select(g => new FileReaderModel()
                      {
                          WireDef = new WireDefinition()
                          {
                              WireColor = g.Key.WireColor,
                              WireCrossSection = g.Key.WireCrossSection,
                              WireTypeName = g.Key.WireTypeName,
                              ConnectionLocation = g.Key.ConnectionLocation
                          },
                          CountOfElements = bufferWireList.Count(s =>
                              s.WireDef.WireColor == g.Key.WireColor &&
                              s.WireDef.WireCrossSection == g.Key.WireCrossSection &&
                              s.WireDef.WireTypeName == g.Key.WireTypeName &&
                              s.WireDef.ConnectionLocation == g.Key.ConnectionLocation)
                      })
                      .ToList();

                currBuff.ForEach(item =>
                {
                    var _ravi = response.GridSource.NewRow();
                    _ravi["Conn.Location"] = item.location;
                    _ravi["Conn.Qty"] = item.qty;
                    response.GridSource.Rows.Add(_ravi);
                });
            }
            catch (Exception ex)
            {
                response.ExceptionCode = -1;
                response.ExceptionMessage = ex.ToString();
            }

            return response;
        }
    }
}
