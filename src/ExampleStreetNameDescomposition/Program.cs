using Infodinamica.Framework.Exportable.Engines;
using Infodinamica.Framework.Exportable.Engines.Excel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExampleStreetNameDescomposition
{
    class Program
    {
        private static readonly Regex _regexStreetNumber = new Regex(@"(\d+)");
        private static readonly Regex _regexFinalStreetNumber = new Regex(@"\b(número|numero|nro|n|#)\z", RegexOptions.IgnoreCase);
        private static readonly Regex _regexFinalStreetTokens = new Regex(@"\b(número|numero|nro|n|#)\z", RegexOptions.IgnoreCase);
        private static readonly Regex _regexHighRoadStreet = new Regex(@"\b(avenida |av |av. )", RegexOptions.IgnoreCase);
        private static readonly Regex _regexHomeZoneStreet = new Regex(@"\b(calle |pasaje |psj |psj. |población |poblacion |villa )", RegexOptions.IgnoreCase);

        static void Main(string[] args)
        {
            var addressesInfo = new ConcurrentBag<AddressInfo>();

            foreach (var address in AddressData.Streets)
            //Parallel.ForEach(AddressData.Streets, address =>
            //foreach (var address in new List<string>() { "Calle 1 4372 Villa Gildemaister" })
            {
                AddressInfo addressInfo;
                var matches = _regexStreetNumber.Matches(address);
                if (matches.Count > 0)
                {
                    //Take match for High Roads and Home Zone
                    var highRoadStreet = _regexHighRoadStreet.Match(address);
                    var homeZoneStreet = _regexHomeZoneStreet.Match(address);
                    var streetName = GetStreetName(address, matches, highRoadStreet, homeZoneStreet);

                    //If match is null, then the address don't have number
                    if (!string.IsNullOrWhiteSpace(streetName))
                    {
                        addressInfo = new AddressInfo
                        {
                            OriginalAddress = address,
                            StreetType = GetStreetType(highRoadStreet, homeZoneStreet),
                            StreetName = streetName,
                            Number = GetStreetNumber(address, matches, highRoadStreet, homeZoneStreet)
                        };
                    }
                    else
                        addressInfo = new AddressInfo
                        {
                            OriginalAddress = address,
                            StreetType = string.Empty,
                            StreetName = address,
                            Number = null
                        };
                }
                else
                {
                    addressInfo = new AddressInfo
                    {
                        OriginalAddress = address,
                        StreetType = string.Empty,
                        StreetName = address,
                        Number = null
                    };
                }

                //Add address info to array
                addressesInfo.Add(addressInfo);
            //});
            }

            CreateFile(addressesInfo.OrderBy(x => x.OriginalAddress).ToList());

            Console.WriteLine("Proceso finalizado. Presione cualquier tecla para finalizar");
            Console.ReadKey();
        }

        private static string GetStreetName(string street, MatchCollection matches, Match highRoadStreet, Match homeZoneStreet)
        {
            var streetName = street;

            if (matches == null || matches.Count == 0)
                return string.Empty;

            var matchNumber = matches[0];
            if (!matchNumber.Success)
                return string.Empty;

            if(matches.Count > 1 && (matchNumber.Index <= highRoadStreet.Index + highRoadStreet.Length + 1 || matchNumber.Index <= homeZoneStreet.Index + homeZoneStreet.Length + 1))
                matchNumber = matches[1];

            if (matchNumber == null || !matchNumber.Success)
                return string.Empty;

            streetName = street.Substring(0, matchNumber.Index).Trim();
            var match = _regexFinalStreetNumber.Match(streetName);
            
            if (match.Success)            
                streetName = street.Substring(0, match.Index).Trim();
            match = _regexFinalStreetTokens.Match(streetName);
            if (match.Success)            
                streetName = street.Substring(0, match.Index).Trim();
                
            if (highRoadStreet.Success && highRoadStreet.Index==0 && streetName.Length > highRoadStreet.Length + 2)
                streetName = streetName.Substring(highRoadStreet.Index + highRoadStreet.Length).Trim();

            if (homeZoneStreet.Success && homeZoneStreet.Index == 0 && streetName.Length > homeZoneStreet.Length + 2)
                streetName = streetName.Substring(homeZoneStreet.Index + homeZoneStreet.Length).Trim();

            return streetName;
        }

        private static string GetStreetType(Match highRoadStreet, Match homeZoneStreet)
        {   
            if (highRoadStreet.Success)
                return "Avenida";            
            if (homeZoneStreet.Success)
                return "Calle";

            return string.Empty;
        }

        private static int GetStreetNumber(string street, MatchCollection matches, Match highRoadStreet, Match homeZoneStreet)
        {
            if (matches == null || matches.Count == 0)
                return 0;

            var matchNumber = matches[0];
            if (!matchNumber.Success)
                return 0;

            if (matches.Count > 1 && (matchNumber.Index <= highRoadStreet.Index + highRoadStreet.Length + 1 || matchNumber.Index <= homeZoneStreet.Index + homeZoneStreet.Length + 1))
                matchNumber = matches[1];

            if (matchNumber == null || !matchNumber.Success)
                return 0;

            return Int32.Parse(matchNumber.Value);
        }

        private static void CreateFile(IList<AddressInfo> addresses)
        {
            IExportEngine engine = new ExcelExportEngine();
            engine.AsExcel().AddData<AddressInfo>(addresses, "sheet1");
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var folderPath = string.Format(@"{0}\output\", path);

            if (!Directory.Exists(folderPath))
                Directory.CreateDirectory(folderPath);

            engine.Export(folderPath + Guid.NewGuid().ToString() + ".xlsx");
        }
    }
}
