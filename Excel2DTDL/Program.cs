using Azure.DigitalTwins.Core;
using Azure.Identity;
using Microsoft.Azure.DigitalTwins.Parser;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Excel2DTDL
{
    class Program
    {
        //private static DigitalTwinsClient client;

        static void Main(string[] args)
        {
            //Uri adtInstanceUrl;
            //try
            //{
            //    // Read configuration data from the 
            //    IConfiguration config = new ConfigurationBuilder()
            //        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
            //        .Build();
            //    adtInstanceUrl = new Uri(config["instanceUrl"]);
            //}
            //catch (Exception ex) when (ex is FileNotFoundException || ex is UriFormatException)
            //{
            //    Log.Error($"Could not read configuration. Have you configured your ADT instance URL in appsettings.json?\n\nException message: {ex.Message}");
            //    return;
            //}

            //Log.Ok("Authenticating...");
            //var credential = new DefaultAzureCredential();
            //client = new DigitalTwinsClient(adtInstanceUrl, credential);

            //Log.Ok($"Service client created – ready to go");
            Log.Ok($"Usage: Excel2DTDL <excelfilename.xls>");

            if (args.Length != 1 || !args[0].EndsWith(".xlsx"))
            {
                Log.Error("Please input excel filename.");
                Environment.Exit(0);
            }

            var filename = args[0];

            var excelParser = new ExcelParser();
            var dtdl = excelParser.Parse(filename);

            string jsonDtdl = JsonConvert.SerializeObject(dtdl.InterfaceArray,
                                        Formatting.Indented,
                                        new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });

            // Validate DTDL
            Log.Ok("Validating ... ");
            ModelParser parser = new ModelParser();
            parser.DtmiResolver = Resolver;

            try
            {
                IReadOnlyDictionary<Dtmi, DTEntityInfo> om = parser.ParseAsync(new List<string>(1) { jsonDtdl }).GetAwaiter().GetResult();

                Log.Ok("DTDL is valid.");
            }
            catch (ParsingException pe)
            {
                Log.Error($"*** Error parsing models");
                int derrcount = 1;
                foreach (ParsingError err in pe.Errors)
                {
                    Log.Error($"Error {derrcount}:");
                    Log.Error($"{err.Message}");
                    Log.Error($"Primary ID: {err.PrimaryID}");
                    Log.Error($"Secondary ID: {err.SecondaryID}");
                    Log.Error($"Property: {err.Property}\n");
                    derrcount++;
                }
                Environment.Exit(0);
            }
            catch (ResolutionException)
            {
                Log.Error("Could not resolve required references");
                Environment.Exit(0);
            }

            File.WriteAllText(@"dtdl.json", jsonDtdl, System.Text.Encoding.UTF8);
            Log.Ok("DTDL is saved to 'dtdl.json' file.");
        }

        static async Task<IEnumerable<string>> Resolver(IReadOnlyCollection<Dtmi> dtmis)
        {
            Log.Error($"*** Error parsing models. Missing:");
            foreach (Dtmi d in dtmis)
            {
                Log.Error($"  {d}");
            }
            return null;
        }
    }
}
