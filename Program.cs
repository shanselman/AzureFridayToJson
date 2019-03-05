using System;
using System.IO;
using ExcelDataReader;
using System.Text;
using Newtonsoft.Json;

namespace AzureFridayToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            var inFilePath = args[0];
            var outFilePath = args[1];

            using (var inFile = File.Open(inFilePath, FileMode.Open, FileAccess.Read))
            using (var outFile = File.CreateText(outFilePath))
            {
                using (var reader = ExcelReaderFactory.CreateReader(inFile, new ExcelReaderConfiguration()
                    { FallbackEncoding = Encoding.GetEncoding(1252) }))
                using (var writer = new JsonTextWriter(outFile))
                {
                    writer.Formatting = Formatting.Indented; //I likes it tidy
                    writer.WriteStartArray();
                    reader.Read(); //SKIP FIRST ROW, it's TITLES.
                    do
                    {
                        while (reader.Read())
                        {
                            //peek ahead? Bail before we start anything so we don't get an empty object
                            var status = reader.GetString(0);
                            if (string.IsNullOrEmpty(status)) break; 

                            writer.WriteStartObject();
                                writer.WritePropertyName("Status");
                                writer.WriteValue(status);

                                writer.WritePropertyName("Title");
                                writer.WriteValue(reader.GetString(1));

                                writer.WritePropertyName("Host");
                                writer.WriteValue(reader.GetString(6));

                                writer.WritePropertyName("Guest");
                                var str = reader.GetString(7);
                                writer.WriteValue(str);

                                writer.WritePropertyName("Episode");
                                writer.WriteValue(Convert.ToInt32(reader.GetDouble(2)));

                                writer.WritePropertyName("Live");
                                writer.WriteValue(reader.GetDateTime(5));

                                writer.WritePropertyName("Url");
                                writer.WriteValue(reader.GetString(11));

                                writer.WritePropertyName("EmbedUrl");
                                writer.WriteValue($"{reader.GetString(11)}player");
                                /*
                                <iframe src="https://channel9.msdn.com/Shows/Azure-Friday/Erich-Gamma-introduces-us-to-Visual-Studio-Online-integrated-with-the-Windows-Azure-Portal-Part-1/player" width="960" height="540" allowFullScreen frameBorder="0"></iframe>
                                 */

                            writer.WriteEndObject();
                        }
                    } while (reader.NextResult());
                    writer.WriteEndArray();
                }
            }
        }
    }
}
