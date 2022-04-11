using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace Serializable.Classes
{
    public enum SerializatorType
    {
        JSON,
        XML,
        EXCEL
    }
    public static class Serializator
    {
        public static int _i = 0;
        public static void Serialize(this SerializatorType type, List<SerializableElement> content, string path)
        {
            switch (type)
            {
                case SerializatorType.JSON:
                    SerializeJson(content, path);
                    break;
                case SerializatorType.XML:
                    SerializeXML(content, path);
                    break;
                case SerializatorType.EXCEL:
                    SerializeExcel(content, path);
                    break;
            }
        }

        private static void SerializeJson(List<SerializableElement> _elements, string path) 
        {
            string json = JsonConvert.SerializeObject(_elements, Formatting.Indented);
            string fileName = (Path.GetFileName(path) ?? $"default{_i++}.txt").Replace(".txt", "");

            using StreamWriter stream = new StreamWriter(Directory.GetCurrentDirectory() + $"\\{fileName}.json");
            stream.WriteLine(json);
            stream.Close();
            try
            {
                stream.Dispose();
            }
            catch (System.Exception)
            {
                return;
            }
        }

        private static void SerializeXML(List<SerializableElement> _elements, string path)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<SerializableElement>));

            string fileName = (Path.GetFileName(path) ?? $"default{_i++}.txt").Replace(".txt", "");

            using (FileStream stream = new FileStream(Directory.GetCurrentDirectory() + $"\\{fileName}.xml", FileMode.OpenOrCreate))
            {
                serializer.Serialize(stream, _elements);
                stream.Close();
                try
                {
                    stream.Dispose();
                }
                catch (System.Exception)
                {
                    throw;
                }
            }
        }

        private static void SerializeExcel(List<SerializableElement> _elements, string path)
        {
            Excel excel = new Excel();
            excel.SaveSucces += Excel_SaveSucces;
            excel.ErrHandle += Excel_ErrHandle;
            excel.AddRow(new string[] {"Name", "Age", "Group"});

            foreach (SerializableElement item in _elements)
            {
                excel.AddRow(new string[] { item.Name, item.Age.ToString(), item.Group });
            }

            string otherPath = Directory.GetCurrentDirectory() + $"\\{Path.GetFileName(path).Replace(".txt", "")}.xlsx";
            excel.FileSave(otherPath);
        }

        private static void Excel_ErrHandle(int code, string msg)
        {
            return;
        }

        private static void Excel_SaveSucces(string mess)
        {
            return;
        }
    }

}
