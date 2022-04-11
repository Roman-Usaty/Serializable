

using Microsoft.Win32;
using System.IO;
using System.Windows;

namespace Serializable.Classes
{
    public static class OpenSaveFile
    {
        static int _i = 0;
        public static OpenFileDialog? ShowFileOpen(string defExt, string filter, string title, bool multiSelect, int countFile)
        {
            OpenFileDialog openFile = new();
            openFile.DefaultExt = defExt;
            openFile.Multiselect = multiSelect;
            openFile.Filter = filter;
            openFile.Title = title;
#pragma warning disable CS8629 // Тип значения, допускающего NULL, может быть NULL.
            if (!(bool)openFile.ShowDialog())
            {
                return null;
            }
#pragma warning restore CS8629 // Тип значения, допускающего NULL, может быть NULL.

            if (openFile.FileNames.Length <= (countFile - 1) || openFile.FileNames.Length > countFile)
            {
                MessageBox.Show($"Выберите {countFile} файл(а)", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return null;
            }
            return openFile;
        }

        public static int SaveFile(string path, string content, string ext)
        {
            string fileName = (Directory.GetFiles(path).ToString() ?? $"default{_i++}.txt").Replace(".txt", "");
            
            using (StreamWriter stream = new StreamWriter(Directory.GetCurrentDirectory() + $"\\{fileName}{ext}"))
            {
                stream.WriteLine(content);
                stream.Close();
                try
                {
                    stream.Dispose();
                }
                catch (System.Exception)
                {
                    return 1;
                }
            }

            return 0;
        }
    }
}
