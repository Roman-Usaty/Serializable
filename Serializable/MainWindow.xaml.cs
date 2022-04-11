using Microsoft.Win32;
using Serializable.Classes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Newtonsoft.Json;
using System.Xml.Serialization;

namespace Serializable
{
    public delegate void ErrHandler(int code, string msg);
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string? _pathToSelectFile;
        private string _formatSerialize;

        private List<SerializableElement> _elements;

        public event ErrHandler Error;
        public MainWindow()
        {
            InitializeComponent();
            _formatSerialize = "JSON";
            _elements = new List<SerializableElement>();

            Error += MainWindow_Error;
        }

        private void MainWindow_Error(int code, string msg)
        {
            _ = MessageBox.Show($"{code}: {msg}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            Close();
        }

        private void Serializable_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not ComboBox comboBox) return;

            if (comboBox.SelectedItem is not TextBlock item) return;

            _formatSerialize = item.Text;
        }

        private void ButtonSelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog? openFileDialog = OpenSaveFile.ShowFileOpen("*.txt", "Текстовый документ (*.txt) | *.txt | Все файлы (*.*) | *.*", "Открыть файл", false, 1);
            
            if (openFileDialog == null || openFileDialog.FileNames.Length <= 0) return;

            _pathToSelectFile = openFileDialog.FileName;

            using StreamReader stream = new StreamReader(_pathToSelectFile);
            string[] content = stream.ReadToEnd().Split('\n');

            foreach (string line in content)
            {
                string[] arrLine = line.Split(' ');
                try
                {
                    SerializableElement element = new SerializableElement()
                    {
                        Name = arrLine[0],
                        Age = int.Parse(arrLine[1]),
                        Group = arrLine[2]
                    };
                    _elements.Add(element);
                }
                catch (Exception ex)
                {
                    Error(1, $"Error from create serializable element: {ex.Message}");
                    continue;
                }
            }
        }

        private void ButtonSerialize_Click(object sender, RoutedEventArgs e)
        {
            switch (_formatSerialize)
            {
                case "JSON":
                    SerializatorType.JSON.Serialize(_elements, _pathToSelectFile ?? "default.txt");
                    break;
                case "XML":
                    SerializatorType.XML.Serialize(_elements, _pathToSelectFile ?? "default.txt");
                    break;
                case "EXCEL":
                    SerializatorType.EXCEL.Serialize(_elements, _pathToSelectFile ?? "default.txt");
                    break;
                default:
                    Error(2, "Неверный формат сериализации");
                    break;
            }
        }
    }
}
