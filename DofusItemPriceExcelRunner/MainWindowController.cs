using DofusItemPriceExcelPj;
using Microsoft.Win32;
using System;
using System.IO;

namespace DofusItemPriceExcelRunner
{
    internal class MainWindowController
    {
        private readonly string AppdataDirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\DofusItemPrice";
        private readonly string AppdataFileName = "lastPath.txt";
        private bool runBtnEnabled;

        public ProgramRunner Runner { get; set; } = new ProgramRunner();
        public string FilePath { get; set; }
        public bool RunBtnEnabled
        {
            get => runBtnEnabled;
            set
            {
                runBtnEnabled = value;
                OnRunBtnValueChanged?.Invoke(value);
            }
        }
        public Action<bool> OnRunBtnValueChanged { get; set; }
        public Action OnWorkDone { get; set; }

        private string AppdataFilePath => AppdataDirectoryPath + AppdataFileName;

        public MainWindowController(Action<bool> onRunBtnValueChanged)
        {
            OnRunBtnValueChanged = onRunBtnValueChanged;
            if(File.Exists(AppdataFilePath))
            {
                FilePath = File.ReadAllText(AppdataFilePath);
                RunBtnEnabled = true;
            }
        }

        public void OnSelectBtnClicked()
        {
            string path;
            if(File.Exists(AppdataFilePath))
            {
                path = File.ReadAllText(AppdataFilePath);
            }
            else
            {
                path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }

            var fileDialog = new OpenFileDialog
            {
                Title = "Sélection du fichier excel contenant les prix",
                Filter = "Fichier excel|*xlsx",
                DefaultExt = ".xlsx",
                Multiselect = false,
                InitialDirectory = path.Substring(0, path.LastIndexOf("\\")),
                FileName = path.Substring(path.LastIndexOf("\\") + 1)
            };
            if(fileDialog.ShowDialog().Value)
            {
                FilePath = fileDialog.FileName;
                SaveChosenFilepath();
                RunBtnEnabled = true;
            }
        }

        private void SaveChosenFilepath()
        {
            Directory.CreateDirectory(AppdataDirectoryPath);
            File.WriteAllText(AppdataFilePath, FilePath);
        }

        internal void OnRunButtonClicked()
        {
            Runner.Run(FilePath);
            OnWorkDone?.Invoke();
        }
    }
}
