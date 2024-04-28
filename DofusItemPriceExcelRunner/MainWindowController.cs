using DofusItemPriceExcel.Objects;
using DofusItemPriceExcelPj;
using Microsoft.Win32;
using Newtonsoft.Json;
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
        public bool SeuilBuySellChecked { get; set; }
        public RunOptions RunOptions { get; set; }
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
        public string BuySellThreshold { get; set; } = "10";

        private string AppdataFilePath => AppdataDirectoryPath + AppdataFileName;

        public MainWindowController(Action<bool> onRunBtnValueChanged)
        {
            OnRunBtnValueChanged = onRunBtnValueChanged;
            if(File.Exists(AppdataFilePath))
            {
                try
                {
                    RunOptions = JsonConvert.DeserializeObject<RunOptions>(File.ReadAllText(AppdataFilePath));
                    BuySellThreshold = $"{RunOptions.BuySellThresholdPercent}";
                    SeuilBuySellChecked = RunOptions.BuySellThresholdPercent != 0;
                }
                catch(JsonReaderException e)
                {
                    RunOptions = new RunOptions();
                    if(e.Message.StartsWith("Unexpected character encountered while parsing value:"))
                    {
                        RunOptions.FilePath = File.ReadAllText(AppdataFilePath);
                    }
                }
                RunBtnEnabled = true;
            }
            else
            {
                RunOptions = new RunOptions();
            }
        }

        public void OnSelectBtnClicked()
        {
            string path;

            if(File.Exists(AppdataFilePath))
            {
                RunOptions options;
                try
                {
                    options = JsonConvert.DeserializeObject<RunOptions>(File.ReadAllText(AppdataFilePath));
                }
                catch(JsonReaderException e)
                {
                    options = new RunOptions();
                    if(e.Message.StartsWith("Unexpected character encountered while parsing value:"))
                    {
                        options.FilePath = File.ReadAllText(AppdataFilePath);
                    }
                }
                path = options.FilePath;
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
                RunOptions.FilePath = fileDialog.FileName;
                RunBtnEnabled = true;
            }
        }

        private void SaveChosenOptions()
        {
            Directory.CreateDirectory(AppdataDirectoryPath);
            File.WriteAllText(AppdataFilePath, JsonConvert.SerializeObject(RunOptions));
        }

        internal void OnRunButtonClicked()
        {
            if(SeuilBuySellChecked)
            {
                if(int.TryParse(BuySellThreshold, out int threshold))
                {
                    RunOptions.BuySellThresholdPercent = threshold;
                }
            }
            else
            {
                RunOptions.BuySellThresholdPercent = 0;
            }
            SaveChosenOptions();
            Runner.Run(RunOptions);
            OnWorkDone?.Invoke();
        }
    }
}
