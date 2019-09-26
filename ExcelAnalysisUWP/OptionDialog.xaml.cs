using System;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;


namespace ExcelAnalysisUWP
{
    public sealed partial class OptionDialog : ContentDialog
    {
        public ExcutionMode Mode { get; private set; }

        public ExcutionMethod ExcutionMethod { get; private set; }

        public StorageFile InputFile { get; private set; }

        public bool IsCoverOriginFile { get; private set; } = true;

        public OptionDialog()
        {
            InitializeComponent();

            ModeCombo.Items.Add("计算数据和绘制直线");
            ModeCombo.Items.Add("仅计算数据");
            ModeCombo.SelectedIndex = 0;
            MethodCombo.Items.Add("退一");
            MethodCombo.Items.Add("错位");
            MethodCombo.Items.Add("退一-1");
            MethodCombo.Items.Add("错位-1");
            MethodCombo.Items.Add("错位-20");
            MethodCombo.SelectedIndex = 0;
        }

        private void ModeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (ModeCombo.SelectedItem.ToString())
            {
                case "仅计算数据":
                    Mode = ExcutionMode.ComputeDataOnly;
                    break;
                case "计算数据和绘制直线":
                    Mode = ExcutionMode.ComputeDataAndLine;
                    break;
            }
        }

        private void MethodCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (MethodCombo.SelectedItem.ToString())
            {
                case "退一":
                    ExcutionMethod = ExcutionMethod.Primary;
                    break;
                case "错位":
                    ExcutionMethod = ExcutionMethod.Secondary;
                    break;
                case "退一-1":
                    ExcutionMethod = ExcutionMethod.Third;
                    break;
                case "错位-1":
                    ExcutionMethod = ExcutionMethod.Forth;
                    break;
                case "错位-20":
                    ExcutionMethod = ExcutionMethod.Sixth;
                    break;
            }
        }
    }
}
