using AnimationEffectProvider;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using Windows.Foundation;
using Windows.Storage;
using Windows.Storage.Streams;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

namespace ExcelAnalysisUWP
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private ExcelPackage Package;
        private ExcelWorkbook Workbook { get => Package?.Workbook; }
        private ExcelWorksheets Worksheets { get => Workbook?.Worksheets; }
        private ExcelWorksheet CurrentSheet { get; set; }
        private ExcutionMode Mode;
        private const int DataStartRow = 70;
        private const int DataStartColumn = 3;
        private const int GoupDistance = 6;
        private ExcutionMethod ExcutionMethod;
        private Action<int, int> ProcessDelegate;
        private int TotalDataLength;
        private IProgress<object> Pro;
        private int Tick;
        private StorageFile File;
        private EntranceAnimationEffect EntranceEffectProvider;
        private bool IsCoverOriginFile;

        public MainPage()
        {
            InitializeComponent();
            Window.Current.SetTitleBar(TitleBar);
            Loaded += MainPage_Loaded;
        }

        private void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            EntranceEffectProvider.StartEntranceEffect();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            if (e.Parameter is Rect SplashRect)
            {
                EntranceEffectProvider = new EntranceAnimationEffect(this, Gr, SplashRect);
                EntranceEffectProvider.PrepareEntranceEffect();
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void DrawLineCore(ExcelRange UpCell, ExcelRange DownCell)
        {
            System.Diagnostics.Debug.WriteLine(UpCell.Start.Row + "/" + UpCell.Start.Column);
            lock (this)
            {
                ExcelShape LineShape = CurrentSheet.Drawings.AddShape(Guid.NewGuid().ToString(), eShapeStyle.Line);
                if (UpCell.Start.Column < DownCell.Start.Column)
                {
                    LineShape.From.Row = UpCell.Start.Row;
                    LineShape.From.Column = UpCell.Start.Column;
                    LineShape.From.RowOff = Convert.ToInt32(CurrentSheet.Row(UpCell.Start.Row).Height);
                    LineShape.From.ColumnOff = Convert.ToInt32(CurrentSheet.Column(UpCell.Start.Column).Width / 2);

                    LineShape.To.Row = DownCell.Start.Row;
                    LineShape.To.Column = DownCell.Start.Column;
                    LineShape.To.RowOff = Convert.ToInt32(CurrentSheet.Row(DownCell.Start.Row).Height);
                    LineShape.To.ColumnOff = Convert.ToInt32(CurrentSheet.Column(DownCell.Start.Column).Width / 2);
                }
                else
                {
                    LineShape.To.Row = UpCell.Start.Row;
                    LineShape.To.Column = UpCell.Start.Column;
                    LineShape.To.RowOff = Convert.ToInt32(CurrentSheet.Row(UpCell.Start.Row).Height);
                    LineShape.To.ColumnOff = Convert.ToInt32(CurrentSheet.Column(UpCell.Start.Column).Width / 2);

                    LineShape.From.Row = DownCell.Start.Row;
                    LineShape.From.Column = DownCell.Start.Column;
                    LineShape.From.RowOff = Convert.ToInt32(CurrentSheet.Row(DownCell.Start.Row).Height);
                    LineShape.From.ColumnOff = Convert.ToInt32(CurrentSheet.Column(DownCell.Start.Column).Width / 2);
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void DrawLineTask(int CurrentRow, int DataLength)
        {
            for (int i = DataStartColumn; i < DataLength + DataStartColumn; i++)
            {
                if (CurrentSheet.Cells[CurrentRow, i].Text == CurrentSheet.Cells[CurrentRow + 1, i].Text)
                {
                    ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, i];
                    ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, i - 1];
                    if (ZeroObject.Text == "0")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
                else
                {
                    ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, i];
                    ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, i - 1];
                    if (OneObject.Text == "1")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void GroupProcessorPrimaryMethod(int CurrentRow, int DataLength)
        {
            ExcelRange LeftMove = CurrentSheet.Cells[CurrentRow, DataStartColumn + 1, CurrentRow, DataStartColumn + DataLength];
            LeftMove.Copy(CurrentSheet.Cells[CurrentRow + 1, DataStartColumn]);

            for (int i = DataStartColumn; i < DataLength + DataStartColumn; i++)
            {
                if (CurrentSheet.Cells[CurrentRow, i].Text == CurrentSheet.Cells[CurrentRow + 1, i].Text)
                {
                    CurrentSheet.Cells[CurrentRow + 2, i].Value = "1";
                }
                else
                {
                    CurrentSheet.Cells[CurrentRow + 4, i].Value = "0";
                }
            }

            if (Mode == ExcutionMode.ComputeDataAndLine)
            {
                DrawLineTask(CurrentRow, DataLength);

                Pro.Report(null);
            }

            ExcelRange Range = CurrentSheet.Cells[CurrentRow + 4, DataStartColumn, CurrentRow + 4, DataLength + 2];
            Range.Copy(CurrentSheet.Cells[CurrentRow + GoupDistance, DataStartColumn]);

            for (int i = DataStartColumn; i < DataLength + DataStartColumn; i++)
            {
                if (CurrentSheet.Cells[CurrentRow + 4, i].Text != "0")
                {
                    CurrentSheet.Cells[CurrentRow + GoupDistance, i].Value = "1";
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void GroupProcessorSecondaryMethod(int CurrentRow, int DataLength)
        {
            int Row = CurrentRow, Length = DataLength;
            for (int i = DataStartColumn; i < Length + DataStartColumn; i++)
            {
                if (CurrentSheet.Cells[Row, i].Text == CurrentSheet.Cells[Row + 1, i].Text)
                {
                    CurrentSheet.Cells[Row + 2, i].Value = "1";
                }
                else
                {
                    CurrentSheet.Cells[Row + 4, i].Value = "0";
                }
            }

            if (DataLength != 1)
            {
                ExcelRange Range = CurrentSheet.Cells[CurrentRow, DataStartColumn, CurrentRow, DataStartColumn + TotalDataLength];
                Range.Copy(CurrentSheet.Cells[CurrentRow + 6, DataStartColumn]);

                ExcelRange Left = CurrentSheet.Cells[CurrentRow + 1, DataStartColumn + 1, CurrentRow + 1, DataStartColumn + DataLength];
                Left.Copy(CurrentSheet.Cells[CurrentRow + 7, DataStartColumn]);
            }

            if (Mode == ExcutionMode.ComputeDataAndLine)
            {
                DrawLineTask(Row, Length);
            }
            Pro.Report(null);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void GroupProcessorThirdMethod(int CurrentRow, int DataLength)
        {
            ExcelRange AddObject = CurrentSheet.Cells[CurrentRow, DataStartColumn + DataLength];

            int CompareIndex = DataStartColumn + DataLength - 1;
            CurrentSheet.Cells[CurrentRow + 1, CompareIndex].Value = AddObject.Text;

            if (CurrentSheet.Cells[CurrentRow, CompareIndex].Text == CurrentSheet.Cells[CurrentRow + 1, CompareIndex].Text)
            {
                ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, CompareIndex];
                OneObject.Value = "1";
                CurrentSheet.Cells[CurrentRow + 6, CompareIndex].Value = (OneObject.Text);

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, CompareIndex - 1];
                    if (ZeroObject.Text == "0")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }
            else
            {
                ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, CompareIndex];
                ZeroObject.Value = "0";
                CurrentSheet.Cells[CurrentRow + 6, CompareIndex].Value = (ZeroObject);

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, CompareIndex - 1];
                    if (OneObject.Text == "1")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }

            Pro.Report(null);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void GroupProcessorForthMethod(int CurrentRow, int DataLength)
        {
            int CompareIndex = DataStartColumn + DataLength - 1;

            if (CurrentSheet.Cells[CurrentRow, CompareIndex].Text == CurrentSheet.Cells[CurrentRow + 1, CompareIndex].Text)
            {
                ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, CompareIndex];
                OneObject.Value = "1";

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, CompareIndex - 1];
                    if (ZeroObject.Text == "0")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }
            else
            {
                ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, CompareIndex];
                ZeroObject.Value = "0";

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, CompareIndex - 1];
                    if (OneObject.Text == "1")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }

            Pro.Report(null);

            if (DataLength != 1)
            {
                ExcelRange Range = CurrentSheet.Cells[DataStartRow, DataStartColumn, DataStartRow, DataStartColumn + TotalDataLength];
                Range.Copy(CurrentSheet.Cells[CurrentRow + 6, DataStartColumn]);
                ExcelRange Origin = CurrentSheet.Cells[CurrentRow, DataStartColumn + TotalDataLength];
                CurrentSheet.Cells[CurrentRow + 7, CompareIndex - 1].Value = Origin.Text;
            }
        }

        private void GroupProcessorSixthMethod(int CurrentRow, int DataLength)
        {
            //复制错位行最后三角形区域的数
            if (DataLength != 1)
            {
                //这里加个删除前面不要的数据的方法就好了
                if ((DataStartColumn + DataLength - 21) < DataStartColumn)
                {
                    ExcelRange Range = CurrentSheet.Cells[CurrentRow + 1, DataStartColumn + 1, CurrentRow + 1, DataStartColumn + DataLength];
                    Range.Copy(CurrentSheet.Cells[CurrentRow + 7, DataStartColumn]);
                }
            }

            //操作一行
            if ((DataStartColumn + DataLength - 20) >= DataStartColumn)
            {
                for (int i = DataStartColumn + DataLength - 20; i < DataLength + DataStartColumn; i++)
                {
                    if (CurrentSheet.Cells[CurrentRow, i].Text == CurrentSheet.Cells[CurrentRow + 1, i].Text)
                    {
                        CurrentSheet.Cells[CurrentRow + 2, i].Value = "1";
                    }
                    else
                    {
                        CurrentSheet.Cells[CurrentRow + 4, i].Value = "0";
                    }
                }
            }
            else
            {
                for (int i = DataStartColumn; i < DataLength + DataStartColumn; i++)
                {
                    if (CurrentSheet.Cells[CurrentRow, i].Text == CurrentSheet.Cells[CurrentRow + 1, i].Text)
                    {
                        CurrentSheet.Cells[CurrentRow + 2, i].Value = "1";
                    }
                    else
                    {
                        CurrentSheet.Cells[CurrentRow + 4, i].Value = "0";
                    }
                }
            }

            //画线操作
            if (Mode == ExcutionMode.ComputeDataAndLine)
            {
                if ((DataStartColumn + DataLength - 20) >= DataStartColumn)
                {
                    for (int i = DataStartColumn + DataLength - 20; i < DataLength + DataStartColumn; i++)
                    {
                        if (CurrentSheet.Cells[CurrentRow, i].Text == CurrentSheet.Cells[CurrentRow + 1, i].Text)
                        {
                            ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, i];
                            ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, i - 1];
                            if (ZeroObject.Text == "0")
                            {
                                DrawLineCore(OneObject, ZeroObject);
                            }
                        }
                        else
                        {
                            ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, i];
                            ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, i - 1];
                            if (OneObject.Text == "1")
                            {
                                DrawLineCore(OneObject, ZeroObject);
                            }
                        }
                    }
                }
                else
                {
                    for (int i = DataStartColumn; i < DataLength + DataStartColumn; i++)
                    {
                        if (CurrentSheet.Cells[CurrentRow, i].Text == CurrentSheet.Cells[CurrentRow + 1, i].Text)
                        {
                            ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, i];
                            ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, i - 1];
                            if (ZeroObject.Text == "0")
                            {
                                DrawLineCore(OneObject, ZeroObject);
                            }
                        }
                        else
                        {
                            ExcelRange ZeroObject = CurrentSheet.Cells[CurrentRow + 4, i];
                            ExcelRange OneObject = CurrentSheet.Cells[CurrentRow + 2, i - 1];
                            if (OneObject.Text == "1")
                            {
                                DrawLineCore(OneObject, ZeroObject);
                            }
                        }
                    }
                }
            }

            Pro.Report(null);
        }

        private async Task Start_Process()
        {
            if (Stack.Visibility == Visibility.Collapsed)
            {
                Stack.Visibility = Visibility.Visible;
                Drag.Visibility = Visibility.Collapsed;
            }

            Progress.IsIndeterminate = true;

            await Task.Delay(3000);

            Pro = new Progress<object>(async (o) =>
            {
                double CurrentText = 100 - (Tick-- - 1) * (100f / TotalDataLength);
                await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, () =>
                 {
                     if (Progress.IsIndeterminate == true)
                     {
                         Progress.IsIndeterminate = false;
                     }

                     Progress.Value = CurrentText;

                     double CeilingText = Math.Ceiling(CurrentText);

                     if (CeilingText > 100)
                     {
                         return;
                     }
                 });
            });

            await Task.Run(() =>
            {
                try
                {
                    int ColumnsCount = 0;

                    for (int i = DataStartColumn; ; i++)
                    {
                        if (CurrentSheet.Cells[DataStartRow, i].Text == "")
                        {
                            ColumnsCount = i - 4;
                            break;
                        }
                    }
                    TotalDataLength = ColumnsCount < 1000 ? ColumnsCount : 1000;
                    Tick = TotalDataLength;

                    switch (ExcutionMethod)
                    {
                        case ExcutionMethod.Primary:
                            {
                                ExcelRange Range = CurrentSheet.Cells[DataStartRow, DataStartColumn, DataStartRow, DataStartColumn + ColumnsCount];
                                Range.Copy(CurrentSheet.Cells[DataStartRow + 2, DataStartColumn]);

                                ProcessDelegate = GroupProcessorPrimaryMethod;
                                break;
                            }
                        case ExcutionMethod.Secondary:
                            {
                                ExcelRange Range = CurrentSheet.Cells[DataStartRow, DataStartColumn, DataStartRow, DataStartColumn + ColumnsCount];
                                Range.Copy(CurrentSheet.Cells[DataStartRow + 2, DataStartColumn]);

                                ExcelRange LeftMove = CurrentSheet.Cells[DataStartRow + 2, DataStartColumn + 1, DataStartRow + 2, DataStartColumn + ColumnsCount];
                                LeftMove.Copy(CurrentSheet.Cells[DataStartRow + 3, DataStartColumn]);

                                ProcessDelegate = GroupProcessorSecondaryMethod;
                                break;
                            }
                        case ExcutionMethod.Third:
                            {
                                ExcelRange AddObject = CurrentSheet.Cells[DataStartRow, DataStartColumn + ColumnsCount];
                                CurrentSheet.Cells[DataStartRow + 2, DataStartColumn + ColumnsCount].Value = AddObject.Text;

                                ProcessDelegate = GroupProcessorThirdMethod;
                                break;
                            }
                        case ExcutionMethod.Forth:
                            {
                                ExcelRange AddObject = CurrentSheet.Cells[DataStartRow, DataStartColumn + ColumnsCount];
                                CurrentSheet.Cells[DataStartRow + 2, DataStartColumn + ColumnsCount].Value = AddObject.Text;

                                CurrentSheet.Cells[DataStartRow + 3, DataStartColumn + ColumnsCount - 1].Value = AddObject.Text;

                                ProcessDelegate = GroupProcessorForthMethod;
                                break;
                            }
                        case ExcutionMethod.Sixth:
                            {
                                //直接生成250个模块的原始数据
                                ExcelRange Origin = CurrentSheet.Cells[DataStartRow, DataStartColumn, DataStartRow, DataStartColumn + ColumnsCount];
                                for (int i = 0; i < 250; i++)
                                {
                                    Origin.Copy(CurrentSheet.Cells[DataStartRow + 2 + 6 * i, DataStartColumn]);
                                }

                                ExcelRange Left = CurrentSheet.Cells[DataStartRow, DataStartColumn + ColumnsCount - 19, DataStartRow, DataStartColumn + ColumnsCount];
                                for (int i = 0; i < 250; i++)
                                {
                                    Left.Copy(CurrentSheet.Cells[DataStartRow + 3 + i * 6, DataStartColumn + ColumnsCount - 20 - i]);
                                    //复制错位行能完整复制时候的数
                                    if ((DataStartColumn + ColumnsCount - 20 - i) <= DataStartColumn)
                                    {
                                        break;
                                    }
                                }

                                ExcelRange LeftMove = CurrentSheet.Cells[DataStartRow + 2, DataStartColumn + 1 + ColumnsCount - 20, DataStartRow + 2, DataStartColumn + ColumnsCount];
                                LeftMove.Copy(CurrentSheet.Cells[DataStartRow + 3, DataStartColumn + ColumnsCount - 20]);

                                ProcessDelegate = GroupProcessorSixthMethod;
                                break;
                            }
                    }

                    for (int CurrentRow = DataStartRow + 2, Count = 0; ColumnsCount > 0 && Count < 250; CurrentRow += GoupDistance, ColumnsCount--, Count++)
                    {
                        ProcessDelegate(CurrentRow, ColumnsCount);
                    }

                    StorageFile TempFile = ApplicationData.Current.TemporaryFolder.CreateFileAsync("ResultTemp.xlsx", CreationCollisionOption.ReplaceExisting).AsTask().Result;
                    using (var Stream = TempFile.OpenAsync(FileAccessMode.ReadWrite).AsTask().Result.AsStream())
                    {
                        Package.SaveAs(Stream);
                    }

                    if (IsCoverOriginFile)
                    {
                        TempFile.MoveAndReplaceAsync(File).AsTask().Wait();
                    }
                    else
                    {
                        string NewName = File.DisplayName + "-已处理" + File.FileType;
                        TempFile.CopyAsync(File.GetParentAsync().AsTask().Result, NewName, NameCollisionOption.GenerateUniqueName).AsTask().Wait();
                    }

                    Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, async () =>
                    {
                        ContentDialog dialog = new ContentDialog
                        {
                            Title = "提示",
                            Content = "针对Excel文件的处理已成功完成",
                            CloseButtonText = "确定",
                            Background = Application.Current.Resources["DialogAcrylicBrush"] as Brush
                        };
                        _ = await dialog.ShowAsync();
                    }).AsTask().Wait();
                }
                catch (Exception ex)
                {
                    Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, async () =>
                    {
                        ContentDialog dialog = new ContentDialog
                        {
                            Title = "提示",
                            Content = "出现错误，将撤销所有更改：" + ex.Message,
                            CloseButtonText = "确定",
                            Background = Application.Current.Resources["DialogAcrylicBrush"] as Brush
                        };
                        _ = await dialog.ShowAsync();
                    }).AsTask().Wait();
                }
                finally
                {
                    Package.Dispose();
                    Package = null;
                }
            });

            Stack.Visibility = Visibility.Collapsed;
            Drag.Visibility = Visibility.Visible;
        }

        private async void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.DataView.Contains(StandardDataFormats.StorageItems))
            {
                IReadOnlyList<IStorageItem> FileList = await e.DataView.GetStorageItemsAsync();

                if (FileList.Count > 1)
                {
                    ContentDialog dialog = new ContentDialog
                    {
                        Title = "错误",
                        Content = "同时传入多个文件是不允许的",
                        CloseButtonText = "确定",
                        Background = Application.Current.Resources["DialogAcrylicBrush"] as Brush
                    };
                    _ = await dialog.ShowAsync();
                }
                else
                {
                    if (FileList.FirstOrDefault() is StorageFile InputFile)
                    {
                        switch (InputFile.FileType)
                        {
                            case ".xlsx":
                                File = await StorageFile.GetFileFromPathAsync(InputFile.Path);
                                OptionDialog Dialog = new OptionDialog();
                                if (await Dialog.ShowAsync() == ContentDialogResult.Primary)
                                {
                                    Mode = Dialog.Mode;
                                    ExcutionMethod = Dialog.ExcutionMethod;
                                    IsCoverOriginFile = Dialog.IsCoverOriginFile;

                                    using (IRandomAccessStream FileSteam = await File.OpenAsync(FileAccessMode.Read))
                                    {
                                        Package = new ExcelPackage(FileSteam.AsStream());
                                        CurrentSheet = Worksheets[0];
                                    }
                                    await Start_Process();
                                }
                                break;
                            default:
                                ContentDialog dia = new ContentDialog
                                {
                                    Title = "错误",
                                    Content = "文件格式错误，仅允许.xls格式的文件",
                                    CloseButtonText = "确定",
                                    Background = Application.Current.Resources["DialogAcrylicBrush"] as Brush
                                };
                                _ = await dia.ShowAsync();
                                break;
                        }
                    }

                    else
                    {
                        ContentDialog dialog = new ContentDialog
                        {
                            Title = "错误",
                            Content = "不允许传入文件夹，仅允许传入文件",
                            CloseButtonText = "确定",
                            Background = Application.Current.Resources["DialogAcrylicBrush"] as Brush
                        };
                        _ = await dialog.ShowAsync();
                    }
                }
            }
        }

        private void Grid_DragEnter(object sender, DragEventArgs e)
        {
            e.DragUIOverride.Caption = "添加文件 o(^▽^)o";
            e.DragUIOverride.IsCaptionVisible = true;
            e.DragUIOverride.IsContentVisible = true;
            e.DragUIOverride.IsGlyphVisible = true;

            e.AcceptedOperation = DataPackageOperation.Copy;
            e.Handled = true;
        }
    }

    public enum ExcutionMode
    {
        ComputeDataOnly = 0,
        ComputeDataAndLine = 1
    }

    public enum ExcutionMethod
    {
        Primary = 0,
        Secondary = 1,
        Third = 2,
        Forth = 3,
        Sixth = 5
    }
}

