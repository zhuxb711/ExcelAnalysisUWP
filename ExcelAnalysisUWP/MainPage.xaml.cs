using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;
using Windows.Storage.Streams;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace ExcelAnalysisUWP
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public sealed partial class MainPage : Page
    {
        private IWorkbook WorkBook;
        private ISheet Sheet;
        private CancellationTokenSource Cancellation;
        private ExcutionMode Mode;
        private bool IsRunning = false;
        private const int DataStartRow = 70;
        private const int DataStartColumn = 3;
        private const int GoupDistance = 6;
        private DateTime StartTime;
        private ExcutionMethod ExcutionMethod;
        private Action<int, int> ProcessDelegate;
        private int TotalDataLength;
        private IProgress<object> Pro;
        private int Tick;
        private List<Task> TaskList;
        public ExcelType InputType;

        public MainPage()
        {
            InitializeComponent();
            Loaded += MainPage_Loaded;
        }

        private async void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            Cancellation = new CancellationTokenSource();
            TaskList = new List<Task>();

            OptionDialog dialog = new OptionDialog(Visibility.Visible);
            if (await dialog.ShowAsync() == ContentDialogResult.Primary)
            {
                Mode = dialog.Mode;
                ExcutionMethod = dialog.ExcutionMethod;
                IRandomAccessStream FileSteam = await dialog.InputFile.OpenAsync(FileAccessMode.ReadWrite);
                if (dialog.InputFile.FileType == ".xlsx")
                {
                    WorkBook = new XSSFWorkbook(FileSteam.AsStream());
                    InputType = ExcelType.XLSX;
                }
                else
                {
                    WorkBook = new HSSFWorkbook(FileSteam.AsStream());
                    InputType = ExcelType.XLS;
                }
                Sheet = WorkBook.GetSheetAt(0);
            }
        }

        private void DrawLineCore(ICell UpCell,ICell DownCell)
        {
            if (InputType == ExcelType.XLS)
            {
                HSSFPatriarch Patriarch = (HSSFPatriarch)Sheet.CreateDrawingPatriarch();
                HSSFClientAnchor Anchor = new HSSFClientAnchor(Sheet.GetColumnWidth(UpCell.ColumnIndex) / 2, UpCell.Row.Height, Sheet.GetColumnWidth(DownCell.ColumnIndex) / 2, 0, DownCell.ColumnIndex, UpCell.RowIndex, DownCell.ColumnIndex, DownCell.RowIndex);
                HSSFSimpleShape Line = Patriarch.CreateSimpleShape(Anchor);

                Line.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
                Line.LineStyle = HSSFShape.LINESTYLE_SOLID;
                Line.LineWidth = 6350;
            }
            else
            {
                XSSFDrawing Patriarch = (XSSFDrawing)Sheet.CreateDrawingPatriarch();
                XSSFClientAnchor Anchor = new XSSFClientAnchor(Sheet.GetColumnWidth(UpCell.ColumnIndex) / 2, UpCell.Row.Height, Sheet.GetColumnWidth(DownCell.ColumnIndex) / 2, 0, UpCell.ColumnIndex, UpCell.RowIndex, DownCell.ColumnIndex, DownCell.RowIndex);
                XSSFSimpleShape Line = Patriarch.CreateSimpleShape(Anchor);

                Line.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
                Line.LineStyle = HSSFShape.LINESTYLE_SOLID;
                Line.LineWidth = 6350;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void DrawLineTask(int CurrentRow, int DataLength)
        {
            for (int i = DataStartColumn; i < DataLength + DataStartColumn && !Cancellation.IsCancellationRequested; i++)
            {
                if (Sheet.GetRow(CurrentRow).Cells[i].ToString() == Sheet.GetRow(CurrentRow + 1).Cells[i].ToString())
                {
                    ICell OneObject = Sheet.GetRow(CurrentRow + 2).Cells[i];
                    ICell ZeroObject = Sheet.GetRow(CurrentRow + 4).Cells[i - 1];
                    if (ZeroObject.ToString() == "0")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
                else
                {
                    ICell ZeroObject = Sheet.GetRow(CurrentRow + 4).Cells[i];
                    ICell OneObject = Sheet.GetRow(CurrentRow + 2).Cells[i - 1];
                    if (OneObject.ToString() == "1")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void GroupProcessorPrimaryMethod(int CurrentRow, int DataLength)
        {
            CellRangeAddress LeftMove = new CellRangeAddress(CurrentRow, CurrentRow, DataStartColumn + 1, DataStartColumn + DataLength);
            LeftMove.CopyTo(Sheet.GetRow(CurrentRow + 1).Cells[DataStartColumn]);

            for (int i = DataStartColumn; i < DataLength + DataStartColumn && !Cancellation.IsCancellationRequested; i++)
            {
                if (Sheet.GetRow(CurrentRow).Cells[i].ToString() == Sheet.GetRow(CurrentRow + 1).Cells[i].ToString())
                {
                    Sheet.GetRow(CurrentRow + 2).Cells[i].SetCellValue("1");
                }
                else
                {
                    Sheet.GetRow(CurrentRow + 4).Cells[i].SetCellValue("0");
                }
            }

            if (Mode == ExcutionMode.ComputeDataAndLine)
            {
                TaskList.Add(Task.Factory.StartNew((Para) =>
                {
                    var Pair = (KeyValuePair<int, int>)Para;
                    DrawLineTask(Pair.Key, Pair.Value);

                    if (!Cancellation.IsCancellationRequested)
                    {
                        Pro.Report(null);
                    }
                }, new KeyValuePair<int, int>(CurrentRow, DataLength)));
            }

            CellRangeAddress Range = new CellRangeAddress(CurrentRow + 4, CurrentRow + 4, DataStartColumn, DataLength + 2);
            Range.CopyTo(Sheet.GetRow(CurrentRow + GoupDistance).Cells[DataStartColumn]);

            for (int i = DataStartColumn; i < DataLength + DataStartColumn && !Cancellation.IsCancellationRequested; i++)
            {
                if (Sheet.GetRow(CurrentRow + 4).Cells[i].ToString() != "0")
                {
                    Sheet.GetRow(CurrentRow + GoupDistance).Cells[i].SetCellValue("1");
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void GroupProcessorSecondaryMethod(int CurrentRow, int DataLength)
        {
            TaskList.Add(Task.Factory.StartNew((Para) =>
            {
                KeyValuePair<int, int> Pair = (KeyValuePair<int, int>)Para;
                int Row = Pair.Key, Length = Pair.Value;
                for (int i = DataStartColumn; i < Length + DataStartColumn && !Cancellation.IsCancellationRequested; i++)
                {
                    if (Sheet.GetRow(Row).Cells[i].ToString() == Sheet.GetRow(Row + 1).Cells[i].ToString())
                    {
                        Sheet.GetRow(Row + 2).Cells[i].SetCellValue("1");
                    }
                    else
                    {
                        Sheet.GetRow(Row + 4).Cells[i].SetCellValue("0");
                    }
                }
                return Pair;
            }, new KeyValuePair<int, int>(CurrentRow, DataLength)).ContinueWith((task) =>
            {
                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    DrawLineTask(task.Result.Key, task.Result.Value);
                }
                if (!Cancellation.IsCancellationRequested)
                {
                    Pro.Report(null);
                }
            }));

            if (DataLength != 1)
            {
                CellRangeAddress Range = new CellRangeAddress(CurrentRow, CurrentRow, DataStartColumn, DataStartColumn + TotalDataLength);
                Range.CopyTo(Sheet.GetRow(CurrentRow + 6).Cells[DataStartColumn]);

                CellRangeAddress Left = new CellRangeAddress(CurrentRow + 1, CurrentRow + 1, DataStartColumn + 1, DataStartColumn + DataLength);
                Left.CopyTo(Sheet.GetRow(CurrentRow + 7).Cells[DataStartColumn]);
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void GroupProcessorThirdMethod(int CurrentRow, int DataLength)
        {
            ICell AddObject = Sheet.GetRow(CurrentRow).Cells[DataStartColumn + DataLength];

            int CompareIndex = DataStartColumn + DataLength - 1;
            Sheet.GetRow(CurrentRow + 1).Cells[CompareIndex].SetCellValue(AddObject.ToString());

            if (Sheet.GetRow(CurrentRow).Cells[CompareIndex].ToString() == Sheet.GetRow(CurrentRow + 1).Cells[CompareIndex].ToString())
            {
                ICell OneObject = Sheet.GetRow(CurrentRow + 2).Cells[CompareIndex];
                OneObject.SetCellValue("1");
                Sheet.GetRow(CurrentRow + 6).Cells[CompareIndex].SetCellValue(OneObject.ToString());

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ICell ZeroObject = Sheet.GetRow(CurrentRow + 4).Cells[CompareIndex - 1];
                    if (ZeroObject.ToString() == "0")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }
            else
            {
                ICell ZeroObject = Sheet.GetRow(CurrentRow + 4).Cells[CompareIndex];
                ZeroObject.SetCellValue("0");
                Sheet.GetRow(CurrentRow + 6).Cells[CompareIndex].SetCellValue(ZeroObject.ToString());

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ICell OneObject = Sheet.GetRow(CurrentRow + 2).Cells[CompareIndex - 1];
                    if (OneObject.ToString() == "1")
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

            if (((Range)Sheet.Cells[CurrentRow, CompareIndex]).Text == ((Range)Sheet.Cells[CurrentRow + 1, CompareIndex]).Text)
            {
                Range OneObject = Sheet.Cells[CurrentRow + 2, CompareIndex];
                OneObject.Value = "1";
                //((Range)Sheet.Cells[CurrentRow + 6, CompareIndex]).Value = OneObject.Text;

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    Range ZeroObject = (Range)Sheet.Cells[CurrentRow + 4, CompareIndex - 1];
                    if (ZeroObject.Text == "0")
                    {
                        _ = Sheet.Shapes.AddLine(OneObject.Left + OneObject.Width / 2, OneObject.Top + OneObject.Height, ZeroObject.Left + ZeroObject.Width / 2, ZeroObject.Top);
                    }
                }
            }
            else
            {
                Range ZeroObject = Sheet.Cells[CurrentRow + 4, CompareIndex];
                ZeroObject.Value = "0";
                //((Range)Sheet.Cells[CurrentRow + 6, CompareIndex]).Value = ZeroObject.Text;

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    Range OneObject = Sheet.Cells[CurrentRow + 2, CompareIndex - 1];
                    if (OneObject.Text == "1")
                    {
                        _ = Sheet.Shapes.AddLine(OneObject.Left + OneObject.Width / 2, OneObject.Top + OneObject.Height, ZeroObject.Left + ZeroObject.Width / 2, ZeroObject.Top);
                    }
                }
            }

            Pro.Report(null);

            if (DataLength != 1)
            {
                _ = ((Range)Sheet.Range[Sheet.Cells[DataStartRow, DataStartColumn], Sheet.Cells[DataStartRow, DataStartColumn + TotalDataLength]]).Copy(Sheet.Cells[CurrentRow + 6, DataStartColumn]);
                Range Origin = Sheet.Cells[CurrentRow, DataStartColumn + TotalDataLength];
                ((Range)Sheet.Cells[CurrentRow + 7, CompareIndex - 1]).Value = Origin.Text;
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
                    Range Left = Sheet.Range[Sheet.Cells[CurrentRow + 1, DataStartColumn + 1], Sheet.Cells[CurrentRow + 1, DataStartColumn + DataLength]];
                    _ = Left.Copy(Sheet.Cells[CurrentRow + 7, DataStartColumn]);
                }
            }

            //操作一行
            if ((DataStartColumn + DataLength - 20) >= DataStartColumn)
            {
                for (int i = DataStartColumn + DataLength - 20; i < DataLength + DataStartColumn && !Cancellation.IsCancellationRequested; i++)
                {
                    if (((Range)Sheet.Cells[CurrentRow, i]).Text == ((Range)Sheet.Cells[CurrentRow + 1, i]).Text)
                    {
                        Sheet.Cells[CurrentRow + 2, i] = "1";
                    }
                    else
                    {
                        Sheet.Cells[CurrentRow + 4, i] = "0";
                    }
                }
            }
            else
            {
                for (int i = DataStartColumn; i < DataLength + DataStartColumn && !Cancellation.IsCancellationRequested; i++)
                {
                    if (((Range)Sheet.Cells[CurrentRow, i]).Text == ((Range)Sheet.Cells[CurrentRow + 1, i]).Text)
                    {
                        Sheet.Cells[CurrentRow + 2, i] = "1";
                    }
                    else
                    {
                        Sheet.Cells[CurrentRow + 4, i] = "0";
                    }
                }
            }

            //画线操作
            if (Mode == ExcutionMode.ComputeDataAndLine)
            {
                if ((DataStartColumn + DataLength - 20) >= DataStartColumn)
                {
                    for (int i = DataStartColumn + DataLength - 20; i < DataLength + DataStartColumn && !Cancellation.IsCancellationRequested; i++)
                    {
                        if (((Range)Sheet.Cells[CurrentRow, i]).Text == ((Range)Sheet.Cells[CurrentRow + 1, i]).Text)
                        {
                            Range OneObject = Sheet.Cells[CurrentRow + 2, i];
                            Range ZeroObject = Sheet.Cells[CurrentRow + 4, i - 1];
                            if (ZeroObject.Text == "0")
                            {
                                _ = Sheet.Shapes.AddLine(OneObject.Left + OneObject.Width / 2, OneObject.Top + OneObject.Height, ZeroObject.Left + ZeroObject.Width / 2, ZeroObject.Top);
                            }
                        }
                        else
                        {
                            Range ZeroObject = Sheet.Cells[CurrentRow + 4, i];
                            Range OneObject = Sheet.Cells[CurrentRow + 2, i - 1];
                            if (OneObject.Text == "1")
                            {
                                _ = Sheet.Shapes.AddLine(OneObject.Left + OneObject.Width / 2, OneObject.Top + OneObject.Height, ZeroObject.Left + ZeroObject.Width / 2, ZeroObject.Top);
                            }
                        }
                    }
                }
                else
                {
                    for (int i = DataStartColumn; i < DataLength + DataStartColumn && !Cancellation.IsCancellationRequested; i++)
                    {
                        if (((Range)Sheet.Cells[CurrentRow, i]).Text == ((Range)Sheet.Cells[CurrentRow + 1, i]).Text)
                        {
                            Range OneObject = Sheet.Cells[CurrentRow + 2, i];
                            Range ZeroObject = Sheet.Cells[CurrentRow + 4, i - 1];
                            if (ZeroObject.Text == "0")
                            {
                                _ = Sheet.Shapes.AddLine(OneObject.Left + OneObject.Width / 2, OneObject.Top + OneObject.Height, ZeroObject.Left + ZeroObject.Width / 2, ZeroObject.Top);
                            }
                        }
                        else
                        {
                            Range ZeroObject = Sheet.Cells[CurrentRow + 4, i];
                            Range OneObject = Sheet.Cells[CurrentRow + 2, i - 1];
                            if (OneObject.Text == "1")
                            {
                                _ = Sheet.Shapes.AddLine(OneObject.Left + OneObject.Width / 2, OneObject.Top + OneObject.Height, ZeroObject.Left + ZeroObject.Width / 2, ZeroObject.Top);
                            }
                        }
                    }
                }
            }
            if (!Cancellation.IsCancellationRequested)
            {
                Pro.Report(null);
            }
        }

        private async void Start_Click(object sender, RoutedEventArgs e)
        {
            if (Start.Content.ToString() == "取消")
            {
                Cancellation.Cancel();
                return;
            }

            Progress.IsIndeterminate = true;
            IsRunning = true;

            if (Setting.Visibility == Visibility.Visible)
            {
                Setting.Visibility = Visibility.Collapsed;
                Drag.Visibility = Visibility.Visible;
            }

            Pro = new Progress<object>(async (o) =>
            {
                double CurrentValue = 100 - (Tick-- - 1) * (100f / TotalDataLength);
                await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, () =>
                 {
                     if (Progress.IsIndeterminate == true)
                     {
                         Progress.IsIndeterminate = false;
                     }

                     Progress.Value = CurrentValue;

                     double CeilingValue = Math.Ceiling(CurrentValue);

                     if (CeilingValue > 100)
                     {
                         return;
                     }

                     string CurrentValueText = CeilingValue.ToString() + " %";

                     TimeSpan Span = DateTime.Now - StartTime;

                     double TimeToFinish = Span.TotalSeconds / CeilingValue * (100 - CeilingValue);

                     if (double.IsNaN(TimeToFinish))
                     {
                         ProText.Text = CurrentValueText + "   剩余时间: Unknown";
                     }
                     else
                     {
                         string Duration = TimeSpan.FromSeconds(TimeToFinish).ToString(@"hh\:mm\:ss");
                         ProText.Text = CurrentValueText + "   剩余时间: " + Duration;
                     }


                     if (CeilingValue.ToString() == "100")
                     {
                         ProText.Text = CurrentValueText;
                     }
                 });
            });

            await Task.Factory.StartNew(() =>
            {
                try
                {
                    int ColumnsCount = 0;

                    for (int i = DataStartColumn; ; i++)
                    {
                        if (((Range)Sheet.Cells[DataStartRow, i]).Text == "")
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
                                Range OriginInput = Sheet.Range[Sheet.Cells[DataStartRow, DataStartColumn], Sheet.Cells[DataStartRow, DataStartColumn + ColumnsCount]];
                                _ = OriginInput.Copy(Sheet.Cells[DataStartRow + 2, DataStartColumn]);

                                ProcessDelegate = GroupProcessorPrimaryMethod;
                                break;
                            }
                        case ExcutionMethod.Secondary:
                            {
                                Range Origin = Sheet.Range[Sheet.Cells[DataStartRow, DataStartColumn], Sheet.Cells[DataStartRow, DataStartColumn + ColumnsCount]];
                                _ = Origin.Copy(Sheet.Cells[DataStartRow + 2, DataStartColumn]);

                                Range LeftMove = Sheet.Range[Sheet.Cells[DataStartRow + 2, DataStartColumn + 1], Sheet.Cells[DataStartRow + 2, DataStartColumn + ColumnsCount]];
                                _ = LeftMove.Copy(Sheet.Cells[DataStartRow + 3, DataStartColumn]);

                                ProcessDelegate = GroupProcessorSecondaryMethod;
                                break;
                            }
                        case ExcutionMethod.Third:
                        case ExcutionMethod.Fifth:
                            {
                                Range AddObject = Sheet.Cells[DataStartRow, DataStartColumn + ColumnsCount];
                                ((Range)Sheet.Cells[DataStartRow + 2, DataStartColumn + ColumnsCount]).Value = AddObject.Text;

                                ProcessDelegate = GroupProcessorThirdMethod;
                                break;
                            }
                        case ExcutionMethod.Forth:
                            {
                                Range AddObject = Sheet.Cells[DataStartRow, DataStartColumn + ColumnsCount];
                                ((Range)Sheet.Cells[DataStartRow + 2, DataStartColumn + ColumnsCount]).Value = AddObject.Text;

                                ((Range)Sheet.Cells[DataStartRow + 3, DataStartColumn + ColumnsCount - 1]).Value = AddObject.Text;

                                ProcessDelegate = GroupProcessorForthMethod;
                                break;
                            }
                        case ExcutionMethod.Sixth:
                            {
                                //直接生成250个模块的原始数据
                                Range Origin = Sheet.Range[Sheet.Cells[DataStartRow, DataStartColumn], Sheet.Cells[DataStartRow, DataStartColumn + ColumnsCount]];
                                for (int i = 0; i < 250; i++)
                                {
                                    _ = Origin.Copy(Sheet.Cells[DataStartRow + 2 + 6 * i, DataStartColumn]);
                                }


                                Range Left = Sheet.Range[Sheet.Cells[DataStartRow, DataStartColumn + ColumnsCount - 19], Sheet.Cells[DataStartRow, DataStartColumn + ColumnsCount]];
                                for (int i = 0; i < 250; i++)
                                {
                                    _ = Left.Copy(Sheet.Cells[DataStartRow + 3 + i * 6, DataStartColumn + ColumnsCount - 20 - i]);
                                    //复制错位行能完整复制时候的数
                                    if ((DataStartColumn + ColumnsCount - 20 - i) <= DataStartColumn)
                                        break;
                                }

                                ////复制错位行的数
                                //if (ColumnsCount != 1)
                                //{
                                //    //这里加个删除前面不要的数据的方法就好了
                                //    if ((DataStartColumn + ColumnsCount - 21) >= DataStartColumn)
                                //    {
                                //        Range Left = Sheet.Range[Sheet.Cells[DataStartRow + 1, DataStartColumn + 1 + ColumnsCount - 21], Sheet.Cells[DataStartRow + 1, DataStartColumn + ColumnsCount]];
                                //        _ = Left.Copy(Sheet.Cells[DataStartRow + 7, DataStartColumn + ColumnsCount - 21]);
                                //    }
                                //    else
                                //    {
                                //        Range Left = Sheet.Range[Sheet.Cells[DataStartRow + 1, DataStartColumn + 1], Sheet.Cells[DataStartRow + 1, DataStartColumn + ColumnsCount]];
                                //        _ = Left.Copy(Sheet.Cells[DataStartRow + 7, DataStartColumn]);
                                //    }
                                //}


                                Range LeftMove = Sheet.Range[Sheet.Cells[DataStartRow + 2, DataStartColumn + 1 + ColumnsCount - 20], Sheet.Cells[DataStartRow + 2, DataStartColumn + ColumnsCount]];
                                _ = LeftMove.Copy(Sheet.Cells[DataStartRow + 3, DataStartColumn + ColumnsCount - 20]);

                                ProcessDelegate = GroupProcessorSixthMethod;
                                break;
                            }
                    }

                    StartTime = DateTime.Now;
                    for (int CurrentRow = DataStartRow + 2, Count = 0; ColumnsCount > 0 && Count < 250 && !Cancellation.IsCancellationRequested; CurrentRow += GoupDistance, ColumnsCount--, Count++)
                    {
                        ProcessDelegate(CurrentRow, ColumnsCount);

                        if (TaskList.Count >= 4)
                        {
                            Task.WaitAll(TaskList.ToArray());
                            TaskList.Clear();
                        }
                    }

                    if (Cancellation.IsCancellationRequested)
                    {
                        _ = MessageBox.Show("任务已取消，已撤销所有更改", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        Cancellation.Dispose();
                        Cancellation = null;
                        Cancellation = new CancellationTokenSource();
                    }
                    else
                    {
                        _ = MessageBox.Show("针对Excel文件的处理已成功完成\r总时间: " + (DateTime.Now - StartTime).ToString(@"hh\:mm\:ss"), "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        if (ExcutionMethod != ExcutionMethod.Fifth)
                        {
                            WorkBook.Save();
                        }
                    }
                }
                catch (Exception ex)
                {
                    _ = MessageBox.Show("出现错误，将撤销所有更改：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    WorkBook.Close();
                    WorkBook = null;
                }
            }, TaskCreationOptions.LongRunning);
        }

        private void Grid_DragOver(object sender, DragEventArgs e)
        {
            e.AcceptedOperation = DataPackageOperation.Copy;
            e.DragUIOverride.Caption = "放开以导入";
            e.DragUIOverride.IsCaptionVisible = true;
            e.Handled = true;
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
                        CloseButtonText = "确定"
                    };
                    _ = await dialog.ShowAsync();
                }
                else
                {
                    if (FileList.FirstOrDefault() is StorageFile InputFile)
                    {
                        OptionDialog dialog = new OptionDialog(Visibility.Collapsed);
                        if (await dialog.ShowAsync() == ContentDialogResult.Primary)
                        {
                            Mode = dialog.Mode;
                            ExcutionMethod = dialog.ExcutionMethod;
                            IRandomAccessStream FileSteam = await dialog.InputFile.OpenAsync(FileAccessMode.ReadWrite);

                            switch (InputFile.FileType)
                            {
                                case ".xlsx":
                                    WorkBook = new XSSFWorkbook(FileSteam.AsStream());
                                    Sheet = WorkBook.GetSheetAt(0);
                                    InputType = ExcelType.XLSX;
                                    break;
                                case ".xls":
                                    WorkBook = new HSSFWorkbook(FileSteam.AsStream());
                                    Sheet = WorkBook.GetSheetAt(0);
                                    InputType = ExcelType.XLS;
                                    break;
                                default:
                                    ContentDialog dia = new ContentDialog
                                    {
                                        Title = "错误",
                                        Content = "文件格式错误，仅允许.xlsx和.xls格式的文件",
                                        CloseButtonText = "确定"
                                    };
                                    _ = await dia.ShowAsync();
                                    break;
                            }
                        }
                    }
                    else
                    {
                        ContentDialog dialog = new ContentDialog
                        {
                            Title = "错误",
                            Content = "不允许传入文件夹，仅允许传入文件",
                            CloseButtonText = "确定"
                        };
                        _ = await dialog.ShowAsync();
                    }
                }
            }
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
        Fifth = 4,
        Sixth = 5
    }

    public enum ExcelType
    {
        XLS = 0,
        XLSX = 1
    }

    public static class Extention
    {
        public static void CopyTo(this CellRangeAddress Range, ICell DestinationCell)
        {
            for (var RowNum = Range.FirstRow; RowNum <= Range.LastRow; RowNum++)
            {
                IRow DestinationRow = DestinationCell.Row;

                for (int ColNum = Range.FirstColumn, Length = 0; ColNum <= Range.LastColumn; ColNum++, Length++)
                {
                    DestinationCell.Row.Cells[DestinationCell.ColumnIndex + Length].SetCellValue(Sheet.GetRow(RowNum).Cells[ColNum].ToString());
                }
            }
        }

    }
}

