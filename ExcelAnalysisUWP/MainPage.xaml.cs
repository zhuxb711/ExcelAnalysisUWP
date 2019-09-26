using AnimationEffectProvider;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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
        private HSSFWorkbook HWorkBook;
        private ISheet Sheet;
        private ExcutionMode Mode;
        private const int DataStartRow = 69;
        private const int DataStartColumn = 2;
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

        private void DrawLineCore(ICell UpCell, ICell DownCell)
        {
            lock (this)
            {
                HSSFPatriarch Patriarch = (HSSFPatriarch)Sheet.CreateDrawingPatriarch();
                HSSFClientAnchor Anchor = new HSSFClientAnchor(Sheet.GetColumnWidth(UpCell.ColumnIndex) / 2, UpCell.Row.Height, Sheet.GetColumnWidth(DownCell.ColumnIndex) / 2, 0, UpCell.ColumnIndex, UpCell.RowIndex, DownCell.ColumnIndex, DownCell.RowIndex);
                HSSFSimpleShape Line = Patriarch.CreateSimpleShape(Anchor);

                Line.ShapeType = HSSFSimpleShape.OBJECT_TYPE_LINE;
                Line.LineStyle = HSSFShape.LINESTYLE_SOLID;
                Line.LineWidth = 12700;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void DrawLineTask(int CurrentRow, int DataLength)
        {
            for (int i = DataStartColumn; i < DataLength + DataStartColumn; i++)
            {
                if (Sheet.GetRowOrCreate(CurrentRow).GetCellOrCreate(i).ToString() == Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(i).ToString())
                {
                    ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(i);
                    ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i - 1);
                    if (ZeroObject.ToString() == "0")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
                else
                {
                    ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i);
                    ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(i - 1);
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
            LeftMove.CopyTo(Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(DataStartColumn));

            for (int i = DataStartColumn; i < DataLength + DataStartColumn; i++)
            {
                if (Sheet.GetRowOrCreate(CurrentRow).GetCellOrCreate(i).ToString() == Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(i).ToString())
                {
                    Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(i).SetCellValue("1");
                }
                else
                {
                    Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i).SetCellValue("0");
                }
            }

            if (Mode == ExcutionMode.ComputeDataAndLine)
            {
                DrawLineTask(CurrentRow, DataLength);

                Pro.Report(null);
            }

            CellRangeAddress Range = new CellRangeAddress(CurrentRow + 4, CurrentRow + 4, DataStartColumn, DataLength + 2);
            Range.CopyTo(Sheet.GetRowOrCreate(CurrentRow + GoupDistance).GetCellOrCreate(DataStartColumn));

            for (int i = DataStartColumn; i < DataLength + DataStartColumn; i++)
            {
                if (Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i).ToString() != "0")
                {
                    Sheet.GetRowOrCreate(CurrentRow + GoupDistance).GetCellOrCreate(i).SetCellValue("1");
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void GroupProcessorSecondaryMethod(int CurrentRow, int DataLength)
        {
            int Row = CurrentRow, Length = DataLength;
            for (int i = DataStartColumn; i < Length + DataStartColumn; i++)
            {
                if (Sheet.GetRowOrCreate(Row).GetCellOrCreate(i).ToString() == Sheet.GetRowOrCreate(Row + 1).GetCellOrCreate(i).ToString())
                {
                    Sheet.GetRowOrCreate(Row + 2).GetCellOrCreate(i).SetCellValue("1");
                }
                else
                {
                    Sheet.GetRowOrCreate(Row + 4).GetCellOrCreate(i).SetCellValue("0");
                }
            }

            if (DataLength != 1)
            {
                CellRangeAddress Range = new CellRangeAddress(CurrentRow, CurrentRow, DataStartColumn, DataStartColumn + TotalDataLength);
                Range.CopyTo(Sheet.GetRowOrCreate(CurrentRow + 6).GetCellOrCreate(DataStartColumn));

                CellRangeAddress Left = new CellRangeAddress(CurrentRow + 1, CurrentRow + 1, DataStartColumn + 1, DataStartColumn + DataLength);
                Left.CopyTo(Sheet.GetRowOrCreate(CurrentRow + 7).GetCellOrCreate(DataStartColumn));
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
            ICell AddObject = Sheet.GetRowOrCreate(CurrentRow).GetCellOrCreate(DataStartColumn + DataLength);

            int CompareIndex = DataStartColumn + DataLength - 1;
            Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(CompareIndex).SetCellValue(AddObject.ToString());

            if (Sheet.GetRowOrCreate(CurrentRow).GetCellOrCreate(CompareIndex).ToString() == Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(CompareIndex).ToString())
            {
                ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(CompareIndex);
                OneObject.SetCellValue("1");
                Sheet.GetRowOrCreate(CurrentRow + 6).GetCellOrCreate(CompareIndex).SetCellValue(OneObject.ToString());

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(CompareIndex - 1);
                    if (ZeroObject.ToString() == "0")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }
            else
            {
                ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(CompareIndex);
                ZeroObject.SetCellValue("0");
                Sheet.GetRowOrCreate(CurrentRow + 6).GetCellOrCreate(CompareIndex).SetCellValue(ZeroObject.ToString());

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(CompareIndex - 1);
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

            if (Sheet.GetRowOrCreate(CurrentRow).GetCellOrCreate(CompareIndex).ToString() == Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(CompareIndex).ToString())
            {
                ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(CompareIndex);
                OneObject.SetCellValue("1");

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(CompareIndex - 1);
                    if (ZeroObject.ToString() == "0")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }
            else
            {
                ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(CompareIndex);
                ZeroObject.SetCellValue("0");

                if (Mode == ExcutionMode.ComputeDataAndLine)
                {
                    ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(CompareIndex - 1);
                    if (OneObject.ToString() == "1")
                    {
                        DrawLineCore(OneObject, ZeroObject);
                    }
                }
            }

            Pro.Report(null);

            if (DataLength != 1)
            {
                CellRangeAddress Range = new CellRangeAddress(DataStartRow, DataStartRow, DataStartColumn, DataStartColumn + TotalDataLength);
                Range.CopyTo(Sheet.GetRowOrCreate(CurrentRow + 6).GetCellOrCreate(DataStartColumn));
                ICell Origin = Sheet.GetRowOrCreate(CurrentRow).GetCellOrCreate(DataStartColumn + TotalDataLength);
                Sheet.GetRowOrCreate(CurrentRow + 7).GetCellOrCreate(CompareIndex - 1).SetCellValue(Origin.ToString());
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
                    CellRangeAddress Range = new CellRangeAddress(CurrentRow + 1, CurrentRow + 1, DataStartColumn + 1, DataStartColumn + DataLength);
                    Range.CopyTo(Sheet.GetRowOrCreate(CurrentRow + 7).GetCellOrCreate(DataStartColumn));
                }
            }

            //操作一行
            if ((DataStartColumn + DataLength - 20) >= DataStartColumn)
            {
                for (int i = DataStartColumn + DataLength - 20; i < DataLength + DataStartColumn; i++)
                {
                    if (Sheet.GetRowOrCreate(CurrentRow).Cells[i].ToString() == Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(i).ToString())
                    {
                        Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(i).SetCellValue("1");
                    }
                    else
                    {
                        Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i).SetCellValue("0");
                    }
                }
            }
            else
            {
                for (int i = DataStartColumn; i < DataLength + DataStartColumn; i++)
                {
                    if (Sheet.GetRowOrCreate(CurrentRow).GetCellOrCreate(i).ToString() == Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(i).ToString())
                    {
                        Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(i).SetCellValue("1");
                    }
                    else
                    {
                        Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i).SetCellValue("0");
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
                        if (Sheet.GetRowOrCreate(CurrentRow).GetCellOrCreate(i).ToString() == Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(i).ToString())
                        {
                            ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(i);
                            ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i - 1);
                            if (ZeroObject.ToString() == "0")
                            {
                                DrawLineCore(OneObject, ZeroObject);
                            }
                        }
                        else
                        {
                            ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i);
                            ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(i - 1);
                            if (OneObject.ToString() == "1")
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
                        if (Sheet.GetRowOrCreate(CurrentRow).GetCellOrCreate(i).ToString() == Sheet.GetRowOrCreate(CurrentRow + 1).GetCellOrCreate(i).ToString())
                        {
                            ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(i);
                            ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i - 1);
                            if (ZeroObject.ToString() == "0")
                            {
                                DrawLineCore(OneObject, ZeroObject);
                            }
                        }
                        else
                        {
                            ICell ZeroObject = Sheet.GetRowOrCreate(CurrentRow + 4).GetCellOrCreate(i);
                            ICell OneObject = Sheet.GetRowOrCreate(CurrentRow + 2).GetCellOrCreate(i - 1);
                            if (OneObject.ToString() == "1")
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
                 });
            });

            await Task.Run(() =>
            {
                try
                {
                    int ColumnsCount = 0;

                    for (int i = DataStartColumn; ; i++)
                    {
                        if (Sheet.GetRowOrCreate(DataStartRow).GetCellOrCreate(i).ToString() == "")
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
                                CellRangeAddress Range = new CellRangeAddress(DataStartRow, DataStartRow, DataStartColumn, DataStartColumn + ColumnsCount);
                                Range.CopyTo(Sheet.GetRowOrCreate(DataStartRow + 2).GetCellOrCreate(DataStartColumn));

                                ProcessDelegate = GroupProcessorPrimaryMethod;
                                break;
                            }
                        case ExcutionMethod.Secondary:
                            {
                                CellRangeAddress Range = new CellRangeAddress(DataStartRow, DataStartRow, DataStartColumn, DataStartColumn + ColumnsCount);
                                Range.CopyTo(Sheet.GetRowOrCreate(DataStartRow + 2).GetCellOrCreate(DataStartColumn));

                                CellRangeAddress LeftMove = new CellRangeAddress(DataStartRow + 2, DataStartRow + 2, DataStartColumn + 1, DataStartColumn + ColumnsCount);
                                LeftMove.CopyTo(Sheet.GetRowOrCreate(DataStartRow + 3).GetCellOrCreate(DataStartColumn));

                                ProcessDelegate = GroupProcessorSecondaryMethod;
                                break;
                            }
                        case ExcutionMethod.Third:
                            {
                                ICell AddObject = Sheet.GetRowOrCreate(DataStartRow).GetCellOrCreate(DataStartColumn + ColumnsCount);
                                Sheet.GetRowOrCreate(DataStartRow + 2).GetCellOrCreate(DataStartColumn + ColumnsCount).SetCellValue(AddObject.ToString());

                                ProcessDelegate = GroupProcessorThirdMethod;
                                break;
                            }
                        case ExcutionMethod.Forth:
                            {
                                ICell AddObject = Sheet.GetRowOrCreate(DataStartRow).GetCellOrCreate(DataStartColumn + ColumnsCount);
                                Sheet.GetRowOrCreate(DataStartRow + 2).GetCellOrCreate(DataStartColumn + ColumnsCount).SetCellValue(AddObject.ToString());

                                Sheet.GetRowOrCreate(DataStartRow + 3).GetCellOrCreate(DataStartColumn + ColumnsCount - 1).SetCellValue(AddObject.ToString());

                                ProcessDelegate = GroupProcessorForthMethod;
                                break;
                            }
                        case ExcutionMethod.Sixth:
                            {
                                //直接生成250个模块的原始数据
                                CellRangeAddress Origin = new CellRangeAddress(DataStartRow, DataStartRow, DataStartColumn, DataStartColumn + ColumnsCount);
                                for (int i = 0; i < 250; i++)
                                {
                                    Origin.CopyTo(Sheet.GetRowOrCreate(DataStartRow + 2 + 6 * i).GetCellOrCreate(DataStartColumn));
                                }

                                CellRangeAddress Left = new CellRangeAddress(DataStartRow, DataStartRow, DataStartColumn + ColumnsCount - 19, DataStartColumn + ColumnsCount);
                                for (int i = 0; i < 250; i++)
                                {
                                    Left.CopyTo(Sheet.GetRowOrCreate(DataStartRow + 3 + i * 6).GetCellOrCreate(DataStartColumn + ColumnsCount - 20 - i));
                                    //复制错位行能完整复制时候的数
                                    if ((DataStartColumn + ColumnsCount - 20 - i) <= DataStartColumn)
                                    {
                                        break;
                                    }
                                }

                                CellRangeAddress LeftMove = new CellRangeAddress(DataStartRow + 2, DataStartRow + 2, DataStartColumn + 1 + ColumnsCount - 20, DataStartColumn + ColumnsCount);
                                LeftMove.CopyTo(Sheet.GetRowOrCreate(DataStartRow + 3).GetCellOrCreate(DataStartColumn + ColumnsCount - 20));

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
                        HWorkBook.Write(Stream);
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
                    HWorkBook.Close();
                    HWorkBook = null;
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
                            case ".xls":
                                File = await StorageFile.GetFileFromPathAsync(InputFile.Path);
                                OptionDialog Dialog = new OptionDialog();
                                if (await Dialog.ShowAsync() == ContentDialogResult.Primary)
                                {
                                    Mode = Dialog.Mode;
                                    ExcutionMethod = Dialog.ExcutionMethod;
                                    IsCoverOriginFile = Dialog.IsCoverOriginFile;

                                    using (IRandomAccessStream FileSteam = await File.OpenAsync(FileAccessMode.Read))
                                    {
                                        HWorkBook = new HSSFWorkbook(FileSteam.AsStream());
                                        Sheet = HWorkBook.GetSheetAt(0);
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

    public static class Extention
    {
        public static void CopyTo(this CellRangeAddress Range, ICell DestinationCell)
        {
            for (var RowNum = Range.FirstRow; RowNum <= Range.LastRow; RowNum++)
            {
                IRow DestinationRow = DestinationCell.Row;

                for (int ColNum = Range.FirstColumn, Length = 0; ColNum <= Range.LastColumn; ColNum++, Length++)
                {
                    DestinationCell.Row.GetCellOrCreate(DestinationCell.ColumnIndex + Length).SetCellValue(DestinationCell.Sheet.GetRow(RowNum).GetCellOrCreate(ColNum).ToString());
                }
            }
        }

        public static ICell GetCellOrCreate(this IRow Row, int cellnum)
        {
            ICell Cell = Row.GetCell(cellnum, MissingCellPolicy.CREATE_NULL_AS_BLANK);
            return Cell ?? Row.CreateCell(cellnum, CellType.String);
        }

        public static IRow GetRowOrCreate(this ISheet sheet, int rownum)
        {
            IRow Row = sheet.GetRow(rownum);
            return Row ?? sheet.CreateRow(rownum);
        }
    }
}

