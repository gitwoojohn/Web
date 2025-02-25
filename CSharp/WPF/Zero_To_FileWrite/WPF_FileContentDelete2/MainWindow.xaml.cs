﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Shell;
using WPF_FileContentDelete.ViewModels;

namespace WPF_FileContentDelete
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private GridViewColumnHeader listViewSortCol = null;
        private SortAdorner listViewSortAdorner = null;

        public OpenFileDialog OpenFileDialog { get; set; } = null;
        public Stack<string> DeleteSubDirs { get; set; } = new Stack<string>();

        //static bool _fileContents = false;

        public MainWindow()
        {
            // 실행 성능 향상.
            ProfileOptimization.SetProfileRoot(@"..\..\bin\Release");
            ProfileOptimization.StartProfile("profile");

            InitializeComponent();
            InitializeOpenFileDialog();

            // 작업표시줄 진행바 ( 초기화 안해주면 진행바 작동 불가 )
            TaskbarItemInfo.ProgressState = TaskbarItemProgressState.Normal;
        }

        private void InitializeOpenFileDialog()
        {
            OpenFileDialog = new OpenFileDialog
            {
                // 여러 파일 선택하기
                Multiselect = true,

                // Set the file dialog to filter for All files.
                Filter = "All files (*.*)|*.*",
                InitialDirectory = @"C:\",
                Title = "파일 선택",
                FileName = string.Empty
            };

            button_Delete.IsEnabled = false;
        }

        private void button_Select_Click(object sender, RoutedEventArgs e)
        {
            if (OpenFileDialog.ShowDialog() == true)
            {
                ListView_AddFile(OpenFileDialog.FileNames);
            }
        }

        private async void button_ZeroFill_Click(object sender, RoutedEventArgs e)
        {
            if (listView.Items.Count > 0)
            {
                button_ZeroFill.IsEnabled = false;

                progressBar.Value = 0;
                await ZeroFillFile(listView.Items.Count, new Progress<int>(ReportProgress));

                button_Delete.IsEnabled = true;
            }
            else
            {
                MessageBox.Show("Zero Fill 할 파일이 없습니다.");
            }
        }

        private async void Button_Delete_Click(object sender, RoutedEventArgs e)
        {
            button_Delete.IsEnabled = false;

            progressBar.Value = 0;
            await DeleteFileList(listView.Items.Count, new Progress<int>(ReportProgress));

            button_ZeroFill.IsEnabled = true;
        }

        private async ValueTask<bool> DeleteFileList(int totalCount, IProgress<int> progress)
        {
            int tempCount = 1;
            try
            {
                for (int i = 0; i < totalCount; i++)
                {
                    string NewFileName = Path.GetDirectoryName(listView.Items[i].ToString()) + @"\1.tmp";
                    string OldFileName = listView.Items[i].ToString();

                    // 생성, 마지막 액세스, 쓰기 현재 시간으로 설정.
                    File.SetCreationTime(OldFileName, DateTime.Now);
                    File.SetLastAccessTime(OldFileName, DateTime.Now);
                    File.SetLastWriteTime(OldFileName, DateTime.Now);

                    // 파일 이름 변경, 삭제
                    File.Move(OldFileName, NewFileName);
                    File.Delete(NewFileName);

                    if (progress != null)
                    {
                        progress.Report((tempCount * 100 / totalCount));
                    }
                    await Task.Delay(10);
                    tempCount++;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                Debug.WriteLine(e.Message);
                return false;
            }
            await Task.Delay(10);

            await DeleteSubDirectory(DeleteSubDirs);

            if (listView.Items.Count >= 0)
            {
                listView.Items.Clear();
            }

            return true;
        }

        private async ValueTask<bool> ZeroFillFile(int totalCount, IProgress<int> progress)
        {
            int tempCount = 1;
            try
            {
                //await the processing and delete file logic here
                for (int i = 0; i < totalCount; i++)
                {
                    if (progress != null)
                    {
                        progress.Report((tempCount * 100 / totalCount));
                    }
                    await StreamFileWrite(listView.Items[i].ToString());
                    tempCount++;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                Debug.WriteLine(e.Message);
                return false;
            }

            return true;

            //int tempCount = 1;
            //const int blockSize = 1024 * 64; // 8; // 16; 
            //byte[] data = new byte[ blockSize ];
            //List<Task> tasks = new List<Task>();
            //List<FileStream> sourceStreams = new List<FileStream>();

            //try
            //{
            //    //await the processing and delete file logic here
            //    for( int i = 0; i < totalCount; i++ )
            //    {
            //        //await StreamFileWrite( lstView.Items[ i ].ToString() );

            //        FileStream sourceStream = 
            //            new FileStream( lstView.Items[ i ].ToString(), 
            //                            FileMode.Open, 
            //                            FileAccess.Write, 
            //                            FileShare.None, 
            //                            bufferSize:4096, 
            //                            useAsync: true );

            //        if( progress != null )
            //        {
            //            progress.Report( ( tempCount * 100 / totalCount ) );
            //        }

            //        Task theTask = sourceStream.WriteAsync( data, 0, data.Length );
            //        sourceStreams.Add( sourceStream );

            //        tasks.Add( theTask );
            //        tempCount++;
            //    }
            //    await Task.WhenAll( tasks );
            //}

            //catch( Exception e )
            //{
            //    Debug.Write( e.Message );
            //}
            //finally
            //{
            //    foreach(FileStream sourceStream in sourceStreams)
            //    {
            //        sourceStream.Close();
            //    }
            //}
        }

        // 관리자 권한 실행중인지 판단
        //public static bool IsAdministrator()
        //{
        //    WindowsIdentity identity = WindowsIdentity.GetCurrent();

        //    if( null != identity )
        //    {
        //        WindowsPrincipal principal = new WindowsPrincipal( identity );
        //        return principal.IsInRole( WindowsBuiltInRole.Administrator );
        //    }

        //    return false;
        //}

        private async ValueTask<bool> StreamFileWrite(string filePath)
        {
            const int blockSize = 1024 * 64; // 8; // 16; 
            byte[] data = new byte[blockSize];

            //if( IsAdministrator() == false )
            //{
            //    try
            //    {
            //        ProcessStartInfo procInfo = new ProcessStartInfo();
            //        procInfo.UseShellExecute = true;
            //        procInfo.FileName = Process.GetCurrentProcess().MainModule.FileName;
            //        procInfo.WorkingDirectory = Environment.CurrentDirectory;
            //        procInfo.Verb = "runas";
            //        Process.Start( procInfo );
            //    }
            //    catch( Exception ex )
            //    {
            //        MessageBox.Show( ex.Message.ToString() );
            //    }
            //}
            //// 권한 상승
            //DirectoryInfo dInfo = new DirectoryInfo( filePath );
            //DirectorySecurity dSecurity = dInfo.GetAccessControl();
            //dSecurity.AddAccessRule( new FileSystemAccessRule( 
            //                            new SecurityIdentifier( 
            //                                WellKnownSidType.WorldSid, null ), 
            //                                FileSystemRights.FullControl, 
            //                                InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit, 
            //                                PropagationFlags.NoPropagateInherit, AccessControlType.Allow ) );
            //dInfo.SetAccessControl( dSecurity );

            using (FileStream streamWrite = File.OpenWrite(filePath))
            {
                try
                {
                    long count = (streamWrite.Length / blockSize) + 1;

                    for (int i = 0; i < count; i++)
                    {
                        await streamWrite.WriteAsync(data, 0, data.Length);
                    }

                    streamWrite.SetLength(0);

                    //if( streamWrite.Length > count * blockSize )
                    //{
                    //    byte[] Last_Data = new byte[ streamWrite.Length - ( count * blockSize ) ];
                    //    await streamWrite.WriteAsync( Last_Data, 0, Last_Data.Length );
                    //}
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                    return false;
                }
                return true;
            }
        }

        // 일반적으로 재귀 방식을 쓰지만 복잡하거나 중첩 규모가 크면 스택 오버 플로우 발생 가능성
        private void TraverseTree(string[] SourceDirs)
        {
            Stack<string> dirs = new Stack<string>();

            // 스택에 소스 경로 넣기( Push )
            foreach (string SourceDir in SourceDirs)
            {
                dirs.Push(SourceDir);
            }

            while (dirs.Count > 0)
            {
                string[] SubDirs = null;

                try
                {
                    string CurrentDir = dirs.Pop();
                    SubDirs = Directory.GetDirectories(CurrentDir);
                    foreach (var SubDir in SubDirs)
                    {
                        // 폴더 일반 속성으로 변경
                        File.SetAttributes(SubDir, FileAttributes.Normal);

                        // 스택에 폴더 넣기( 후입선출 - LIFO )
                        DeleteSubDirs.Push(SubDir.ToString());
                    }
                    string[] files = Directory.GetFiles(CurrentDir);
                    ListView_AddFile(files);
                }
                catch (UnauthorizedAccessException e)
                {
                    Debug.WriteLine(e.Message);
                }

                foreach (string SubDir in SubDirs)
                {
                    dirs.Push(SubDir);
                }
            }
        }

        private async ValueTask<bool> DeleteSubDirectory(Stack<string> SubDirs)
        {
            try
            {
                int DirCount = SubDirs.Count;

                for (int i = 0; i < DirCount; i++)
                {
                    // 전체 경로 스택에서 받기
                    string WorkDirectory = SubDirs.Pop();

                    // 경로만 알아내기
                    string path = Path.GetDirectoryName(WorkDirectory);

                    // 변경할 폴더 이름
                    string NewFolderName = path + "\\tmp";

                    Directory.Move(WorkDirectory, NewFolderName);
                    Directory.Delete(NewFolderName);
                    await Task.Delay(10);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return false;
            }

            return true;
        }

        // 파일 드래그 앤 드랍 
        private void lstView_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] DragDropItems = e.Data.GetData(DataFormats.FileDrop, true) as string[];

                    List<string> DropFolder = new List<string>();
                    List<string> DropFile = new List<string>();

                    // 폴더와 파일을 동시에 드래그 했을 때
                    // 폴더와 파일을 각 각의 리스트에 담기.
                    foreach (var item in DragDropItems)
                    {
                        if (Directory.Exists(item.ToString()))
                        {
                            DropFolder.Add(item.ToString());
                        }
                        else
                        {
                            DropFile.Add(item.ToString());
                        }
                    }

                    if (DropFolder.Count != 0)
                    {
                        TraverseTree(DropFolder.ToArray());
                    }

                    if (DropFile.Count != 0)
                    {
                        ListView_AddFile(DropFile.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ListView_AddFile(string[] files)
        {
            foreach (string file in files)
            {
                // 리스트뷰 아이템 추가
                long fileSize = new FileInfo(file).Length;

                try
                {
                    // 파일 속성 변경
                    File.SetAttributes(file, FileAttributes.Normal);

                    // 리스트 뷰 데이터 추가
                    listView.Items.Add(new SelectionFileData(file, (fileSize / 1024 + 1)));
                }
                catch (Exception)
                {
                    MessageBox.Show("파일이 이미 사용 중 이거나 권한이 없습니다.", "파일 속성 변경 에러", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        // 진행바 갱신
        void ReportProgress(int value)
        {
            progressBar.Value = value;
            TaskbarItemInfo.ProgressValue = (double)value / 100;
        }

        // 파일 이름, 크기별로 정렬 가능
        private void ColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            ColumnHeaderSort(sender);
        }

        private void ColumnHeaderSort(object sender)
        {
            // 클릭한 헤더 Tag 알아내기
            GridViewColumnHeader column = (sender as GridViewColumnHeader);
            string sortBy = column.Tag.ToString();

            // 헤더 클릭 이벤트 발생시 
            if (listViewSortCol != null)
            {
                AdornerLayer.GetAdornerLayer(listViewSortCol).Remove(listViewSortAdorner);
                listView.Items.SortDescriptions.Clear();
            }

            ListSortDirection newDirection = ListSortDirection.Ascending;
            if (listViewSortCol == column && listViewSortAdorner.Direction == newDirection)
            {
                newDirection = ListSortDirection.Descending;
            }

            listViewSortCol = column;
            listViewSortAdorner = new SortAdorner(listViewSortCol, newDirection);
            AdornerLayer.GetAdornerLayer(listViewSortCol).Add(listViewSortAdorner);
            listView.Items.SortDescriptions.Add(new SortDescription(sortBy, newDirection));
        }

        private void RemoveListItem_Click(object sender, RoutedEventArgs e)
        {
            RemoveAllListItem(e);
        }

        private void RemoveAllListItem(RoutedEventArgs e)
        {
            // 단일 아이템
            //listView.Items.Remove( listView.SelectedItem );

            // 여러 아이템 ( 1개 이상의 아이템 선택시 선택된 아이템만 리스트에 추출 )
            try
            {
                List<object> selectedItems = new List<object>();
                foreach (var item in listView.SelectedItems)
                {
                    selectedItems.Add(item);
                }

                // 리스트뷰에서 선택된 파일만 삭제
                foreach (var item in selectedItems)
                {
                    listView.Items.Remove(item);
                }
                button_ZeroFill.IsEnabled = true;
            }
            catch (Exception)
            {
                MessageBox.Show(e.ToString());
            }
        }

        // 컬럼을 정렬했을때 업, 다운 상태 표시 이미지.
        public class SortAdorner : Adorner
        {
            // 삼각형 다운 애로우  
            private static readonly Geometry ascGeometry = Geometry.Parse("M 0 4 L 3.5 0 L 7 4 Z");

            // 삼각형 업 애로우  
            private static readonly Geometry descGeometry = Geometry.Parse("M 0 0 L 3.5 4 L 7 0 Z");

            public ListSortDirection Direction { get; private set; }

            public SortAdorner(UIElement element, ListSortDirection dir) : base(element)
            {
                Direction = dir;
            }

            protected override void OnRender(DrawingContext drawingContext)
            {
                base.OnRender(drawingContext);

                if (AdornedElement.RenderSize.Width < 20)
                {
                    return;
                }

                TranslateTransform transform = new TranslateTransform
                    (AdornedElement.RenderSize.Width - 15,
                    (AdornedElement.RenderSize.Height - 5) / 2);

                drawingContext.PushTransform(transform);

                Geometry geometry = ascGeometry;
                if (Direction == ListSortDirection.Descending)
                {
                    geometry = descGeometry;
                }
                drawingContext.DrawGeometry(Brushes.Black, null, geometry);
                drawingContext.Pop();
            }
        }
    }
}

