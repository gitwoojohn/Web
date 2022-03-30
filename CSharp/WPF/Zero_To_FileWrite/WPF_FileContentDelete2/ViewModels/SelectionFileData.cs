using System.Collections.ObjectModel;

namespace WPF_FileContentDelete.ViewModels
{
    public class ListViewDataSource : ObservableCollection<SelectionFileData> { }

    public class SelectionFileData
    {
        public string FilePath { get; }

        public long FileSize { get; }

        public override string ToString()
        {
            return FilePath;
        }

        public SelectionFileData( string filePath, long fileSize )
        {
            FilePath = filePath;
            FileSize = fileSize;
        }
    }
}
