using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace fr.avh.archivage
{
    public class Utils
    {

        public delegate void OnErrorCallback(Exception e);

        public delegate void OnInfoCallback(string info, Tuple<int, int> progressIndicator = null);

        public enum SearchType
        {
            Files,
            Directories,
            All
        }

        public static string LongPath(string path)
        {
            return (path.Length > 259 ? "\\\\?\\" : "") + path;
        }
        public static List<FileInfo> LazyRecusiveSearch(
            string path,
            Regex searchedPattern,
            SearchType searchingType = SearchType.All,
            OnInfoCallback onInfo = null
        )
        {
            List<FileInfo> result = new List<FileInfo>();
            Match searchedStructure = searchedPattern.Match(Path.GetFileName(path));
            if (
                searchedStructure.Success
                && (
                    searchingType == SearchType.All
                    || (searchingType == SearchType.Directories && Directory.Exists(LongPath(path)))
                    || (searchingType == SearchType.Files && File.Exists(LongPath(path)))
                )
            )
            {
                result.Add(new FileInfo(LongPath(path)));
            }
            if (Directory.Exists(LongPath(path)) && !(searchedStructure.Success && searchingType == SearchType.Directories))
            {
                onInfo?.Invoke("Recherche dans le dossier " + path);

                string[] subtree = Directory.GetDirectories(
                    LongPath(path)
                );
                foreach (string dir in subtree)
                {
                    result.AddRange(
                        LazyRecusiveSearch(dir, searchedPattern, searchingType, onInfo)
                    );
                }

                if (searchingType == SearchType.All || searchingType == SearchType.Files)
                {
                    string[] files = Directory.GetFiles(
                        LongPath(path)
                    );
                    foreach (string dir in files)
                    {
                        result.AddRange(
                            LazyRecusiveSearch(dir, searchedPattern, searchingType, onInfo)
                        );
                    }
                }
            }
            return result;
        }

    }
}
