using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DutiesSendService
{
    public class DirectoryManager
    {
        public string CreateTemporaryDutiesDirectory()
        {
            var dirPath = $"{Path.GetTempPath()}Duties\\";
            return Directory.CreateDirectory(dirPath).FullName;
        }

        public void RemoveTemporaryDutiesDirectory(string directory)
        {
            Directory.Delete(directory, true);
        }
    }
}
