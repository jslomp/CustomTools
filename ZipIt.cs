using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
// work in progress here 
/*
 * Author: Jacob Slomp
 */

namespace CustomTools
{
    class ZipIt
    {
        private string readDir;

        string theDirName = "";
        public ZipIt(string readDir)
        {

            theDirName = Path.GetFullPath(readDir);
            string[] dirs = theDirName.Split(new string[] { @"\" }, StringSplitOptions.None);

            theDirName = dirs[dirs.Length - 1];
            

            this.readDir = readDir;
            deepRead(readDir);
        }

        private void deepRead(string dir)
        {
            string[] dirs = Directory.GetDirectories(dir);
            foreach(string d in dirs)
            {
                deepRead(d);
            }
            string[] f = Directory.GetFiles(dir);
            foreach(string f1 in f) {
                if (!files.Contains(f1))
                {
                    files.Add(f1);
                }
            }

        }
        List<string> files = new List<string>();

        private void createZip(string FileName)
        {

            if (File.Exists(FileName))
            {
                try
                {
                    File.Delete(FileName);
                }
                catch (IOException e)
                {
                    Console.WriteLine(e.Message);

                    File.Delete(FileName);
                }
            }

            // Create FileStream for output ZIP archive
            using (var fileStream = new FileStream(FileName, FileMode.CreateNew))
            {
                using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Create, true))
                {

                    foreach (string f in files)
                    {
                        var fileBytes = File.ReadAllBytes(f);
                        string name = f;
                        name = name.Replace(theDirName + @"\", "");

                        var zipArchiveEntry = archive.CreateEntry(name, CompressionLevel.Optimal);
                        using (var zipStream = zipArchiveEntry.Open())
                        {
                            zipStream.Write(fileBytes, 0, fileBytes.Length);

                        }

                    }

                }
            }
        }


        public void saveAs(string filename, string dialogDefaultFileName = "", string dialogDefaultFilter = "*.zip|*.zip")
        {

            if(filename == "" || filename == null)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.FileName = dialogDefaultFileName;
                sfd.Filter = dialogDefaultFilter;
                var result = sfd.ShowDialog();
                if(result == DialogResult.OK)
                {
                    filename = sfd.FileName;
                }
            }
            if (filename != "" && filename != null)
            {
                deepRead(readDir);
                createZip(filename);
            }
        }
    }
}
