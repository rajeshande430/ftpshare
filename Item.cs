using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FTP_Share
{
    public class Item
    {
        public string Name { get; set; }
        public string FullPath { get; set; }
        public string RelativePath { get; set; }
        public ItemType Type { get; set; }
        public Folder Folder { get; set; }

        public List<Item> Items { get; set; } = new List<Item>();

    }

    public enum ItemType
    {
        Folder,
        File
    }
}
