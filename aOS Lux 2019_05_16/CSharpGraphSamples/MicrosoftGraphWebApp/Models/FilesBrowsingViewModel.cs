using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MicrosoftGraphWebApp.Models
{

    public class PathSegment
    {
        public string Segment { get; set; }
        public string FullPath { get; set; }
    }

    public class FilesBrowsingViewModel
    {
        public string CurrentPath { get; set; }
        public PathSegment[] PathSegments { get; set; }
        public IList<DriveItem> Files { get; set; }
    }
}