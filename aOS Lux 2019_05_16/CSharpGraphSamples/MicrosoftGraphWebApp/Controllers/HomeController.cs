using MicrosoftGraphWebApp.Helpers;
using MicrosoftGraphWebApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace MicrosoftGraphWebApp.Controllers
{
    public class HomeController : BaseController
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Error(string message, string debug)
        {
            Flash(message, debug);
            return RedirectToAction("Index");
        }

        public async Task<ActionResult> MyEmails()
        {
            var myEmails = await GraphHelper.GetMyEmails();
            return View(myEmails);
        }

        public async Task<ActionResult> MyFiles(string path=null)
        {
            var myFiles = await GraphHelper.GetMyFiles(path);
            var model = new FilesBrowsingViewModel()
            {
                CurrentPath = path,
                PathSegments = GetSegmentsFromPath(path),
                Files = myFiles.ToList()
            };

            return View(model);
        }

        private PathSegment[] GetSegmentsFromPath(string path)
        {
            List<PathSegment> segments = new List<PathSegment>
            {

                // Add root
                new PathSegment()
                {
                    FullPath = "",
                    Segment = "Root"
                }
            };

            if (!string.IsNullOrEmpty(path))
            {
                string[] segmentsAsStrings = path.Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                var currentPathBuilder = new StringBuilder();
                for (int i = 0; i < segmentsAsStrings.Length; i++)
                {
                    string currentSegment = segmentsAsStrings[i];
                    currentPathBuilder.Append($"/{currentSegment}");
                    segments.Add(new PathSegment()
                    {
                        FullPath = currentPathBuilder.ToString(),
                        Segment = currentSegment
                    });
                }
            }

            return segments.ToArray();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}