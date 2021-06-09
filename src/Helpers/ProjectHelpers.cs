using System;
using System.IO;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;

namespace MonikerManifestTools
{
    internal static class ProjectHelpers
    {
        public static void CheckFileOutOfSourceControl(this DTE2 dte, string file)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!File.Exists(file) || dte.Solution.FindProjectItem(file) == null)
            {
                return;
            }

            if (dte.SourceControl.IsItemUnderSCC(file) && !dte.SourceControl.IsItemCheckedOut(file))
            {
                dte.SourceControl.CheckOutItem(file);
            }

            var info = new FileInfo(file)
            {
                IsReadOnly = false
            };
        }

        public static ProjectItem AddFileToProject(this DTE2 dte, Project project, string file, string itemType = null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!File.Exists(file))
            {
                return null;
            }

            ProjectItem item = dte.Solution.FindProjectItem(file);

            if (item == null)
            {
                item = project.ProjectItems.AddFromFile(file);
                item.SetItemType(itemType);
            }

            return item;
        }

        public static void SetItemType(this ProjectItem item, string itemType)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                if (item == null || item.ContainingProject == null)
                {
                    return;
                }

                item.Properties.Item("ItemType").Value = itemType;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Write(ex);
            }
        }
    }
}
