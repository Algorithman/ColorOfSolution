// ---------------------------------------------------------------------------------------
// Solution: ColorOfSolution
// Project: ColorOfSolution
// Filename: ColorOfSolutionPackage.cs
// 
// Last modified: 2024-3-31 15:29
// Created:       2024-3-31 11:28
// 
// Copyright: 2021 Walter Wissing & Co
// ---------------------------------------------------------------------------------------

using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
using Application = System.Windows.Application;
using Project = EnvDTE.Project;

namespace ColorOfSolution
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(UIContextGuids80.NoSolution, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(PackageGuidString)]
    public sealed class ColorOfSolutionPackage : AsyncPackage
    {
        /// <summary>
        ///     ColorOfSolutionPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "fb0d4af9-fcba-490b-b58c-3b6418cfbfc8";

        private static readonly Color[] Backgrounds =
        {
            Color.FromRgb(0, 0, 255), // Blue
            Color.FromRgb(0, 255, 0), // Green
            Color.FromRgb(255, 0, 0), // Red
            Color.FromRgb(255, 255, 0), // Yellow
            Color.FromRgb(255, 0, 255), // Magenta
            Color.FromRgb(0, 255, 255), // Cyan
            Color.FromRgb(128, 0, 0), // Maroon
            Color.FromRgb(0, 128, 0), // Olive
            Color.FromRgb(128, 128, 0), // Brown
            Color.FromRgb(0, 0, 128), // Navy
            Color.FromRgb(128, 0, 128), // Purple
            Color.FromRgb(0, 128, 128), // Teal
            Color.FromRgb(128, 128, 128), // Silver
        };

        private static readonly Color[] Foregrounds =
        {
            Color.FromRgb(0xfa, 0xfa, 0xfa),
            Color.FromRgb(0, 0, 0),
            Color.FromRgb(0xfa, 0xfa, 0xfa),
            Color.FromRgb(0, 0, 0),
            Color.FromRgb(0, 0, 0),
            Color.FromRgb(0, 0, 0),
            Color.FromRgb(0xfa, 0xfa, 0xfa),
            Color.FromRgb(0xfa, 0xfa, 0xfa),
            Color.FromRgb(0xfa, 0xfa, 0xfa),
            Color.FromRgb(0xfa, 0xfa, 0xfa),
            Color.FromRgb(0xfa, 0xfa, 0xfa),
            Color.FromRgb(0xfa, 0xfa, 0xfa),
            Color.FromRgb(0, 0, 0),
        };

        private DTE2 _dte;

        private async void SolutionOpened()
        {
            await SolutionOpenedAsync();
        }


        private async Task UpdateTitleBarColor()
        {
            // Get the current solution's path
            var solutionPath = _dte.Solution.FullName;

            Color foregroundColor, backgroundColor;

            (foregroundColor, backgroundColor) = GetColorForSolution(solutionPath);

            var gg = await FindMainWindowSolutionTitleBar();

            gg.SetColors(new SolidColorBrush(foregroundColor), new SolidColorBrush(backgroundColor));
        }


        private async Task SolutionOpenedAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var mainWindow = Application.Current.MainWindow;
            if (mainWindow != null)
            {
                while (true)
                {
                    await Task.Delay(150);
                    var mainWindowTitleBar = await FindMainWindowSolutionTitleBar();
                    if (mainWindowTitleBar != null)
                    {
                        await UpdateTitleBarColor();
                        break;
                    }
                }
            }
        }


        private async Task<FrameworkElement> FindMainWindowSolutionTitleBar()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var mainWindow = Application.Current.MainWindow;
            if (mainWindow != null)
            {
                var mainWindowTitleBar = mainWindow.FindChild<FrameworkElement>("PART_SolutionNameTextBlock");
                return mainWindowTitleBar;
            }

            throw new Exception("Didn't find PART_SolutionNameTextBlock");
        }

        private (Color, Color) GetColorForSolution(string solutionPath)
        {
            var propertyForeground = GetProjectProperty(GetStartupProject(solutionPath), "Foreground");
            var propertyBackground = GetProjectProperty(GetStartupProject(solutionPath), "Background");

            if (!string.IsNullOrEmpty(propertyForeground) && !string.IsNullOrEmpty(propertyBackground) && propertyForeground.Length >= 7 &&
                propertyBackground.Length >= 7 && propertyForeground.Length <= 9 && propertyBackground.Length <= 9)
            {
                var foregroundBrush = StringToSolidColorBrush(propertyForeground);
                var backgroundBrush = StringToSolidColorBrush(propertyBackground);
                return (foregroundBrush, backgroundBrush);
            }


            // Calculate a hash code from the solution path
            var hashCode = solutionPath.GetHashCode();

            // Map the hash code to an index within the VGA color array
            var index = Math.Abs(hashCode) % Backgrounds.Length;

            // Return the color at the calculated index
            return (Foregrounds[index], Backgrounds[index]);
        }


        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Get an instance of the DTE
            _dte = await GetServiceAsync(typeof(DTE)) as DTE2;

            // Subscribe to solution events
            _dte.Events.SolutionEvents.Opened += SolutionOpened;

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
        }

        private string GetProjectProperty(Project project, string PropName)
        {
            var userFilePath = Path.ChangeExtension(project.FullName, ".csproj.user");

            if (File.Exists(userFilePath))
            {
                // Load the .csproj.user file
                var doc = new XmlDocument();
                doc.Load(userFilePath);

                // Define the namespace manager
                var nsManager = new XmlNamespaceManager(doc.NameTable);
                nsManager.AddNamespace("msb", "http://schemas.microsoft.com/developer/msbuild/2003");

                // Find the PropertyGroup element containing the Background and Foreground properties
                var propertyGroups = doc.GetElementsByTagName("PropertyGroup");
                foreach (XmlNode propertyGroup in propertyGroups)
                {
                    var resultNode = propertyGroup.SelectSingleNode("msb:" + PropName, nsManager);

                    if (resultNode != null)
                    {
                        var result = resultNode.InnerText;
                        return result;
                    }
                }
            }

            return null;
        }

        private Project GetStartupProject(string solutionFilePath)
        {
            // Get the DTE service to access solution and projects
            var dte = GetGlobalService(typeof(SDTE)) as DTE;

            if (dte != null && dte.Solution != null && dte.Solution.Projects.Count > 0)
            {
                // Iterate over the projects in the solution
                foreach (Project project in dte.Solution.Projects)
                {
                    // Check if the project is marked as the startup project
                    if (IsStartupProject(_dte.Solution, project))
                    {
                        // Check if the project is a C# project (you may need additional checks based on your solution's structure)
                        if (Path.GetExtension(project.FileName).Equals(".csproj", StringComparison.OrdinalIgnoreCase))
                        {
                            return project;
                        }
                    }
                }
            }

            return null;
        }

        private bool IsStartupProject(Solution solution, Project project)
        {
            // Check if the project is marked as the startup project
            foreach (string startupProjectUniqueName in (Array)solution.SolutionBuild.StartupProjects)
            {
                if (startupProjectUniqueName == project.UniqueName)
                {
                    return true;
                }
            }

            return false;
        }

        public static Color StringToSolidColorBrush(string colorString)
        {
            if (string.IsNullOrEmpty(colorString) || colorString.Length < 7) // Minimum length of color string is '#rrggbb'
            {
                // Invalid color string, return null or a default color
                return colorString == "Foreground" ? Color.FromRgb(250, 250, 250) : Color.FromRgb(31, 31, 31);
            }

            try
            {
                byte a = 255; // Default alpha value
                if (colorString.Length == 9) // If color string includes alpha channel
                {
                    a = byte.Parse(colorString.Substring(1, 2), NumberStyles.HexNumber);
                    colorString = "#" + colorString.Substring(3); // Remove alpha channel from color string
                }

                var r = byte.Parse(colorString.Substring(1, 2), NumberStyles.HexNumber);
                var g = byte.Parse(colorString.Substring(3, 2), NumberStyles.HexNumber);
                var b = byte.Parse(colorString.Substring(5, 2), NumberStyles.HexNumber);

                return Color.FromArgb(a, r, g, b);
            }
            catch (Exception)
            {
                // Error parsing color string, return null or a default color
                return colorString == "Foreground" ? Color.FromRgb(250, 250, 250) : Color.FromRgb(31, 31, 31);
            }
        }
    }

    public static class VisualTreeHelperExtensions
    {
        public static T FindChild<T>(this DependencyObject parent, string childName) where T : FrameworkElement
        {
            if (parent == null)
            {
                return null;
            }

            var childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (var i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is T frameworkElement && frameworkElement.Name == childName)
                {
                    return frameworkElement;
                }

                var result = FindChild<T>(child, childName);
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }

        private static T FindFirstChild<T>(this DependencyObject parent) where T : FrameworkElement
        {
            if (parent == null)
            {
                return null;
            }

            var childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (var i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);

                if (child is T frameworkElement)
                {
                    return frameworkElement;
                }

                var result = FindFirstChild<T>(child);
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }

        public static void SetColors(this FrameworkElement element, Brush FG, Brush BG)
        {
            // Check if the element has the Foreground and Background properties
            var FGE = element.FindFirstChild<TextBlock>();
            var BGE = element.GetParent();
            var foregroundProperty = FGE.GetType().GetProperty("Foreground");
            var backgroundProperty = BGE.GetType().GetProperty("Background");

            if (foregroundProperty != null && foregroundProperty.CanWrite)
            {
                // Set the Foreground property
                foregroundProperty.SetValue(FGE, FG);
            }

            if (backgroundProperty != null && backgroundProperty.CanWrite)
            {
                // Set the Background property
                backgroundProperty.SetValue(BGE, BG);
            }
        }

        private static DependencyObject GetParent(this DependencyObject child)
        {
            if (child == null)
            {
                return null;
            }

            return VisualTreeHelper.GetParent(child);
        }
    }
}