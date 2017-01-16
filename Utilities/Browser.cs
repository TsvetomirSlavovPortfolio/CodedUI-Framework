// <copyright file="Browser.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>Browser.cs handles all browser related settings.</summary>
namespace INF.CodedUI.TestAutomation.Utilities
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Reflection;
    using Configuration;
    using Entities;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using UI;

    /// <summary>
    /// Type of Browser.
    /// </summary>
    public class Browser
    {
        /// <summary>
        /// Popup Class Name for web pop up.
        /// </summary>
        private const string PopupClassName = "Internet Explorer";

        /// <summary>
        /// Browser window.
        /// </summary>
        private static readonly Dictionary<string, BrowserWindow> Browsers = new Dictionary<string, BrowserWindow>();

        /// <summary>
        /// Browser window title.
        /// </summary>
        private static readonly Dictionary<string, string> Titles = new Dictionary<string, string>();

        /// <summary>
        ///  Launch Browser.
        /// </summary>
        /// <param name="urlString">URL of the page.</param>
        /// <returns>Browser Window.</returns>
        public static BrowserWindow Launch(string urlString)
        {
            return BrowserWindow.Launch(new Uri(urlString));
        }

        /// <summary>
        ///  Locate a browser with a specific title.
        /// </summary>
        /// <param name="title">Title of the page.</param>
        /// <param name="type">Type of the Browser.</param>
        /// <param name="tryLocateInCashe">Locate elements from cache.</param>
        /// <returns>Browser Title.</returns>
        public static BrowserWindow Locate(string title, string type, bool tryLocateInCashe = true)
        {
            try
            {
                if (tryLocateInCashe && Browsers.ContainsKey(title))
                {
                    return Browsers[title];
                }

                var browserWindow = new BrowserWindow();
                browserWindow.SearchProperties[UITestControl.PropertyNames.Name] = title;

                if (type.ToUpper() == UiControls.TypePopup)
                {
                    browserWindow.SearchProperties[UITestControl.PropertyNames.ClassName] = PopupClassName;
                }

                browserWindow.WindowTitles.Add(title);

                if (tryLocateInCashe)
                {
                    Browsers.Add(title, browserWindow);
                }

                return browserWindow;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Browser, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Locate a browser with a specific title or part of a title.
        /// </summary>
        /// <param name="partOfTitle">Part of Title of the page.</param>
        /// <param name="type">Type of the Browser.</param>
        /// <param name="tryLocateInCashe">Locate elements from cache.</param>
        /// <returns>Partial Browser Title.</returns>
        public static string GetTitleFromPartOfTitle(string partOfTitle, string type, bool tryLocateInCashe = true)
        {
            try
            {
                //// Check if we have search for partoftitle last time, if so return lastTitle instead of search once more -> just to speed up things
                if (tryLocateInCashe && Titles.ContainsKey(partOfTitle))
                {
                    return Titles[partOfTitle];
                }

                var browserWindow = new BrowserWindow();
                browserWindow.SearchProperties.Add(UITestControl.PropertyNames.Name, partOfTitle, PropertyExpressionOperator.Contains);

                if (type.ToUpper() == UiControls.TypePopup)
                {
                    browserWindow.SearchProperties[UITestControl.PropertyNames.ClassName] = PopupClassName;
                }

                if (tryLocateInCashe)
                {
                    Titles.Add(partOfTitle, browserWindow.Name);
                }

                return browserWindow.Name;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Browser, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Close all open Browser Processes that match given BrowserType in configuration file.
        /// </summary>
        /// <returns>Status of browser as True or False.</returns>
        public static bool CloseAllBrowsers()
        {
            try
            {
                switch (General.BrowserType.ToUpper())
                {
                    case Constants.Browsers.IeCaps:
                        foreach (var window in Process.GetProcessesByName(Constants.Browsers.Iexplore))
                        {
                            window.Kill();
                        }

                        break;
                    case Constants.Browsers.FirefoxCaps:
                        foreach (var window in Process.GetProcessesByName(Constants.Browsers.Firefox))
                        {
                            window.Kill();
                        }

                        break;
                    case Constants.Browsers.ChromeCaps:
                        foreach (var window in Process.GetProcessesByName(Constants.Browsers.Chrome))
                        {
                            window.Kill();
                        }

                        break;
                    default:
                        return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Browser, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Clear Browser cookies.
        /// </summary>
        /// <returns>Clearing of Cookie is False or True.</returns>
        public static bool ClearCookies()
        {
            try
            {
                BrowserWindow.ClearCookies();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Browser, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Clear Browser cache.
        /// </summary>
        /// <returns>Returns true or false.</returns>
        public static bool ClearCache()
        {
            try
            {
                BrowserWindow.ClearCache();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Browser, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        ///     Clear Browser cache.
        /// </summary>
        public static void SetCurrentBrowser()
        {
            try
            {
                switch (General.BrowserType.ToUpper())
                {
                    case Constants.Browsers.IeCaps:
                        BrowserWindow.CurrentBrowser = Constants.Browsers.IeCaps;
                        break;
                    case Constants.Browsers.ChromeCaps:
                        BrowserWindow.CurrentBrowser = Constants.Browsers.Chrome;
                        break;
                    case Constants.Browsers.FirefoxCaps:
                        BrowserWindow.CurrentBrowser = Constants.Browsers.Firefox;
                        break;
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Browser, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}