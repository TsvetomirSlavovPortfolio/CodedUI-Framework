// <copyright file="Window.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>Window.cs class handles windows</summary>
namespace INF.CodedUI.TestAutomation.Utilities
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Reflection;
    using Configuration;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using UI;

    /// <summary>
    /// Window handles.
    /// </summary>
    public class Window
    {
        /// <summary>
        /// Pop up class name.
        /// </summary>
        private const string PopupClassname = "Internet Explorer";

        /// <summary>
        /// Dictionary value.
        /// </summary>
        //// key = Part of title, value = title
        private static readonly Dictionary<string, string> Titles = new Dictionary<string, string>();

        /// <summary>
        /// Dictionary value for windows.
        /// </summary>
        //// key = title, value = a ApplicationUnderTest
        private static readonly Dictionary<string, ApplicationUnderTest> Windows =
            new Dictionary<string, ApplicationUnderTest>();

        /// <summary>
        /// Gets or sets Dictionary value for application under test.
        /// </summary>
        /// <value>Application name.</value>
        public static ApplicationUnderTest Bw = new ApplicationUnderTest();

        /// <summary>
        /// Launch Window.
        /// </summary>
        /// <param name="urlString">Path for the application.</param>
        /// <returns>Application path.</returns>
        public static ApplicationUnderTest Launch(string urlString)
        {
            return ApplicationUnderTest.Launch(urlString);
        }

        /// <summary>
        /// Locate a Window with a specific title.
        /// </summary>
        /// <param name="title">Window title.</param>
        /// <param name="type">Type of Window.</param>
        /// <param name="tryLocateInCashe">Locate window property from cache.</param>
        /// <returns>Title of Window.</returns>
        public static ApplicationUnderTest Locate(string title, string type, bool tryLocateInCashe = false)
        {
            try
            {
                if (tryLocateInCashe && Windows.ContainsKey(title))
                {
                    return Windows[title];
                }

                Bw.SearchProperties[UITestControl.PropertyNames.Name] = title;

                if (type == UiControls.TypePopup)
                {
                    Bw.SearchProperties[UITestControl.PropertyNames.ClassName] = PopupClassname;
                }

                Bw.WindowTitles.Add(title);

                if (tryLocateInCashe)
                {
                    Windows.Add(title, Bw);
                }

                return Bw;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Window, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Locate a Window with a specific title or part of a title.
        /// </summary>
        /// <param name="partoftitle">Part of title.</param>
        /// <param name="type">Window type.</param>
        /// <param name="tryLocateInCashe">Locate window properties from cache.</param>
        /// <returns>Window title.</returns>
        public static string GetTitleFromPartOfTitle(string partoftitle, string type, bool tryLocateInCashe = false)
        {
            try
            {
                //// Check if we have search for partoftitle last time, if so return lastTitle instead of search once more -> just to speed up things
                var bw = new ApplicationUnderTest();

                if (tryLocateInCashe && Titles.ContainsKey(partoftitle))
                {
                    return Titles[partoftitle];
                }

                bw.SearchProperties.Add(UITestControl.PropertyNames.Name, partoftitle, PropertyExpressionOperator.Contains);

                if (type == UiControls.TypePopup)
                {
                    bw.SearchProperties[UITestControl.PropertyNames.ClassName] = PopupClassname;
                }

                if (tryLocateInCashe)
                {
                    Titles.Add(partoftitle, bw.Name);
                }

                return bw.Name;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Window, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Close all open Window Processes that match given WindowType in configuration file.
        /// </summary>
        /// <returns>Window is closed or not.</returns>
        public static bool CloseAllWindows()
        {
            try
            {
                switch (General.WindowType.ToUpper())
                {
                    case "IE":
                        var appWindows = Process.GetProcessesByName("windowapp");
                        foreach (var window in appWindows)
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
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.Window, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }
    }
}