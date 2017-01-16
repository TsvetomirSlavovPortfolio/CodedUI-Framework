// <copyright file="WorkBookUtility.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>WorkBookUtility.cs Work book related handles.</summary>
namespace INF.CodedUI.TestAutomation.Utilities
{
    using System;
    using Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Work book utility.
    /// </summary>
    public static class WorkBookUtility
    {
        /// <summary>
        /// Open work book.
        /// </summary>
        /// <param name="applicationClass">Read application class.</param>
        /// <param name="fileName">File name.</param>
        /// <returns>Work book.</returns>
        public static Workbook OpenWorkBook(ApplicationClass applicationClass, string fileName)
        {
            var workbook = applicationClass.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            return workbook;
        }
        
        /// <summary>
        /// Close the work book.
        /// </summary>
        /// <param name="workbook">Work book name.</param>
        /// <param name="save">Save the work book before close.</param>
        /// <param name="message">Message to be stored.</param>
        public static void CloseWorkBook(Workbook workbook, bool save = false, string message = "")
        {
            //// Close workbook
            workbook.Close(save);
            General.ReleaseObject(workbook);
        }

        /// <summary>
        /// Close Excel.
        /// </summary>
        /// <param name="applicationClass">Application Class.</param>
        /// <param name="message">Message to be saved.</param>
        public static void CloseExcel(ApplicationClass applicationClass, string message = "")
        {
            applicationClass.Quit();
            General.ReleaseObject(applicationClass);
        }
    }
}