// <copyright file="BaseTestClass.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>BaseTestClass.cs Holds all the basic informations about tests like test initialize and clean up.</summary>
namespace INF.CodedUI.TestAutomation.TestIterations
{
    using System;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using Configuration;
    using Entities;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Utilities;
    
    /// <summary>
    /// Used as Abstract script for other test scripts. Will not contain any test methods. 
    /// Other test scripts will inherit from it so that reusable functions like My Test Initialize, My Test Cleanup can be used.
    /// </summary>
    [TestClass]
    public class BaseTestClass
    {
        /// <summary>
        /// Gets or sets Test Context.
        /// </summary>
        /// <value>Test Context.</value>
        public TestContext TestContext { get; set; }

        /// <summary>
        /// Gets or sets Assembly Initialize.
        /// </summary>
        /// <param name="context">Test Context.</param>
        /// <value>context.</value>
        [AssemblyInitialize]
        public static void AssemblyInitialize(TestContext context)
        {
            try
            {
                //// Initiliaze timing
                Timing.TotalStartTime = DateTime.Now;

                var filePath = Environment.CurrentDirectory.Split('\\');
                filePath =
                    filePath.Where(item => item != filePath[filePath.Length - 1])
                        .Where(item => item != filePath[filePath.Length - 2])
                        .Where(item => item != filePath[filePath.Length - 3])
                        .ToArray();

                var rootPath = string.Join(Constants.DoubleBackslash, filePath);

                TestCase.RootFilePath =
                    string.IsNullOrEmpty(ConfigurationManager.AppSettings.Get(Constants.AppSetting.RootFilePath))
                        ? rootPath
                        : ConfigurationManager.AppSettings.Get(Constants.AppSetting.RootFilePath);

                if (!TestCase.RootFilePath.Last().ToString().Equals(Constants.DoubleBackslash))
                {
                    TestCase.RootFilePath += Constants.DoubleBackslash;
                }

                TestCase.TestReportFileNamePrefix = ConfigurationManager.AppSettings.Get(Constants.AppSetting.FileNamePrefix);
                General.WaitForControlToExistTimeOut =
                    ConfigurationManager.AppSettings.Get(Constants.AppSetting.ExistTimeOut);
                General.BrowserType = ConfigurationManager.AppSettings.Get(Constants.AppSetting.BrowserType);
                Browser.SetCurrentBrowser();

                //// Initiliaze reporting
                var errorMessage = string.Empty;
                if (!Reporting.CreateExcelFile(ref errorMessage))
                {
                    Assert.Inconclusive(Constants.Messages.TestReportError, errorMessage);
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.BaseTestClass, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Assembly Cleanup.
        /// </summary>
        [AssemblyCleanup]
        public static void MyAssemblyCleanup()
        {
            try
            {
                Timing.TotalEndTime = DateTime.Now;
                Timing.Totalduration = Timing.TotalEndTime - Timing.TotalStartTime;
                Reporting.InsertSummaryDetailsAndFormat();
                File.SetAttributes(Reporting.FilePath, FileAttributes.Normal);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.BaseTestClass, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Initialize the test.
        /// </summary>
        [TestInitialize]
        public virtual void TestInitialize()
        {
            try
            {
                //// Initilize timing
                Timing.TestCaseStartTime = DateTime.Now;

                //// Initilize playback
                if (!Playback.IsInitialized)
                {
                    Playback.Initialize();
                }

                //// Initliaze test case
                TestCase.Application = Data.GetValue(TestContext, Constants.TestCase.Application);
                TestCase.Name = Data.GetValue(TestContext, Constants.TestCase.Name);
                TestCase.Description = Data.GetValue(TestContext, Constants.TestCase.Description);
                TestCase.FileName = Data.GetValue(TestContext, Constants.TestCase.FileName);

                //// Initiliaze test step result
                Result.TestStepsResultsCollection.Clear();

                //// Initiliaze test case reporting
                var errorMessage = string.Empty;
                if (!Reporting.CreateExcelSheet(ref errorMessage))
                {
                    Assert.Inconclusive(Constants.Messages.ReportSheetError, errorMessage);
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.BaseTestClass, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        /// Cleaning up the test after iteration.
        /// </summary>
        [TestCleanup]
        public virtual void TestCleanup()
        {
            try
            {
                //// Set end time and duration for test case
                Timing.TestCaseduration = DateTime.Now - Timing.TestCaseStartTime;

                //// Insert information about test case in summary sheet
                Reporting.InsertResultSummary();

                //// Insert information about test step to test case sheet
                var errorMessage = string.Empty;
                if (!Reporting.InsertTestStepResult(ref errorMessage))
                {
                    Assert.Inconclusive(Constants.Messages.ReportInsertError, errorMessage);
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.BaseTestClass, MethodBase.GetCurrentMethod().Name);
                throw;
            }
            finally
            {
                Playback.Cleanup();
            }
        }
    }
}