// <copyright file="TestCases.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>TestCases.cs class handles all test cases and its steps as per test iteration excel.</summary>
namespace INF.CodedUI.TestAutomation.TestCases
{
    using System;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using System.Windows.Forms;
    using Configuration;
    using Entities;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using UI;
    using Utilities;

    /// <summary>
    /// Reads Test cases and executes.
    /// </summary>
    public class TestCases
    {
        /// <summary>
        /// Gets or sets Hyper link for the failed test steps.
        /// </summary>
        /// <value>File path.</value>
        public static string Hyplink { get; set; }

        /// <summary>
        /// Gets or sets Test Session ID.
        /// </summary>
        /// <value>ID of test session.</value>
        public static Guid TestSessionId { get; set; }

        /// <summary>
        /// Executes test steps.
        /// </summary>
        public static void Execute()
        {
            try
            {
                TestSessionId = Guid.NewGuid();

                //// Initiliaze test data
                TestCase.TestDataCount = Constants.Zero;

                var errorMessage = string.Empty;
                if (!Data.InitiliazeTestCaseAndTestData(ref errorMessage))
                {
                    Assert.Inconclusive(Constants.Messages.TestInitializationError, errorMessage);
                }

                //// Run the test case once for every test data that exist
                var testCaseCount = TestCase.TestDataCount;
                for (var testCaseIndex = 1; testCaseIndex <= testCaseCount; testCaseIndex++)
                {
                    var isSuccess = true;

                    var testStepLength = TestCase.TestStepList.Count;
                    for (var testStepCount = 0; testStepCount <= testStepLength - 1; testStepCount++)
                    {
                        var teststep = new TestStep();
                        var testStep = teststep;
                        testStep.TestDataKeyToUse = Convert.ToString(testCaseIndex);
                        testStep.Action = TestCase.TestStepList[testStepCount].Action;
                        testStep.TestData = TestCase.TestStepList[testStepCount].TestData;
                        testStep.TestStepNumber = TestCase.TestStepList[testStepCount].TestStepNumber;
                        testStep.UiControl = TestCase.TestStepList[testStepCount].UiControl;
                        testStep.Verification = TestCase.TestStepList[testStepCount].Verification;
                        testStep.Remarks = TestCase.TestStepList[testStepCount].Remarks;

                        switch (testStep.Action.ToUpper())
                        {
                            case Constants.TestStepAction.CloseBrowsers:
                                isSuccess = UiActions.CloseAllBrowsers(testStep);
                                break;
                            case Constants.TestStepAction.ClearBrowserCookies:
                                isSuccess = UiActions.ClearBrowserCookies(testStep);
                                break;
                            case Constants.TestStepAction.ClearBrowserCashe:
                                isSuccess = UiActions.ClearBrowserCache(testStep);
                                break;
                            case Constants.TestStepAction.LaunchBrowser:
                                isSuccess = UiActions.LaunchBrowser(testStep);
                                break;
                            case Constants.TestStepAction.LaunchWindow:
                                isSuccess = UiActions.LaunchWindow(testStep);
                                break;
                            case Constants.TestStepAction.EditUiControl:
                                isSuccess = UiActions.EditUiControl(testStep);
                                break;
                            case Constants.TestStepAction.Verify:
                                isSuccess = UiActions.Verify(testStep, true);
                                break;
                            case Constants.TestStepAction.WaitForUi:
                                isSuccess = UiActions.WaitForUi(testStep);
                                break;
                            case Constants.TestStepAction.SendKeys:
                                isSuccess = UiActions.SendKeys(testStep);
                                break;
                            case Constants.TestStepAction.SaveUiControlValue:
                                isSuccess = UiActions.SaveUiControlValue(testStep);
                                break;
                            default:
                                Result.PassStepOutandSave(
                                    testStep.TestDataKeyToUse,
                                    testStep.TestStepNumber,
                                    Constants.TestIterations,
                                    Constants.Fail,
                                    string.Format(Constants.Messages.NotSupported, testStep.Action),
                                    testStep.Remarks);

                                isSuccess = false;
                                break;
                        }

                        if (!isSuccess)
                        {
                            break;
                        }

                        //// Exit if one test step failed
                    }

                    if (!isSuccess && Result.GetTestScriptResult() == Constants.Fail)
                    {
                        //// Some test step failed, raise Assert.Fail after all test data iterations are completed
                        var bounds = Screen.GetBounds(Point.Empty);

                        using (var bitmap = new Bitmap(bounds.Width, bounds.Height))
                        {
                            using (var objGraphic = Graphics.FromImage(bitmap))
                            {
                                objGraphic.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
                            }

                            bitmap.Save(
                                new StringBuilder()
                                    .Append(Reporting.PathString)
                                    .Append(Constants.DoubleBackslash)
                                    .Append(Path.GetFileNameWithoutExtension(TestCase.Name))
                                    .Append(Constants.Jpg).ToString(),
                                ImageFormat.Jpeg);

                            Hyplink = new StringBuilder()
                                .Append(Constants.Hyperlink)
                                .Append(Reporting.PathString)
                                .Append(Constants.DoubleBackslash)
                                .Append(Path.GetFileNameWithoutExtension(TestCase.Name))
                                .Append(Constants.Jpg)
                                .Append(@""",""")
                                .Append(Path.GetFileNameWithoutExtension(TestCase.Name))
                                .Append(@""")").ToString();
                        }

                        Assert.Fail(Constants.Messages.TestCaseFailedError, Reporting.FilePath);
                    }
                }

                LogHelper.EventLog(
                    string.Format(
                        Constants.Messages.SuccessfullCompletion,
                        MethodBase.GetCurrentMethod().Name),
                        Constants.ClassName.TestCases,
                        MethodBase.GetCurrentMethod().Name);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.TestCases, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}