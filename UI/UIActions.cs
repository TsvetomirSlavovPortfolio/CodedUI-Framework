// <copyright file="UIActions.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>UIActions.cs Reads and performs the User Interface action needs to be done while under test.</summary>
namespace INF.CodedUI.TestAutomation.UI
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using System.Windows.Forms;
    using System.Windows.Input;
    using Configuration;
    using Entities;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;   
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Utilities;
    using Verifications;
    using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
    using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
    using Point = System.Drawing.Point;

    /// <summary>
    /// User interface actions.
    /// </summary>
    public class UiActions
    {
        /// <summary>
        /// Gets or sets Application class.
        /// </summary>
        /// <value>Application class as value.</value>
        public static ApplicationClass ApplicationClass = new ApplicationClass();

        /// <summary>
        /// Gets or sets Workbook as null.
        /// </summary>
        /// <value>Workbook value.</value>
        public static Workbook Workbook = null;

        /// <summary>
        /// Gets or sets workbook header value.
        /// </summary>
        /// <value>Header Value.</value>
        public static string HeaderValue = string.Empty;

        /// <summary>
        /// Gets or sets Data value.
        /// </summary>
        /// <value>Data Value.</value>
        public static string DataValue = string.Empty;

        /// <summary>
        /// Gets or sets value attribute.
        /// </summary>
        /// <value>Value Attribute.</value>
        public static string ValueAttribute = string.Empty;

        /// <summary>
        /// Gets or sets Html control.
        /// </summary>
        /// <value>Html control.</value>
        public static dynamic HtmlControl { get; set; }

        /// <summary>
        /// Close all browsers.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool CloseAllBrowsers(TestStep testStep)
        {
            try
            {
                Assert.IsTrue(Browser.CloseAllBrowsers(), "All browsers couldn't be closed");
                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "CloseAllBrowsers",
                    Entities.Constants.Pass,
                    "All browsers closed sucessfully",
                    testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "CloseAllBrowsers", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Clean browser cookies.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool ClearBrowserCookies(TestStep testStep)
        {
            try
            {
                Assert.IsTrue(Browser.ClearCookies(), "Browser cookies couldn't be cleared");
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "ClearBrowserCookies", Entities.Constants.Pass, "Browser cookies was cleared sucessfully", testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "ClearBrowserCookies", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Clean browser Cache.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool ClearBrowserCache(TestStep testStep)
        {
            try
            {
                Assert.IsTrue(Browser.ClearCache(), "Browser cache couldn't be cleared");
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "ClearBrowserCache", Entities.Constants.Pass, "Browser cache was cleared sucessfully", testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "ClearBrowserCache", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Launch browser.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool LaunchBrowser(TestStep testStep)
        {
            try
            {
                Browser.Launch(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]));
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "LaunchBrowser", Entities.Constants.Pass, "Browser launched sucessfully", testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "LaunchBrowser", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Launch browser window.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool LaunchWindow(TestStep testStep)
        {
            try
            {
                Utilities.Window.Launch(Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]));
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "LaunchWindow", Entities.Constants.Pass, "Window launched sucessfully", testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(testStep.TestDataKeyToUse, testStep.TestStepNumber, "LaunchWindow", Entities.Constants.Fail, string.Format(Entities.Constants.Messages.DueToException, ex.Message), testStep.Remarks);
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Edit user interface control.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool EditUiControl(TestStep testStep)
        {
            try
            {
                if (testStep.UiControl == null)
                {
                    throw new Exception("UiControl couldn't be found.");
                }

                //// Get testdata value, and check if UI control is radiobutton because then it's a special case
                var valueToEdit = Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]);
                if (testStep.TestData != null)
                {
                    if (!string.IsNullOrEmpty(valueToEdit) &&
                        (testStep.UiControl.UiControlType.ToUpper() == UiControls.HtmlRadioButton ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.HtmlInputButton ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.HtmlButton ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.HtmlCustom ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.HtmlDiv ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.HtmlImage ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinTree ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinTreeItem ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinButton ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinRadioButton ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinCheckBox ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinCheckBoxTreeItem ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinCalendar ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinCell ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinClient ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinComboBox ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinCustom ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinDateTimePicker ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinGroup ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinHyperlink ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinList ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinListItem ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinMenu ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinMenuBar ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinPane ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinProgressBar ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinRow ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinRowHeader ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinScrollBar ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinStatusBar ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinTable ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinTabList ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinTabPage ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinText ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinTitleBar ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinToolBar ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinToolTip ||
                         testStep.UiControl.UiControlType.ToUpper() == UiControls.WinWindow))
                    {
                        if (testStep.Verification != null)
                        {
                            if (valueToEdit.ToUpper() != Entities.Constants.UiActions.LeftClick &&
                                valueToEdit.ToUpper() != Entities.Constants.UiActions.RightClick &&
                                valueToEdit.ToUpper() != Entities.Constants.UiActions.MouseHover)
                            {
                                testStep.UiControl.UiControlSearchValue = valueToEdit;
                            }
                        }
                    }
                }

                //// Create htmlcontrol
                dynamic htmlControl = UiControls.CreateControl(
                    testStep.UiControl.UiControlType,
                    testStep.UiControl.UiTitle,
                    testStep.UiControl.UiType,
                    testStep.UiControl.UiControlSearchProperty,
                    testStep.UiControl.UiControlSearchValue,
                    testStep.UiControl.UiControlFilterProperty,
                    testStep.UiControl.UiControlFilterValue);

                //// Edit
                var failMessage = "UiControl with search property " + testStep.UiControl.UiControlSearchProperty +
                                  " and search value " + testStep.UiControl.UiControlSearchValue + " and filter property " +
                                  testStep.UiControl.UiControlFilterProperty + " and filter value " +
                                  testStep.UiControl.UiControlFilterValue + " is not enabled or does not exist";
                if (valueToEdit != null && valueToEdit.ToUpper() == "{LEFTCLICK}")
                {
                    Mouse.Click(htmlControl, MouseButtons.Left);
                }
                else if (valueToEdit != null && valueToEdit.ToUpper() == "{RIGHTCLICK}")
                {
                    Mouse.Click(htmlControl, MouseButtons.Right);
                }
                else if (valueToEdit != null && valueToEdit.ToUpper() == "{MOUSEHOVER}")
                {
                    Mouse.Hover(htmlControl);
                }
                else if (valueToEdit != null && valueToEdit.ToUpper() == "{DBLCLICK}")
                {
                    Mouse.DoubleClick(htmlControl, MouseButtons.Middle);
                }
                else
                {
                    switch (testStep.UiControl.UiControlType.ToUpper())
                    {
                        case UiControls.HtmlInputButton:
                        case UiControls.HtmlButton:
                            Assert.IsTrue(htmlControl.WaitForControlExist(Convert.ToInt32(General.WaitForControlToExistTimeOut)));
                            Mouse.Click(htmlControl);
                            break;
                        case UiControls.HtmlDiv:
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.EnterBracket, ModifierKeys.None);
                            break;
                        case UiControls.HtmlCustom:
                            Mouse.Click(htmlControl);
                            break;
                        case UiControls.HtmlRadioButton:
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.SpaceBracket, ModifierKeys.None);
                            break;
                        case UiControls.HtmlImage:
                            if (valueToEdit == null)
                            {
                                Mouse.Click(htmlControl);
                            }
                            else if (valueToEdit.ToUpper().Contains("Y") & valueToEdit.ToUpper().Contains("X"))
                            {
                                //// There are some coordinates to click on
                                var relativeCoordinate = new Point();
                                dynamic indexOfY = valueToEdit.IndexOf("Y", StringComparison.CurrentCulture);
                                dynamic indexOfX = valueToEdit.IndexOf("X", StringComparison.CurrentCulture);
                                if (indexOfY < indexOfX)
                                {
                                    relativeCoordinate.Y = valueToEdit.Substring(indexOfY + 1, indexOfX - 1);
                                    relativeCoordinate.X = valueToEdit.Substring(indexOfX + 1,
                                        valueToEdit.Length - indexOfX - 1);
                                }
                                else
                                {
                                    relativeCoordinate.X = valueToEdit.Substring(indexOfX + 1, indexOfY - 1);
                                    relativeCoordinate.Y = valueToEdit.Substring(indexOfY + 1,
                                        valueToEdit.Length - indexOfY - 1);
                                }

                                Mouse.Click(htmlControl, relativeCoordinate);
                            }
                            else
                            {
                                throw new Exception("Wrong kind of test data entered for UiControl of type " +
                                                    testStep.UiControl.UiControlType + " in test step " +
                                                    testStep.TestStepNumber + ".");
                            }

                            break;
                        case UiControls.HtmlComboBox:
                            Assert.IsTrue(
                                !string.IsNullOrEmpty(valueToEdit),
                                string.Format(Entities.Constants.Messages.UiControlError, testStep.UiControl.UiControlType, testStep.TestStepNumber));

                            var foundItem = false;
                            for (var item = 1; item <= htmlControl.ItemCount; item++)
                            {
                                //// Use below line to make sure that some events will happen that show/hide other fields dependent on the list value
                                Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.DownBracket, ModifierKeys.None);

                                //// Remove any dots or comma because sometimes they can be difficult to compare because of format issues
                                string displayText =
                                    htmlControl.SelectedItem.Replace(Entities.Constants.FullStop, string.Empty)
                                        .Replace(Entities.Constants.Comma, string.Empty)
                                        .Replace(Entities.Constants.Hyphen, string.Empty)
                                        .Replace(Entities.Constants.Space, string.Empty)
                                        .Replace(Entities.Constants.Percentage, string.Empty)
                                        .Replace(Entities.Constants.Asteric, string.Empty)
                                        .Replace(Entities.Constants.Equal, string.Empty)
                                        .Replace(Entities.Constants.Colon, string.Empty)
                                        .ToUpper();
                                var value =
                                    valueToEdit.Replace(Entities.Constants.FullStop, string.Empty)
                                        .Replace(Entities.Constants.Comma, string.Empty)
                                        .Replace(Entities.Constants.Hyphen, string.Empty)
                                        .Replace(Entities.Constants.Space, string.Empty)
                                        .Replace(Entities.Constants.Space, string.Empty)
                                        .Replace(Entities.Constants.Percentage, string.Empty)
                                        .Replace(Entities.Constants.Asteric, string.Empty)
                                        .Replace(Entities.Constants.Equal, string.Empty)
                                        .Replace(Entities.Constants.Colon, string.Empty)
                                        .ToUpper();
                                if (displayText == value)
                                {
                                    foundItem = true;
                                    break; //// TODO: might not be correct. Was : Exit For
                                }
                            }

                            Assert.IsTrue(foundItem, failMessage);
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.TabBracket, ModifierKeys.None);
                            break;
                        case UiControls.HtmlComboBoxSelectDirect:
                            Assert.IsTrue(
                                !string.IsNullOrEmpty(valueToEdit),
                                string.Format(Entities.Constants.Messages.UiControlError, testStep.UiControl.UiControlType, testStep.TestStepNumber));

                            Keyboard.SendKeys(htmlControl, "T", ModifierKeys.None);
                            Keyboard.SendKeys(htmlControl, valueToEdit, ModifierKeys.None);
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.EnterBracket, ModifierKeys.None);
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.TabBracket, ModifierKeys.None);
                            break;
                        case UiControls.HtmlCheckBox:
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.SpaceBracket, ModifierKeys.None);
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.TabBracket, ModifierKeys.None);
                            break;
                        case UiControls.HtmlLabel:
                        case UiControls.HtmlFrame:
                            Mouse.Hover(htmlControl);
                            break;
                        case UiControls.HtmlHyperlink:
                            Playback.Wait(1000);
                            Mouse.Click(htmlControl);
                            break;
                        case UiControls.HtmlCell:
                            Mouse.Click(htmlControl);
                            break;
                        case UiControls.HtmlTable:
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.EnterBracket, ModifierKeys.None);
                            break;
                        case UiControls.HtmlEdit:
                        case UiControls.HtmlTextArea:
                            Assert.IsTrue(
                                !string.IsNullOrEmpty(valueToEdit),
                                string.Format(Entities.Constants.Messages.UiControlError, testStep.UiControl.UiControlType, testStep.TestStepNumber));

                            Keyboard.SendKeys(htmlControl, "A", ModifierKeys.Control);
                            var isScanCode = Playback.PlaybackSettings.SendKeysAsScanCode;
                            Playback.PlaybackSettings.SendKeysAsScanCode = false;
                            Keyboard.SendKeys(htmlControl, valueToEdit, ModifierKeys.None);
                            Playback.PlaybackSettings.SendKeysAsScanCode = isScanCode;
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.TabBracket, ModifierKeys.None);
                            break;
                        case UiControls.WinEdit:
                            Assert.IsTrue(
                                !string.IsNullOrEmpty(valueToEdit),
                                string.Format(Entities.Constants.Messages.UiControlError, testStep.UiControl.UiControlType, testStep.TestStepNumber));

                            Keyboard.SendKeys(htmlControl, "A", ModifierKeys.Control);
                            Keyboard.SendKeys(htmlControl, valueToEdit, ModifierKeys.None);
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.TabBracket, ModifierKeys.None);
                            break;
                        case UiControls.WinTreeItem:
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.SpaceBracket, ModifierKeys.None);
                            break;
                        case UiControls.WinTree:
                            Keyboard.SendKeys(htmlControl, valueToEdit, ModifierKeys.None);
                            break;
                        case UiControls.WinButton:
                            Playback.Wait(1000);
                            Mouse.Click(htmlControl, MouseButtons.Left);
                            break;
                        case UiControls.WinRadioButton:
                        case UiControls.WinCheckBox:
                        case UiControls.WinCheckBoxTreeItem:
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.SpaceBracket, ModifierKeys.None);
                            break;
                        case UiControls.WinCalendar:
                        case UiControls.WinCell:
                        case UiControls.WinClient:
                            Playback.Wait(1000);
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.EnterBracket, ModifierKeys.None);
                            break;
                        case UiControls.WinComboBox:
                            Assert.IsTrue(
                                !string.IsNullOrEmpty(valueToEdit),
                                string.Format(Entities.Constants.Messages.UiControlError, testStep.UiControl.UiControlType, testStep.TestStepNumber));

                            var foundItem1 = false;
                            for (var item = 1; item <= htmlControl.ItemCount; item++)
                            {
                                Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.DownBracket, ModifierKeys.None);
                                string displayText =
                                    htmlControl.SelectedItem.Replace(Entities.Constants.FullStop, string.Empty)
                                        .Replace(Entities.Constants.Comma, string.Empty)
                                        .Replace(Entities.Constants.Hyphen, string.Empty)
                                        .Replace(Entities.Constants.Space, string.Empty)
                                        .Replace(Entities.Constants.Percentage, string.Empty)
                                        .Replace(Entities.Constants.Asteric, string.Empty)
                                        .Replace(Entities.Constants.Equal, string.Empty)
                                        .Replace(Entities.Constants.Colon, string.Empty)
                                        .ToUpper();
                                dynamic value =
                                    valueToEdit.Replace(Entities.Constants.FullStop, string.Empty)
                                        .Replace(Entities.Constants.Comma, string.Empty)
                                        .Replace(Entities.Constants.Hyphen, string.Empty)
                                        .Replace(Entities.Constants.Space, string.Empty)
                                        .Replace(Entities.Constants.Space, string.Empty)
                                        .Replace(Entities.Constants.Percentage, string.Empty)
                                        .Replace(Entities.Constants.Asteric, string.Empty)
                                        .Replace(Entities.Constants.Equal, string.Empty)
                                        .Replace(Entities.Constants.Colon, string.Empty)
                                        .ToUpper();
                                if (displayText == value)
                                {
                                    break;
                                }
                            }

                            Assert.IsTrue(foundItem1, failMessage);
                            break;
                        case UiControls.WinCustom:
                        case UiControls.WinDateTimePicker:
                        case UiControls.WinGroup:
                        case UiControls.WinList:
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.SpaceBracket, ModifierKeys.None);
                            break;
                        case UiControls.WinHyperlink:
                        case UiControls.WinListItem:
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.EnterBracket, ModifierKeys.None);
                            break;
                        case UiControls.WinMenu:
                            Mouse.Click(htmlControl, MouseButtons.Left);
                            break;
                        case UiControls.WinMenuBar:
                            Assert.IsTrue(
                                !string.IsNullOrEmpty(valueToEdit),
                                string.Format(Entities.Constants.Messages.UiControlError, testStep.UiControl.UiControlType, testStep.TestStepNumber));
                            Mouse.Click(htmlControl, MouseButtons.Left);
                            Keyboard.SendKeys(htmlControl, valueToEdit, ModifierKeys.None);
                            break;
                        case UiControls.WinMenuItem:
                            Assert.IsTrue(
                                !string.IsNullOrEmpty(valueToEdit),
                                string.Format(Entities.Constants.Messages.UiControlError, testStep.UiControl.UiControlType, testStep.TestStepNumber));
                            Mouse.Click(htmlControl, MouseButtons.Left);
                            break;
                        case UiControls.WinProgressBar:
                        case UiControls.WinRow:
                        case UiControls.WinRowHeader:
                        case UiControls.WinScrollBar:
                        case UiControls.WinStatusBar:
                        case UiControls.WinTable:
                        case UiControls.WinTabList:
                        case UiControls.WinText:
                        case UiControls.WinToolTip:
                        case UiControls.WinTitleBar:
                            Keyboard.SendKeys(htmlControl, Entities.Constants.UiActions.SpaceBracket, ModifierKeys.None);
                            break;
                        case UiControls.WinTabPage:
                        case UiControls.WinToolBar:
                            Mouse.Click(htmlControl, MouseButtons.Left);
                            break;
                        case UiControls.WinWindow:
                            Keyboard.SendKeys(htmlControl, valueToEdit, ModifierKeys.None);
                            break;
                        default:
                            throw new Exception("UiControlType " + testStep.UiControl.UiControlType + " is not supported");
                    }
                }

                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "EditUIControl",
                    Entities.Constants.Pass,
                    "Edit search property " + testStep.UiControl.UiControlSearchProperty + " with search value " + testStep.UiControl.UiControlSearchValue + " completed successfully",
                    testStep.Remarks);

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "EditUIControl",
                    Entities.Constants.Fail,
                    string.Format(Entities.Constants.Messages.DueToException, ex.Message),
                    testStep.Remarks);

                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);
                return false;
            }
        }

        /// <summary>
        /// Wait for user interface control.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool WaitForUi(TestStep testStep)
        {
            try
            {
                if (!string.IsNullOrEmpty(testStep.TestData.ContainsValue(testStep.TestDataKeyToUse).ToString()))
                {
                    var secondsToWaitForUi = Convert.ToInt32(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)]);
                    if (secondsToWaitForUi > 0)
                    {
                        Playback.Wait(secondsToWaitForUi * 1000);
                    }
                }

                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "WaitForUI",
                    Entities.Constants.Pass,
                    "Wait for UI to load completed successfully.",
                    testStep.Remarks);
                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "WaitForUI",
                    Entities.Constants.Fail,
                    string.Format(Entities.Constants.Messages.DueToException, ex.Message),
                    testStep.Remarks);

                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }

        /// <summary>
        /// Verify user interface control.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <param name="passOutAndSave">Step pass or fail and save.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool Verify(TestStep testStep, bool passOutAndSave)
        {
            try
            {
                switch (testStep.Verification.VerificationType.ToUpper())
                {
                    case Entities.Constants.UiActions.BrowerExist:
                        return Verifications.BrowserExist(testStep);
                    case Entities.Constants.UiActions.BrowerNotExists:
                        return Verifications.BrowserNotExist(testStep);
                    case Entities.Constants.UiActions.UiControlExist:
                        return Verifications.UiControlExist(testStep, passOutAndSave);
                    case Entities.Constants.UiActions.UiControlProperty:
                        return Verifications.UiControlProperty(testStep);
                    case Entities.Constants.UiActions.DataBaseValue:
                        return Verifications.DatabaseValue(testStep);
                    default:
                        throw new Exception("VerificationType " + testStep.Verification.VerificationType + " is not supported");
                }
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "Verify",
                    Entities.Constants.Fail,
                    string.Format(Entities.Constants.Messages.DueToException, ex.Message),
                    testStep.Remarks);

                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }

        /// <summary>
        /// Open file option.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool OpenfileOption(TestStep testStep)
        {
            try
            {
                Playback.Wait(5000);

                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "SendKeys",
                    Entities.Constants.Pass,
                    "Send keys " + testStep.TestData.ContainsValue(testStep.TestDataKeyToUse + " completed successfully"),
                    testStep.Remarks);

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "SendKeys",
                    Entities.Constants.Fail,
                    string.Format(Entities.Constants.Messages.DueToException, ex.Message),
                    testStep.Remarks);

                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }

        /// <summary>
        /// Send keys.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool SendKeys(TestStep testStep)
        {
            try
            {
                Assert.IsTrue(
                    testStep.TestData != null,
                    "Test data object does not exist for test step " + testStep.TestStepNumber + ".");

                Assert.IsTrue(
                    testStep.TestData.ContainsValue(testStep.TestDataKeyToUse).ToString() != null,
                    "Test data with key " + testStep.TestDataKeyToUse + "does not exist for test step " + testStep.TestStepNumber + Entities.Constants.FullStop);

                var valueToSend =
                    Convert.ToString(testStep.TestData[Convert.ToInt32(testStep.TestDataKeyToUse)])
                        .ToUpper()
                        .Replace(Entities.Constants.Space, string.Empty);

                dynamic htmlControl = null;
                var test = testStep.UiControl;
                if (test != null)
                {
                    if (!string.IsNullOrEmpty(testStep.UiControl.UiControlType))
                    {
                        htmlControl = UiControls.CreateControl(
                            testStep.UiControl.UiControlType,
                            testStep.UiControl.UiTitle,
                            testStep.UiControl.UiType,
                            testStep.UiControl.UiControlSearchProperty,
                            testStep.UiControl.UiControlSearchValue,
                            testStep.UiControl.UiControlFilterProperty,
                            testStep.UiControl.UiControlFilterValue);
                    }
                }

                if (htmlControl != null)
                {
                    //// Pass the htmlcontrol with the sendkeys command
                    if (valueToSend.Contains(Entities.Constants.UiActions.Alt))
                    {
                        Keyboard.SendKeys(
                            htmlControl,
                            valueToSend.Replace(Entities.Constants.UiActions.Alt, string.Empty),
                            ModifierKeys.Alt);
                    }
                    else if (valueToSend.Contains(Entities.Constants.UiActions.Shift))
                    {
                        Keyboard.SendKeys(
                            htmlControl,
                            valueToSend.Replace(Entities.Constants.UiActions.Shift, string.Empty),
                            ModifierKeys.Shift);
                    }
                    else if (valueToSend.Contains(Entities.Constants.UiActions.Control))
                    {
                        Keyboard.SendKeys(
                            htmlControl,
                            valueToSend.Replace(Entities.Constants.UiActions.Control, string.Empty),
                            ModifierKeys.Control);
                    }
                    else
                    {
                        Keyboard.SendKeys(htmlControl, valueToSend);
                    }
                }
                else
                {
                    if (valueToSend.Contains(Entities.Constants.UiActions.Alt))
                    {
                        Keyboard.SendKeys(valueToSend.Replace(Entities.Constants.UiActions.Alt, string.Empty), ModifierKeys.Alt);
                    }
                    else if (valueToSend.Contains(Entities.Constants.UiActions.Shift))
                    {
                        Keyboard.SendKeys(valueToSend.Replace(Entities.Constants.UiActions.Shift, string.Empty), ModifierKeys.Shift);
                    }
                    else if (valueToSend.Contains(Entities.Constants.UiActions.Control))
                    {
                        Keyboard.SendKeys(valueToSend.Replace(Entities.Constants.UiActions.Control, string.Empty), ModifierKeys.Control);
                    }
                    else
                    {
                        Keyboard.SendKeys(valueToSend);
                    }
                }

                Playback.Wait(5000);

                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "SendKeys",
                    Entities.Constants.Pass,
                    "Send keys " + testStep.TestData.ContainsValue(testStep.TestDataKeyToUse) + " completed successfully",
                    testStep.Remarks);

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "SendKeys",
                    Entities.Constants.Fail,
                    string.Format(Entities.Constants.Messages.DueToException, ex.Message),
                    testStep.Remarks);

                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }

        /// <summary>
        /// Save user interface control.
        /// </summary>
        /// <param name="testStep">Current Test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool SaveUiControlValue(TestStep testStep)
        {
            try
            {
                var errorMessage = string.Empty;
                dynamic key = testStep.TestData.ContainsValue(testStep.TestDataKeyToUse).ToString();

                if (string.IsNullOrEmpty(key))
                {
                    throw new Exception(
                        "A key (id) needs to be set as test data when using action SaveUIControlValueAttribute");
                }

                HtmlControl = UiControls.CreateControl(
                    testStep.UiControl.UiControlType,
                    testStep.UiControl.UiTitle,
                    testStep.UiControl.UiType,
                    testStep.UiControl.UiControlSearchProperty,
                    testStep.UiControl.UiControlSearchValue,
                    testStep.UiControl.UiControlFilterProperty,
                    testStep.UiControl.UiControlFilterValue);

                switch (testStep.UiControl.UiControlType.ToUpper())
                {
                    case UiControls.HtmlControl:
                        var htmlcontrol = new HtmlControl();
                        ValueAttribute = htmlcontrol.ValueAttribute;
                        break;
                    case UiControls.HtmlTable:
                        var htmltable = new HtmlTable();
                        ValueAttribute = htmltable.ValueAttribute;
                        break;
                    case UiControls.HtmlEdit:
                        var htmledit = new HtmlEdit();
                        ValueAttribute = htmledit.ValueAttribute;
                        break;
                    case UiControls.HtmlInputButton:
                        var htmlInputbutton = new HtmlInputButton();
                        ValueAttribute = htmlInputbutton.ValueAttribute;
                        break;
                    case UiControls.HtmlButton:
                        var htmlbutton = new HtmlButton();
                        ValueAttribute = htmlbutton.ValueAttribute;
                        break;
                    case UiControls.HtmlRadioButton:
                        var htmlradiobutton = new HtmlRadioButton();
                        ValueAttribute = htmlradiobutton.ValueAttribute;
                        break;
                    case UiControls.HtmlHyperlink:
                        var htmlhyperlink = new HtmlHyperlink();
                        ValueAttribute = htmlhyperlink.ValueAttribute;
                        break;
                    case UiControls.HtmlList:
                        var htmllist = new HtmlList();
                        ValueAttribute = htmllist.ValueAttribute;
                        break;
                    case UiControls.HtmlComboBox:
                    case UiControls.HtmlComboBoxSelectDirect:
                        var htmlcombobox = new HtmlComboBox();
                        ValueAttribute = htmlcombobox.ValueAttribute;
                        break;
                    case UiControls.HtmlCheckBox:
                        var htmlcheckbox = new HtmlCheckBox();
                        ValueAttribute = htmlcheckbox.ValueAttribute;
                        break;
                    case UiControls.HtmlTextArea:
                        var htmlcheckarea = new HtmlTextArea();
                        ValueAttribute = htmlcheckarea.ValueAttribute;
                        break;
                    case UiControls.HtmlDiv:
                        var htmldiv = new HtmlDiv();
                        ValueAttribute = htmldiv.ValueAttribute;
                        break;
                    case UiControls.HtmlLabel:
                        var htmllable = new HtmlLabel();
                        ValueAttribute = htmllable.ValueAttribute;
                        break;
                    case UiControls.HtmlImage:
                        var htmlimage = new HtmlImage();
                        ValueAttribute = htmlimage.ValueAttribute;
                        break;
                    case UiControls.HtmlSpan:
                        var htmlspan = new HtmlSpan();
                        ValueAttribute = htmlspan.ValueAttribute;
                        break;
                    case UiControls.HtmlCustom:
                        var htmlcustom = new HtmlCustom();
                        ValueAttribute = htmlcustom.ValueAttribute;
                        break;
                    case UiControls.HtmlCell:
                        var htmlcell = new HtmlCell();
                        ValueAttribute = htmlcell.ValueAttribute;
                        break;
                    default:
                        throw new Exception("UiControlType " + testStep.UiControl.UiControlType + " is not supported");
                }

                if (TestCase.TestDataSavedValues.ContainsKey(key))
                {
                    TestCase.TestDataSavedValues.Remove(key);
                }

                if (!string.IsNullOrEmpty(ValueAttribute))
                {
                    TestCase.TestDataSavedValues.Add(key, ValueAttribute);
                }

                UiActions.BufferApps(key, ValueAttribute, ref errorMessage);

                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "SaveUIControlValueAttribute",
                    Entities.Constants.Pass,
                    "SaveUIControlValueAttribute = " + ValueAttribute + " completed successfully",
                    testStep.Remarks);

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    testStep.TestDataKeyToUse,
                    testStep.TestStepNumber,
                    "SaveUIControlValueAttribute",
                    Entities.Constants.Fail,
                    string.Format(Entities.Constants.Messages.DueToException, ex.Message),
                    testStep.Remarks);

                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);

                return false;
            }
            finally
            {
                if (testStep.TestData != null)
                {
                    General.ReleaseObject(testStep.TestData);
                }
            }
        }

        /// <summary>
        /// Verify user interface control.
        /// </summary>
        /// <param name="configStep">Configuration of test Step.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool BufferValues(ConfigStep configStep)
        {
            try
            {
                dynamic key = configStep.TestVariableName;
                dynamic valueAttribute = configStep.TestDataValue;

                if (TestCase.TestDataSavedValues.ContainsKey(key))
                {
                    TestCase.TestDataSavedValues.Remove(key);
                }

                if (!string.IsNullOrEmpty(valueAttribute))
                {
                    TestCase.TestDataSavedValues.Add(key, valueAttribute);
                }

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    configStep.TestConfigKeyToUse,
                    configStep.TestStepNo,
                    "SaveUIControlValueAttribute",
                    Entities.Constants.Fail,
                    string.Format(Entities.Constants.Messages.DueToException, ex.Message),
                    configStep.Remarks);

                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }

        /// <summary>
        /// Verify user interface control.
        /// </summary>
        /// <param name="bufferKey">Buffer key value.</param>
        /// <param name="bufValueAttribute">Buffer value attribute.</param>
        /// <param name="errorMessage">Error Message.</param>
        /// <returns>Returns current step s true or false.</returns>
        public static bool BufferApps(string bufferKey, string bufValueAttribute, ref string errorMessage)
        {
            var applicationClass = new ApplicationClass();
            try
            {
                TestCase.TestStepList = new List<TestStep>();
                TestCase.UiControls = new List<UiControl>();
                ConfigName.TestConfigNames = new List<ConfigStep>();
                TestCase.Verifications = new List<Verification>();

                Data.LoadUiControls(applicationClass);

                Data.LoadVerifications(applicationClass);

                Data.LoadTestCases(applicationClass);

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                LogHelper.ErrorLog(ex, Entities.Constants.ClassName.UiActionsClass, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            finally
            {
                WorkBookUtility.CloseExcel(applicationClass);
            }
        }
    }
}
