// <copyright file="Verifications.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>Verifications.cs class handles test verifications.</summary>
namespace INF.CodedUI.TestAutomation.Verifications
{
    using System;
    using System.Reflection;
    using Configuration;
    using Entities;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using UI;
    using Utilities;

    /// <summary>
    /// Test verifications.
    /// </summary>
    public class Verifications
    {
        /// <summary>
        /// Gets or sets Test steps.
        /// </summary>
        /// <value>Test step value.</value>
        public static dynamic Ts = new TestStep();

        /// <summary>
        /// Verifies that a specific Browser exist.
        /// </summary>
        /// <param name="ts">Test step.</param>
        /// <returns>Browser exists or not.</returns>
        public static bool BrowserExist(TestStep ts)
        {
            try
            {
                var title = Browser.GetTitleFromPartOfTitle(
                    ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString(),
                    UiControls.TypeBrowser,
                    false);

                BrowserWindow.Locate(title);
                Result.PassStepOutandSave(
                    ts.TestDataKeyToUse,
                    ts.TestStepNumber,
                    "Verify BrowserExist",
                    Constants.Pass,
                    "Browser with title (or part of title) " + ts.TestData.ContainsValue(ts.TestDataKeyToUse) + " does exist",
                    ts.Remarks);

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    ts.TestDataKeyToUse,
                    ts.TestStepNumber,
                    "Verify BrowserExist",
                    Constants.Fail,
                    "Due to exception - " + ex.Message + string.Empty,
                    ts.Remarks);

                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }

        /// <summary>
        /// Verifies that a specific Browser doesn't exist.
        /// </summary>
        /// <param name="ts">Test step.</param>
        /// <returns>Browser exist or not.</returns>
        public static bool BrowserNotExist(TestStep ts)
        {
            try
            {
                var title = Browser.GetTitleFromPartOfTitle(
                    ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString(),
                    UiControls.TypeBrowser,
                    false);

                BrowserWindow.Locate(title);
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    ts.TestDataKeyToUse,
                    ts.TestStepNumber,
                    "Verify BrowserNotExist",
                    Constants.Pass,
                    "Browser with title (or part of title) " + ts.TestData.ContainsValue(ts.TestDataKeyToUse) + " doesn't exist",
                    ts.Remarks);

                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);

                return true;
            }

            Result.PassStepOutandSave(
                ts.TestDataKeyToUse,
                ts.TestStepNumber,
                "Verify BrowserNotExist",
                Constants.Fail,
                "Browser with title (or part of title) " + ts.TestData.ContainsValue(ts.TestDataKeyToUse) + " does exist",
                ts.Remarks);

            return false;
        }

        /// <summary>
        /// Verifies that a specific Control exist.
        /// </summary>
        /// <param name="ts">Test step.</param>
        /// <param name="passOutAndSave">Result pass or fail and save it.</param>
        /// <returns>Returns true or false.</returns>
        public static bool UiControlExist(TestStep ts, bool passOutAndSave)
        {
            try
            {
                //// In function UiControls.CreateHtmlControl there will be a check that the UiControl exist
                UiControls.CreateControl(
                    ts.UiControl.UiControlType,
                    ts.UiControl.UiTitle,
                    ts.UiControl.UiType,
                    ts.UiControl.UiControlSearchProperty,
                    ts.UiControl.UiControlSearchValue,
                    ts.UiControl.UiControlFilterProperty,
                    ts.UiControl.UiControlFilterValue);

                if (passOutAndSave)
                {
                    var msg = "UiControl with search property " + ts.UiControl.UiControlSearchProperty + " and search value " + ts.UiControl.UiControlSearchValue + " and filter property " + ts.UiControl.UiControlFilterProperty + " and filter value " + ts.UiControl.UiControlFilterValue + " exist";
                    Result.PassStepOutandSave(
                        ts.TestDataKeyToUse,
                        ts.TestStepNumber,
                        "Verify UIControlExist",
                        Constants.Pass,
                        msg,
                        ts.Remarks);
                }

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    ts.TestDataKeyToUse,
                    ts.TestStepNumber,
                    "Verify UIControlExist",
                    Constants.Fail,
                    "Due to exception - " + ex.Message + string.Empty,
                    ts.Remarks);

                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }

        /// <summary>
        /// Verifies that a value exist for a specific Control.
        /// </summary>
        /// <param name="ts">Test Step.</param>
        /// <returns>Returns true or false.</returns>
        public static bool UiControlProperty(TestStep ts)
        {
            try
            {
                //// Check that mandatory values exist
                if (string.IsNullOrEmpty(ts.Verification.UiControlProperty))
                {
                    throw new Exception(
                        "UiControlProperty for given VerificationId must contain a value when using action Verify");
                }

                if (string.IsNullOrEmpty(ts.Verification.OperatorToUse))
                {
                    throw new Exception("Operator for given VerificationId must contain a value when using action Verify");
                }

                if (ts.TestData.ContainsValue(ts.TestDataKeyToUse))
                {
                    throw new Exception("TestData must contain a value when using action Verify");
                }

                //// Create htmlcontrol
                dynamic htmlcontrol = UiControls.CreateControl(
                    ts.UiControl.UiControlType,
                    ts.UiControl.UiTitle,
                    ts.UiControl.UiType,
                    ts.UiControl.UiControlSearchProperty,
                    ts.UiControl.UiControlSearchValue,
                    ts.UiControl.UiControlFilterProperty,
                    ts.UiControl.UiControlFilterValue);

               //// Verify that UiControl has correct value 
               //var failMessage = "UiControl with search property " + ts.UiControl.UiControlSearchProperty + " and search value " + ts.UiControl.UiControlSearchValue + " and filter property. " + ts.UiControl.UiControlFilterProperty + " and filter value " + ts.UiControl.UiControlFilterValue + " " + ts.TestData.ContainsValue(ts.TestDataKeyToUse) + " hasn't the correct value for property " + ts.Verification.UiControlProperty + ".";
               var failMessage = "UiControl with search property " + ts.UiControl.UiControlSearchProperty + " and search value " + ts.UiControl.UiControlSearchValue + " " + ts.TestData.ContainsValue(ts.TestDataKeyToUse) + " hasn't the correct value for property " + ts.Verification.UiControlProperty + ".";
               switch (ts.Verification.OperatorToUse.ToUpper())
                {
                    case Constants.Operator.EQuals:
                        Assert.IsTrue(
                            htmlcontrol.GetProperty(ts.Verification.UiControlProperty) == ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString(),
                            failMessage);

                        break;

                    case Constants.Operator.Contains:
                        Assert.IsTrue(
                            htmlcontrol.GetProperty(ts.Verification.UiControlProperty).ToString.Contains(ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString()),
                            failMessage);

                        break;

                    case Constants.Operator.NotEquals:
                        Assert.IsTrue(
                            htmlcontrol.GetProperty(ts.Verification.UiControlProperty) != ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString(),
                            failMessage);

                        break;

                    case Constants.Operator.NotContains:
                        Assert.IsFalse(
                            !htmlcontrol.GetProperty(ts.Verification.UiControlProperty).ToString.Contains(ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString()),
                            failMessage);

                        break;

                    default:
                        throw new Exception("Operator " + ts.Verification.OperatorToUse + " is not supported");
                }

                var msg = "UiControl with search property " + ts.UiControl.UiControlSearchProperty + " and search value " + ts.UiControl.UiControlSearchValue + " and filter property " + ts.UiControl.UiControlFilterProperty + " and filter value " + ts.UiControl.UiControlFilterValue + " has the correct value for property " + ts.Verification.UiControlProperty + string.Empty;

                Result.PassStepOutandSave(
                    ts.TestDataKeyToUse,
                    ts.TestStepNumber,
                    "Verify UiControlProperty",
                    Constants.Pass,
                    msg,
                    ts.Remarks);

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    ts.TestDataKeyToUse,
                    ts.TestStepNumber,
                    "Verify UiControlProperty",
                    Constants.Fail,
                    "Due to exception - " + ex.Message + string.Empty,
                    ts.Remarks);

                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }

        /// <summary>
        /// Verifies a specific database value.
        /// </summary>
        /// <param name="ts">Test step.</param>
        /// <returns>Returns true or false.</returns>
        public static bool DatabaseValue(TestStep ts)
        {
            try
            {
                //// Check that mandatory values exist
                if (string.IsNullOrEmpty(ts.Verification.DatabaseQuery))
                {
                    throw new Exception(
                        "DatabaseQuery for given VerificationId must contain a value when using action Verify");
                }

                if (string.IsNullOrEmpty(ts.Verification.DatabaseServer))
                {
                    throw new Exception(
                        "DatabaseServer for given VerificationId must contain a value when using action Verify");
                }

                if (string.IsNullOrEmpty(ts.Verification.DatabaseName))
                {
                    throw new Exception(
                        "DatabaseName for given VerificationId must contain a value when using action Verify");
                }

                if (string.IsNullOrEmpty(ts.Verification.OperatorToUse))
                {
                    throw new Exception("Operator for given VerificationId must contain a value when using action Verify");
                }

                if (ts.TestData.ContainsValue(ts.TestDataKeyToUse))
                {
                    throw new Exception("TestData must contain a value when using action Verify");
                }

                //// Validate query
                Db.ValidateQuery(ts.Verification.DatabaseQuery);

                //// Replace #<value># in the query with current test case start time or from a saved value
                dynamic pos1 = ts.Verification.DatabaseQuery.IndexOf("#", StringComparison.CurrentCulture) + 1;

                while (pos1 > 0)
                {
                    dynamic pos2 = ts.Verification.DatabaseQuery.IndexOf("#", pos1);

                    //// Get the tag value from query 
                    dynamic tagValueFromQuery = ts.Verification.DatabaseQuery.Substring(pos1, pos2 - pos1);
                    if (tagValueFromQuery.ToUpper == "testCaseSTARTDATETIME")
                    {
                        //// Get the start time for current test case
                        tagValueFromQuery = "#" + tagValueFromQuery + "#";
                        ts.Verification.DatabaseQuery = ts.Verification.DatabaseQuery.Replace(tagValueFromQuery, "'" + Timing.TestCaseStartTime + "'");
                    }
                    else if (TestCase.TestDataSavedValues.ContainsKey(tagValueFromQuery))
                    {
                        //// Get saved value if it exist
                        if (IsNumeric())
                        {
                            dynamic savedValue = TestCase.TestDataSavedValues.ContainsKey(tagValueFromQuery);
                            ts.Verification.DatabaseQuery = ts.Verification.DatabaseQuery.Replace(tagValueFromQuery, savedValue);
                        }
                        else
                        {
                            //// If string value add char '
                            dynamic savedValue = string.Format("'{0}'", TestCase.TestDataSavedValues.ContainsKey(tagValueFromQuery));
                            ts.Verification.DatabaseQuery = ts.Verification.DatabaseQuery.Replace(tagValueFromQuery, savedValue);
                        }
                    }

                    pos1 = ts.Verification.DatabaseQuery.IndexOf("#", StringComparison.CurrentCulture) + 1;
                }

                //// Run db query
                dynamic databaseValue = Db.ExecuteQuery(
                    ts.Verification.DatabaseQuery,
                    ts.Verification.DatabaseServer,
                    ts.Verification.DatabaseName);

                //// Verify that database value returned from query is correct 
                var failMessage = "Database value wasn't correct.";
                switch (ts.Verification.OperatorToUse.ToUpper())
                {
                    case Constants.Operator.EQuals:
                        Assert.IsTrue(databaseValue == ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString(), failMessage);
                        break;
                    case Constants.Operator.Contains:
                        Assert.IsTrue(databaseValue.Contains(ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString()), failMessage);
                        break;
                    case Constants.Operator.NotEquals:
                        Assert.IsTrue(databaseValue != ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString(), failMessage);
                        break;
                    case Constants.Operator.NotContains:
                        Assert.IsTrue(!databaseValue.Contains(ts.TestData.ContainsValue(ts.TestDataKeyToUse).ToString()), failMessage);
                        break;
                    default:
                        throw new Exception("Operator " + ts.Verification.OperatorToUse + " is not supported");
                }

                Result.PassStepOutandSave(ts.TestDataKeyToUse, ts.TestStepNumber, "Verify database value", Constants.Pass, "Database value was correct.", ts.Remarks);

                return true;
            }
            catch (Exception ex)
            {
                Result.PassStepOutandSave(
                    ts.TestDataKeyToUse,
                    ts.TestStepNumber,
                    "Verify database value",
                    Constants.Fail,
                    "Due to exception - " + ex.Message + string.Empty,
                    ts.Remarks);

                LogHelper.ErrorLog(ex, Constants.ClassName.Verifications, MethodBase.GetCurrentMethod().Name);

                return false;
            }
        }

        /// <summary>
        /// Verifies is numeric or not.
        /// </summary>
        /// <returns>Returns Exception message.</returns>
        private static bool IsNumeric()
        {
            throw new NotImplementedException();
        }
    }
}