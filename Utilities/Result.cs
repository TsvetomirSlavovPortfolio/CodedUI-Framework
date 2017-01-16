// <copyright file="Result.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>Result.cs Generates Result based on test results.</summary>
namespace INF.CodedUI.TestAutomation.Utilities
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;
    using Configuration;
    using Entities;

    /// <summary>
    /// Collection of test results.
    /// </summary>
    public class Result
    {
        /// <summary>
        /// Gets or sets Collection of test results.
        /// </summary>
        /// <value>Test Result value.</value>
        public static Queue<TestResult> TestStepsResultsCollection = new Queue<TestResult>();

        /// <summary>
        /// Called at the end of every step to store results in collection.
        /// </summary>
        /// <param name="testdataiterationnr">Test iteration number.</param>
        /// <param name="stepnr">Step number.</param>
        /// <param name="description">Description of step.</param>
        /// <param name="result">Step Result.</param>
        /// <param name="comments">Step comments.</param>
        /// <param name="remark">Step Remarks.</param>
        public static void PassStepOutandSave(string testdataiterationnr, string stepnr, string description, string result, string comments, string remark)
        {
            try
            {
                var logObj = new TestResult
                {
                    TestDataIterationNr = testdataiterationnr,
                    StepNr = stepnr,
                    Result = result,
                    Description = description,
                    Comment = comments,
                    Remarks = remark
                };
                AddTestScriptResulttoCollection(logObj);
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Result, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }

        /// <summary>
        ///     Enqueue the step result on the queue testStepsResultsCollection.
        /// </summary>
        /// <param name="obj">Refers the test step results.</param>
        public static void AddTestScriptResulttoCollection(TestResult obj)
        {
            TestStepsResultsCollection.Enqueue(obj);
        }

        /// <summary>
        ///     Evaluates the script result from testStepsResultsCollection.
        /// </summary>
        /// <returns>Returns Fail if any Step has Fail result, True otherwise.</returns>
        public static string GetTestScriptResult()
        {
            try
            {
                if (TestStepsResultsCollection.Count == 0)
                {
                    return Constants.Fail;
                }

                foreach (var obj in TestStepsResultsCollection)
                {
                    if (obj.Result == Constants.Fail | obj.Result == Constants.Fail)
                    {
                        return Constants.Fail;
                    }
                }

                return Constants.Pass;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Result, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}