// <copyright file="TestIterations.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>TestIterations.cs class Reads test cases sheets and sets iterations.</summary>
namespace INF.CodedUI.TestAutomation.TestIterations
{
    using System;
    using System.Reflection;
    using Configuration;
    using Entities;
    using Microsoft.VisualStudio.TestTools.UITesting;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using TestCases;

    /// <summary>
    /// Test iterations.
    /// </summary>
    [CodedUITest]
    public class TestIterations : BaseTestClass
    {
        /// <summary>
        /// Initialize the Test.
        /// </summary>
        [TestInitialize]
        public override void TestInitialize()
        {
            base.TestInitialize();
        }

        /// <summary>
        /// Clean up the Test.
        /// </summary>
        [TestCleanup]
        public override void TestCleanup()
        {
            base.TestCleanup();
        }

        /// <summary>
        /// Iterating the Test.
        /// </summary>
        /// 
        [DeploymentItem(Constants.AppSetting.DeploymentItem)]
        [DataSource(Constants.AppSetting.Oledb, Constants.AppSetting.ConnectionString, Constants.AppSetting.TestIterationSheet, DataAccessMethod.Sequential)]
        [TestCategory("CodedUI")]
        [TestMethod]
        [Timeout(TestTimeout.Infinite)]
        public void TestIterationsMethod()
        {
            try
            {
                TestCases.Execute();
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.TestIterationsClass, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}