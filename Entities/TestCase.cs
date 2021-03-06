// <copyright file="TestCase.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>TestCase.cs class handles TestCase programs.</summary>
namespace INF.CodedUI.TestAutomation.Entities
{
    using System.Collections;
    using System.Collections.Generic;

    /// <summary>
    /// Test Case Collection Base.
    /// </summary>
    public class TestCase : CollectionBase
    {
        /// <summary>
        /// Gets or sets Test Case Dictionary for Test data saved values.
        /// </summary>
        /// <value>Value as array.</value>
        public static Dictionary<string, string> TestDataSavedValues = new Dictionary<string, string>();

        /// <summary>
        /// Gets or sets Application value.
        /// </summary>
        /// <value>Application value.</value>
        public static string Application { get; set; }

        /// <summary>
        /// Gets or sets Name value.
        /// </summary>
        /// <value>Name value.</value>
        public static string Name { get; set; }

        /// <summary>
        /// Gets or sets Description value.
        /// </summary>
        /// <value>Description value.</value>
        public static string Description { get; set; }

        /// <summary>
        /// Gets or sets File Name value.
        /// </summary>
        /// <value>File Name value.</value>
        public static string FileName { get; set; }

        /// <summary>
        /// Gets or sets Root File Path.
        /// </summary>
        /// <value>Root File Path value.</value>
        public static string RootFilePath { get; set; }

        /// <summary>
        /// Gets or sets Test Report File Name Prefix.
        /// </summary>
        /// <value>Test Report File Name Prefix value.</value>
        public static string TestReportFileNamePrefix { get; set; }

        /// <summary>
        /// Gets or sets Test Step List.
        /// </summary>
        /// <value>Test Step List value.</value>
        public static List<TestStep> TestStepList { get; set; }

        /// <summary>
        /// Gets or sets Test Data Count.
        /// </summary>
        /// <value>Test Data Count value.</value>
        public static int TestDataCount { get; set; }

        /// <summary>
        /// Gets or sets Controls.
        /// </summary>
        /// <value>Control value.</value>
        public static List<UiControl> UiControls { get; set; }

        /// <summary>
        /// Gets or sets Verifications.
        /// </summary>
        /// <value>Verifications value.</value>
        public static List<Verification> Verifications { get; set; }

        /// <summary>
        /// Gets or sets Remarks.
        /// </summary>
        /// <value>Remarks value.</value>
        public static string Remarks { get; set; }
    }
}