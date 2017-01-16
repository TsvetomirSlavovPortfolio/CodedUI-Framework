// <copyright file="ConfigName.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>ConfigName.cs class reads the values in TestConfigurations file.</summary>
namespace INF.CodedUI.TestAutomation.Entities
{
    using System.Collections;
    using System.Collections.Generic;

    /// <summary>
    /// ConfigName collection base.
    /// </summary>
    public class ConfigName : CollectionBase
    {
        /// <summary>
        /// Gets or sets ConfigName collection base for TestDataConfigCount.
        /// </summary>
        /// <value>Test Data Configuration Count.</value>
        public static int TestDataConfigCount { get; set; }

        /// <summary>
        /// Gets or sets ConfigName collection base for ConfigStep.
        /// </summary>
        /// <value>Test Data Configuration Steps.</value>
        public static List<ConfigStep> TestConfigNames { get; set; }

        /// <summary>
        /// Gets or sets ConfigName collection base for Step Number.
        /// </summary>
        /// <value>Test Data Configuration Step number.</value>
        public static string StepNo { get; set; }

        /// <summary>
        /// Gets or sets ConfigName collection base for Data Type.
        /// </summary>
        /// <value>Data type of Test Data Configuration.</value>
        public static string DataType { get; set; }

        /// <summary>
        /// Gets or sets ConfigName collection base for Variable Name.
        /// </summary>
        /// <value>Variable Name of Test Data Configuration.</value>
        public static string VariableName { get; set; }

        /// <summary>
        /// Gets or sets ConfigName collection base for Test Data Value.
        /// </summary>
        /// <value>Test Data Value of Test Data Configuration.</value>
        public static string TestDataValue { get; set; }

        /// <summary>
        /// Gets or sets ConfigName collection base for Test configurations.
        /// </summary>
        /// <value>Test Data Configuration.</value>
        public static List<TestConfiguration> TestConfigurations { get; set; }

        /// <summary>
        /// Gets or sets ConfigName collection base for Dictionary.
        /// </summary>
        /// <value>Dictionary of Test Data Configuration.</value>
        public static Dictionary<string, string> TestDataSavedValues = new Dictionary<string, string>();
    }
}