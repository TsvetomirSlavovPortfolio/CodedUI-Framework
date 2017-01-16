// <copyright file="ConfigStep.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>ConfigStep.cs class handles Configuration step variables.</summary>
namespace INF.CodedUI.TestAutomation.Entities
{
    using System.Collections;
    using System.Collections.Generic;

    /// <summary>
    /// ConfigStep reads steps from Configuration.
    /// </summary>
    public class ConfigStep : CollectionBase
    {
        /// <summary>
        /// Gets or sets Test Step Number.
        /// </summary>
        /// <value>Test Step Number.</value>
        public string TestStepNo { get; set; }

        /// <summary>
        /// Gets or sets Test Data Type.
        /// </summary>
        /// <value>Test Data Type.</value>
        public string TestDataType { get; set; }

        /// <summary>
        /// Gets or sets Test Variable Name.
        /// </summary>
        /// <value>Test Variable Name.</value>
        public string TestVariableName { get; set; }

        /// <summary>
        /// Gets or sets Test Data Value.
        /// </summary>
        /// <value>Test Data Value.</value>
        public string TestDataValue { get; set; }

        /// <summary>
        /// Gets or sets Configuration Action.
        /// </summary>
        /// <value>Configuration Action.</value>
        public string ConfigAction { get; set; }

        /// <summary>
        /// Gets or sets Test Configuration Key To Use.
        /// </summary>
        /// <value>Test Configuration Key To Use.</value>
        public string TestConfigKeyToUse { get; set; }

        /// <summary>
        /// Gets or sets Test Data Configuration.
        /// </summary>
        /// <value>Test Data Configuration.</value>
        public Dictionary<int, string> TestDataConfig { get; set; }

        /// <summary>
        /// Gets or sets Remarks.
        /// </summary>
        /// <value>Remarks column.</value>
        public string Remarks { get; set; }
    }
}