// <copyright file="Timing.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>Timing.cs Start and End time of each test case execution</summary>
namespace INF.CodedUI.TestAutomation.Utilities
{
    using System;

    /// <summary>
    /// Set up time.
    /// </summary>
    public class Timing
    {
        /// <summary>
        /// Gets or sets up time for its start time.
        /// </summary>
        /// <value>Time for test starts.</value>
        public static DateTime TotalStartTime { get; set; }

        /// <summary>
        /// Gets or sets up time for its total end time.
        /// </summary>
        /// <value>Total end Time.</value>
        public static DateTime TotalEndTime { get; set; }

        /// <summary>
        /// Gets or sets Test case start time.
        /// </summary>
        /// <value>Test case start time.</value>
        public static DateTime TestCaseStartTime { get; set; }

        /// <summary>
        /// Gets or sets Test case end time.
        /// </summary>
        /// <value>Test case end time.</value>
        public static DateTime TestCaseEndTime { get; set; }

        /// <summary>
        /// Gets or sets Duration of test.
        /// </summary>
        /// <value>Test case duration.</value>
        public static TimeSpan TestCaseduration { get; set; }

        /// <summary>
        /// Gets or sets Total duration.
        /// </summary>
        /// <value>Total duration of all tests.</value>
        public static TimeSpan Totalduration { get; set; }
    }
}