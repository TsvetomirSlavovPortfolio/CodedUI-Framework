// <copyright file="Program.cs" company="Metlife">
//  Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>Program.cs class handles System and Test Iterations configuration.</summary>
namespace INF.CodedUI.TestAutomation.Entities
{
    using System;
    using Configuration;
    using TestIterations;

    /// <summary>
    /// Class file Program.
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>      
        public static void Main()
        {
            try
            {
                TestIterations objTestIterations = new TestIterations();
                objTestIterations.TestInitialize();
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Program, System.Reflection.MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}