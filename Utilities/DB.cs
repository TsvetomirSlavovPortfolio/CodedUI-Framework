// <copyright file="DB.cs" company="Metlife">
// Copyright (c) Metlife. All rights reserved.
// </copyright>
// <summary>DB.cs class helps framework to interact with data bases.</summary>
namespace INF.CodedUI.TestAutomation.Utilities
{
    using System;
    using System.Data.SqlClient;
    using System.Diagnostics.CodeAnalysis;
    using System.Reflection;
    using Configuration;
    using Entities;

    /// <summary>
    /// This class enables framework to interact with data base.
    /// </summary>
    public class Db
    {
        /// <summary>
        /// Method runs a database query using a connection to database server and database name given in app dot config file and returns the value.
        /// </summary>
        /// <param name="query">Data base Query.</param>
        /// <param name="server">Data base server name.</param>
        /// <param name="database">Data base name.</param>
        /// <returns>Query results.</returns>
        [SuppressMessage("Microsoft.Security", "CA2100:Review SQL queries for security vulnerabilities", Justification = "Used during unit testing")]
        public static string ExecuteQuery(string query, string server, string database)
        {
            var result = string.Empty;

            try
            {
                //// Run database query
                var numberOfRows = 0;
                using (var con = new SqlConnection("Server=" + server + ";Database=" + database + ";Trusted_Connection=True;"))
                {
                    con.Open();
                    using (var sql = new SqlCommand(query, con))
                    {
                        using (var dr = sql.ExecuteReader())
                        {
                            if (dr.FieldCount > 1)
                            {
                                throw new Exception("Database query must return only 1 column.");
                            }

                            if (dr.HasRows)
                            {
                                while (dr.Read())
                                {
                                    if (numberOfRows > 1)
                                    {
                                        throw new Exception("Database query must return only 1 row.");
                                    }
                                }
                            }
                        }
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Db, MethodBase.GetCurrentMethod().Name);

                throw;
            }
        }

        /// <summary>
        /// Method validate a database query.
        /// </summary>
        /// <param name="query">Database Query.</param>
        public static void ValidateQuery(string query)
        {
            //// Validation is done according to SQL injection (http://msdn.microsoft.com/en-us/library/ms161953(SQL.105).aspx) plus some more
            try
            {
                if (query.ToUpper().Contains(";"))
                {
                    throw new Exception("Database query must not contain character ;");
                }

                if (query.Contains("'"))
                {
                    throw new Exception("Database query must not contain character '");
                }

                if (query.Contains("--"))
                {
                    throw new Exception("Database query must not contain characters --");
                }

                if (query.Contains("/*"))
                {
                    throw new Exception("Database query must not contain characters /*");
                }

                if (query.Contains("*/"))
                {
                    throw new Exception("Database query must not contain characters */");
                }

                if (query.ToUpper().Contains("XP_"))
                {
                    throw new Exception("Database query must not contain character xp_");
                }

                if (query.ToUpper().Contains("DELETE"))
                {
                    throw new Exception("Database query must not contain a delete statement");
                }

                if (query.ToUpper().Contains("UPDATE"))
                {
                    throw new Exception("Database query must not contain an update statement");
                }

                if (query.ToUpper().Contains("CREATE"))
                {
                    throw new Exception("Database query must not contain a create statement");
                }

                if (query.ToUpper().Contains("INSERT"))
                {
                    throw new Exception("Database query must not contain an insert statement");
                }

                if (query.ToUpper().Contains("CURSOR"))
                {
                    throw new Exception("Database query must not contain a cursor statement");
                }

                if (query.ToUpper().Contains("EXEC"))
                {
                    throw new Exception("Database query must not contain a exec statement");
                }

                if (query.ToUpper().Contains("DROP"))
                {
                    throw new Exception("Database query must not contain a drop statement");
                }

                if (query.ToUpper().Contains("DECLARE"))
                {
                    throw new Exception("Database query must not contain a declare statement");
                }

                if (query.ToUpper().Contains("SET"))
                {
                    throw new Exception("Database query must not contain a set statement");
                }

                if (!query.ToUpper().StartsWith("SELECT"))
                {
                    throw new Exception("Database query must contain a select statement");
                }
            }
            catch (Exception ex)
            {
                LogHelper.ErrorLog(ex, Constants.ClassName.Db, MethodBase.GetCurrentMethod().Name);
                throw;
            }
        }
    }
}