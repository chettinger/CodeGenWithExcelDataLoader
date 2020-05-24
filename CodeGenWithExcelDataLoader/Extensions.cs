using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CodeGenWithExcelDataLoader
{
   public static class Extensions
    {
        /// <summary>
        /// Returns list of entities
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="service"></param>
        /// <param name="filterExpression"></param>
        /// <param name="columnSet"></param>
        /// <returns></returns>
        public static List<TEntity> RetrieveEntities<TEntity>(this IOrganizationService service, FilterExpression filterExpression = null, ColumnSet columnSet = null)
        where TEntity : Entity
        {
            var entities = RetrieveEntities(service, typeof(TEntity), columnSet, filterExpression);
            var typedEntities = (from e in entities select e.ToEntity<TEntity>()).ToList();
            return typedEntities;
        }

        /// <summary>
        /// Returns list of entities
        /// </summary>
        /// <param name="service"></param>
        /// <param name="entityType"></param>
        /// <param name="columnSet"></param>
        /// <param name="filterExpression"></param>
        /// <returns></returns>
        private static List<Entity> RetrieveEntities(this IOrganizationService service, Type entityType, ColumnSet columnSet, FilterExpression filterExpression = null)
        {
            entityType.VerifyIsEntity();

            var queryCount = 5000;
            var pageNumber = 1;
            var pageQuery = new QueryExpression
            {
                EntityName = entityType.Name.ToLower(),
                ColumnSet = columnSet ?? new ColumnSet(true)
            };
            if (filterExpression != null)
            {
                pageQuery.Criteria.AddFilter(filterExpression);
            }

            // Assign the pageinfo properties to the query expression.
            pageQuery.PageInfo = new PagingInfo()
            {
                Count = queryCount,
                PageNumber = pageNumber,
                // The current paging cookie. When retrieving the first page, 
                // pagingCookie should be null.
                PagingCookie = null
            };
            var entities = new List<Entity>();
            while (true)
            {
                // Retrieve the page.
                var results = service.RetrieveMultiple(pageQuery);
                if (results.Entities != null)
                {
                    entities.AddRange(results.Entities);
                }

                // Check for more records, if it returns true.
                if (results.MoreRecords)
                {
                    // Increment the page number to retrieve the next page.
                    pageQuery.PageInfo.PageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.
                    pageQuery.PageInfo.PagingCookie = results.PagingCookie;
                }
                else
                {
                    // If no more records are in the result nodes, exit the loop.
                    break;
                }
            }

            return entities;
        }

        /// <summary>
        /// Verifies type is CRM Entity
        /// </summary>
        /// <param name="type"></param>
        public static void VerifyIsEntity(this Type type)
        {
            if (!type.IsSubclassOf(typeof(Entity)))
            {
                throw new Exception("Type must be a subclass of Entity.");
            }
        }
        public static void RetryOnException(int times, TimeSpan delay, Action operation)
        {
            var attempts = 0;
            do
            {
                try
                {
                    attempts++;
                    operation();
                    break; // Sucess! Lets exit the loop!
                }
                catch (Exception ex)
                {
                    if (attempts == times)
                        throw;

                    Console.WriteLine($"Exception caught on attempt {attempts} - will retry after delay {delay}");

                    Task.Delay(delay).Wait();
                }
            } while (true);
        }
    }
}
