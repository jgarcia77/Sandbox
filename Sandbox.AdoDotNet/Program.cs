using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sandbox.AdoDotNet
{
    class Program
    {
        static string connectionString;
        static SqlConnection conn;
        static IQueryable<BoardPost> boardPosts;

        static void Main(string[] args)
        {
            connectionString = ConfigurationManager.ConnectionStrings["localhost"].ConnectionString;
            conn = new SqlConnection(connectionString);
            conn.Open();

            DataReaderToIEnumerable();

            conn.Close();

            Console.Read();
        }

        static void DataReaderToIEnumerable()
        {
            var commandText = "select * from BoardPost";

            using (var command = new SqlCommand(commandText, conn))
            {
                var reader = command.ExecuteReader();

                boardPosts = reader.Cast<IDataRecord>().Select(record => new BoardPost { Id = record.GetInt64(0), Message = record.GetString(1), IsActive = record.GetBoolean(2), CreatedMemberId = record.GetInt64(3), CreatedDate = record.GetDateTime(4), UpdatedMemberId = record.GetInt64(5), UpdatedDate = record.GetDateTime(6), BoardTopicId = record.GetInt64(7) }).AsQueryable();
            }

            Console.WriteLine(boardPosts.Count());

            //var enumerator = boardPosts.GetEnumerator();

            //if (enumerator.MoveNext())
            //{
            //    var post = enumerator.Current.Message;

            //    Console.WriteLine(post);
            //}

            Console.WriteLine(boardPosts.Count());

            //var first = boardPosts.FirstOrDefault(bp => bp.Id == 1);

            //Console.WriteLine(first.Message);

            Console.WriteLine(boardPosts.Count());
        }
    }
}
