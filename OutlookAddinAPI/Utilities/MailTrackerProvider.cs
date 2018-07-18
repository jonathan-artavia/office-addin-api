using OutlookAddinAPI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace OutlookAddinAPI
{
    public class MailTrackerProvider : IDisposable
    {
        #region Fields
        private SqlConnection _connection;
        internal static string _connectionString;
        #endregion

        #region Constructor
        public MailTrackerProvider() { }
        #endregion

        #region Public
        public virtual bool OpenConnection()
        {
            if (_connection == null)
                _connection = new SqlConnection(_connectionString);
            if (_connection.State == ConnectionState.Closed)
                _connection.Open();
            if (_connection.State == ConnectionState.Open)
                return true;
            else
                return false;
        }

        public virtual long StartTracker(string mailItem, string conversationId, string displayName)
        {
            
            DataTable tb = this.FindByConversationId(conversationId);
            if (!this.IsTracked(conversationId) & (tb != null && tb.Rows.Count == 0))
            {
                long id = DateTime.Now.Ticks;
                Dictionary<string, object> pairs = new Dictionary<string, object>();
                pairs.Add("@ID", id);
                pairs.Add("@ConversationID", conversationId);
                pairs.Add("@StartDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                pairs.Add("@StartUserID", displayName);

                SqlCommand cmd = this.Connection.CreateCommand();
                cmd.CommandText = "INSERT INTO [dbo].[tblCmMailTracker] (ID, ConversationID, StartDate, StartUserID) VALUES (@ID, @ConversationID, @StartDate, @StartUserID)";
                cmd.Parameters.AddWithValue("@ID", id);
                cmd.Parameters.AddWithValue("@ConversationID", conversationId);
                cmd.Parameters.AddWithValue("@StartDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                cmd.Parameters.AddWithValue("@StartUserID", displayName);

                cmd.ExecuteNonQuery();

                this.SaveEmail(mailItem, id);

                return id;
            }
            else
            {
                SqlCommand cmd1 = new SqlCommand("UPDATE [dbo].[tblCmMailTracker] SET [StopDate] = NULL, [StopUserID] = NULL WHERE [ID] = @ID", this.Connection);
                cmd1.Parameters.AddWithValue("@ID", tb.Rows[0].Field<long>("ID"));
                cmd1.ExecuteNonQuery();

                this.SaveEmail(mailItem, tb.Rows[0].Field<long>("ID"));
            }
            return tb.Rows[0].Field<long>("ID");
        }

        public virtual long StopTracker(string mailItem, string conversationId, string displayName)
        {

            DataTable tb = this.FindByConversationId(conversationId);
            if (this.IsTracked(conversationId) & (tb != null && tb.Rows.Count > 0))
            {
                SqlCommand cmd1 = new SqlCommand("UPDATE [dbo].[tblCmMailTracker] SET [StopDate] = @StopDate, [StopUserID] = @StopUserID WHERE [ID] = @ID", this.Connection);
                cmd1.Parameters.AddWithValue("@ID", tb.Rows[0].Field<long>("ID"));
                cmd1.Parameters.AddWithValue("@StopDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                cmd1.Parameters.AddWithValue("@StopUserID", displayName);
                cmd1.ExecuteNonQuery();

                this.SaveEmail(mailItem, tb.Rows[0].Field<long>("ID"));

                return tb.Rows[0].Field<long>("ID");
            }
            return 0;
        }

        //public virtual bool UpdateTracker(long trackerId, bool track)
        //{
        //    if (track)
        //    {
        //        SqlCommand cmd1 = new SqlCommand("UPDATE [dbo].[tblCmMailTracker] SET [StopDate] = NULL, [StopUserID] = NULL WHERE [ID] = @ID", this.Connection);
        //        cmd1.Parameters.AddWithValue("@ID", trackerId);
        //        cmd1.ExecuteNonQuery();
        //        return true;
        //    }
        //    else
        //    {
        //        SqlCommand cmd1 = new SqlCommand("UPDATE [dbo].[tblCmMailTracker] SET [StopDate] = @StopDate, [StopUserID] = @StopUserID WHERE [ID] = @ID", this.Connection);
        //        cmd1.Parameters.AddWithValue("@ID", trackerId);
        //        cmd1.Parameters.AddWithValue("@StopDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
        //        cmd1.Parameters.AddWithValue("@StopUserID", OutlookRibbon.Application.Session.CurrentUser.Name.Substring(0, 19));
        //        cmd1.ExecuteNonQuery();
        //        return true;
        //    }
        //}

        public virtual bool SaveEmail(string body, long trackerID)
        {
            MailItem mailItem = MailItem.FromJson(body);
            if (mailItem != null)
            {
                if (this.IsTracked(trackerID))
                {
                    DataTable tb = this.FindByMessageId(mailItem.DataP0.DataP0.Id);
                    if (tb == null || tb.Rows.Count == 0)
                    {
                        byte[] email = System.Text.Encoding.UTF8.GetBytes(body);
                        SqlCommand cmd1 = new SqlCommand("INSERT INTO [dbo].[tblCmMail] (ID, MessageID, TrackerID, [Subject], Sender, Recipients, Mail, Processed) VALUES (@ID, @MessageID, @TrackerID, @Subject, @Sender, @Recipients, @Mail, @Processed)", this.Connection);
                        cmd1.Parameters.AddWithValue("@ID", DateTime.Now.Ticks);
                        cmd1.Parameters.AddWithValue("@MessageID", mailItem.DataP0.DataP0.Id);
                        cmd1.Parameters.AddWithValue("@TrackerID", trackerID);
                        cmd1.Parameters.AddWithValue("@Subject", mailItem.DataP0.DataP0.Subject ?? string.Empty);
                        cmd1.Parameters.AddWithValue("@Sender", mailItem.DataP0.DataP0.Sender.Address);
                        cmd1.Parameters.AddWithValue("@Recipients", string.Join(",", mailItem.DataP0.DataP0.To.Select(x => x.Address)));
                        cmd1.Parameters.Add("@Mail", SqlDbType.VarBinary, email.Length).Value = email;
                        cmd1.Parameters.AddWithValue("@Processed", new DateTime(1970, 1, 1, 0, 0, 0, 0).AddMilliseconds(mailItem.DataP0.DataP0.DateTimeCreated).ToString("yyyy-MM-dd HH:mm:ss.fff"));
                        cmd1.ExecuteNonQuery();
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                throw new InvalidCastException("Mail JSON string arrived with wrong format");
            }
        }

        public virtual bool IsTracked(string conversationId)
        {

            DataTable dt = this.FindByConversationId(conversationId);
            return (!(dt == null || dt.Rows.Count == 0 || !dt.Rows[0].IsNull("StopDate")));
        }

        public virtual bool IsTracked(long trackerId)
        {

            DataTable dt = this.FindByTrackerId(trackerId);
            return (!(dt == null || dt.Rows.Count == 0 || !dt.Rows[0].IsNull("StopDate")));
        }

        public virtual DataTable FindByConversationId(string id)
        {
            SqlDataReader rdr = null;
            DataTable tb = new DataTable();
            SqlCommand cmd = new SqlCommand("SELECT * FROM [dbo].[tblCmMailTracker] WHERE ConversationID = @ConvId", this.Connection);
            cmd.Parameters.AddWithValue("@ConvId", id);
            rdr = cmd.ExecuteReader();
            tb.Load(rdr);
            if (rdr != null)
            {
                rdr.Close();
            }
            return tb;
        }

        public virtual DataTable FindByTrackerId(long trackerId)
        {
            SqlDataReader rdr = null;
            DataTable tb = new DataTable();
            SqlCommand cmd = new SqlCommand("SELECT * FROM [dbo].[tblCmMailTracker] WHERE ID = @TrackerId", this.Connection);
            cmd.Parameters.AddWithValue("@TrackerId", trackerId);
            rdr = cmd.ExecuteReader();
            tb.Load(rdr);
            if (rdr != null)
            {
                rdr.Close();
            }
            return tb;
        }

        public virtual DataTable FindByMessageId(string id)
        {
            SqlDataReader rdr = null;
            DataTable tb = new DataTable();
            SqlCommand cmd = new SqlCommand("SELECT * FROM [dbo].[tblCmMail] WHERE MessageID = @MsgId", this.Connection);
            cmd.Parameters.AddWithValue("@MsgId", id);
            rdr = cmd.ExecuteReader();
            tb.Load(rdr);
            if (rdr != null)
            {
                rdr.Close();
            }
            return tb;
        }

        public virtual void Dispose()
        {
            if (this.Connection != null && this.Connection.State != ConnectionState.Closed)
            {
                this.Connection.Close();
                this._connection = null;
            }
        }
        #endregion

        #region Properties
        public virtual SqlConnection Connection
        {
            get { return this._connection; }
        }
        #endregion
    }
}