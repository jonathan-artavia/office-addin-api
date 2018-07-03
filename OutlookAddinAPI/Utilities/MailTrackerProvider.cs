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

        //public virtual long SaveTracker(MailItem mailItem)
        //{
        //    DataTable tb = this.FindByConversationId(mailItem.ConversationID);
        //    if (tb == null || tb.Rows.Count == 0)
        //    {
        //        long id = DateTime.Now.Ticks;
        //        SqlCommand cmd1 = new SqlCommand("INSERT INTO [dbo].[tblCmMailTracker] (ID, ConversationID, StartDate, StartUserID) VALUES (@ID, @ConversationID, @StartDate, @StartUserID)", this.Connection);
        //        cmd1.Parameters.AddWithValue("@ID", id);
        //        cmd1.Parameters.AddWithValue("@ConversationID", mailItem.ConversationID);
        //        cmd1.Parameters.AddWithValue("@StartDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
        //        cmd1.Parameters.AddWithValue("@StartUserID", OutlookRibbon.Application.Session.CurrentUser.Name.Substring(0, 19));
        //        cmd1.ExecuteNonQuery();
        //        return id;
        //    }
        //    return tb.Rows[0].Field<long>("ID");
        //}

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

        //public virtual bool SaveEmail(Outlook.MailItem mailItem, long trackerID, string senderEmail, string recipientsEmail)
        //{
        //    DataTable tb = this.FindByMessageId(mailItem.EntryID);
        //    if (tb == null || tb.Rows.Count == 0)
        //    {
        //        byte[] email = Utility.ConvertToBytes(mailItem);
        //        SqlCommand cmd1 = new SqlCommand("INSERT INTO [dbo].[tblCmMail] (ID, MessageID, TrackerID, [Subject], Sender, Recipients, Mail, Processed) VALUES (@ID, @MessageID, @TrackerID, @Subject, @Sender, @Recipients, @Mail, @Processed)", this.Connection);
        //        cmd1.Parameters.AddWithValue("@ID", DateTime.Now.Ticks);
        //        cmd1.Parameters.AddWithValue("@MessageID", mailItem.EntryID);
        //        cmd1.Parameters.AddWithValue("@TrackerID", trackerID);
        //        cmd1.Parameters.AddWithValue("@Subject", mailItem.Subject != null ? mailItem.Subject : string.Empty);
        //        cmd1.Parameters.AddWithValue("@Sender", senderEmail);
        //        cmd1.Parameters.AddWithValue("@Recipients", recipientsEmail);
        //        cmd1.Parameters.Add("@Mail", SqlDbType.VarBinary, email.Length).Value = email;
        //        cmd1.Parameters.AddWithValue("@Processed", mailItem.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss.fff"));
        //        cmd1.ExecuteNonQuery();
        //    }
        //    return true;
        //}
        public virtual bool IsTracked(string conversationId)
        {
            DataTable dt = this.FindByConversationId(conversationId);
            return !(dt == null || dt.Rows.Count == 0 || !dt.Rows[0].IsNull("StopDate"));
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