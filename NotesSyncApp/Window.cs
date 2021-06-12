using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Data;
using System.Drawing;
using X.HtmlToRtfConverter;

namespace NotesSyncApp
{
    public partial class MainWindow : Form
    {

        //-------- Window Communication Setup -----------//


        [DllImport("user32", EntryPoint = "FindWindowExA")]
        private static extern int FindWindowEx(int hWnd1, int hWnd2, string lpsz1, string lpsz2);

        [DllImport("user32", EntryPoint = "SendMessageA")]
        private static extern int ZSendMessage(int Hwnd, int wMsg, int wParam, long lParam);


        // The Stickies API is accessed by sending text strings to the main Stickies window. 
        // The Windows message which is used is WM_COPYDATA which is available from Windows 95 onwards, 
        // and is a way of passing small amounts of data directly between two windows.

        private const short WM_COPYDATA = 0x4A;

        // In return, Stickies will reply with an acknowledgement of the command, 
        // and information as to whether the command has worked, 
        // and also possibly some data in return depending on the command used.

        // To use the API, your application will find the Stickies main window, which has a title of ZhornSoftwareStickiesMain, 
        // and then use the Windows API call SendMessage to pass a COPYDATASTRUCT which contains the command to be sent. 

        public struct COPYDATASTRUCT
        {
            public IntPtr dwData;
            public int cbData;
            public IntPtr lpData;
        }

        // When Stickies replies, your application will receive a WM_COPYDATA message, again with a COPYDATASTRUCT filled in. 
        // Prefix strings sent to Stickies using WM_COPYDATA with "api", meaning that for example the full string passed to SendMessage might be "api do ping".

        public int m_commandref;
        public string m_command;
        public bool m_replyreceived;
        public string m_reply;

        /// <summary>
        /// Extracts the address of the object you are passing in?
        /// </summary>
        public long VarPtr(object e)
        {
            GCHandle GC = GCHandle.Alloc(e, GCHandleType.Pinned);
            long gc = GC.AddrOfPinnedObject().ToInt64();
            GC.Free();
            return gc;
        }

        /// <summary>
        /// Process incoming window message
        /// </summary>
        protected override void WndProc(ref Message m)
        {

            //Use to read registered hotkey pressed
            if (m.Msg == WM_HOTKEY)
                handleHotKey((int)m.WParam);


            //Used to read messages from stickies
            switch (m.Msg)
            {
                case WM_COPYDATA:

                    COPYDATASTRUCT CD = (COPYDATASTRUCT)m.GetLParam(typeof(COPYDATASTRUCT));
                    byte[] B = new byte[CD.cbData];
                    IntPtr lpData = CD.lpData;
                    Marshal.Copy(lpData, B, 0, CD.cbData);
                    string strData = Encoding.Default.GetString(B);
                    if (CD.dwData == (IntPtr)m_commandref)
                    {
                        // A reply, so store it
                        m_reply = strData;
                        m_replyreceived = true;
                    }
                    else
                    {
                        // It's an event
                        AddEventLine(strData);
                    }
                    break;

                default:
                    // let the base class deal with it
                    base.WndProc(ref m);
                    break;
            }
        }




        //-------- Window Position Setup -----------//

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string className, string windowTitle);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, ShowWindowState flags);

        [DllImport("user32.dll")]
        private static extern int SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowPlacement(IntPtr hWnd, ref Windowplacement lpwndpl);

        [DllImport("user32.dll")]
        static extern int GetForegroundWindow();

        private enum ShowWindowState
        {
            Hide = 0,
            ShowNormal = 1, ShowMinimized = 2, ShowMaximized = 3,
            Maximize = 3, ShowNormalNoActivate = 4, Show = 5,
            Minimize = 6, ShowMinNoActivate = 7, ShowNoActivate = 8,
            Restore = 9, ShowDefault = 10, ForceMinimized = 11
        };

        private struct Windowplacement
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }




        //-------- HotKey Setup -----------//

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        const int MOD_ALT = 0x0001;
        const int MOD_CONTROL = 0x0002;
        const int MOD_SHIFT = 0x0004;
        const int WM_HOTKEY = 0x0312;


        private void handleHotKey(int noteId)
        {
            Log(LoggingLevel.Info, "Hot key for note ID " + noteId + " pressed");

            string title;
            string desiredLocation;


            try
            {
                //Find title of note by note ID
                var note = notesLocal.Find(x => x.Id == noteId);
                title = note.Title;
                desiredLocation = note.DesiredLocation;
            }
            catch (NullReferenceException)
            {

                if (notesLocal.Count == 0)
                {
                    Log(LoggingLevel.Low, "Stickies have not yet been loaded (Start sync first)");

                }
                else
                {
                    Log(LoggingLevel.Low, "Could not find note with ID " + noteId);

                }


                return;
            }


            Log(LoggingLevel.Debug, "Hot key for '" + title + "' pressed at " + DateTime.Now);


            string reply = SendToStickiesWithLog(String.Format("get desktop {0} handle", noteId));
            string handle = reply.Substring(4);

            Log(LoggingLevel.Debug, "Sticky " + title + " has handle: " + handle);

            //Parse handle (hWnd) of sticky process
            IntPtr wdwIntPtr = (IntPtr)int.Parse(handle);

            //Get window info
            Windowplacement placement = new Windowplacement();
            GetWindowPlacement(wdwIntPtr, ref placement);


            Log(LoggingLevel.Debug, "Sticky " + title + " had placement: " + placement.showCmd);

            string desiredSize = "720x1000";

            //Check if window is minimized
            if (placement.showCmd == 6)
            {
                //The window is minimized so we restore it
                ShowWindow(wdwIntPtr, ShowWindowState.Restore);

                Log(LoggingLevel.Debug, "Sticky " + title + " is now restored ");
            }



            reply = SendToStickiesWithLog(String.Format("get desktop {0} rolled", noteId));
            string rolled = reply.Substring(4);

            if (rolled == "1")
            {
                Log(LoggingLevel.Debug, "Sticky '" + title + "' was rolled ");
                reply = SendToStickiesWithLog(String.Format("do desktop {0} unrolled", noteId));
                Log(LoggingLevel.Info, "Sticky '" + title + "' is now unrolled ");


                reply = SendToStickiesWithLog(String.Format("set desktop {0} position {1}", noteId, desiredLocation));
                Log(LoggingLevel.Info, "Sticky '" + title + "' is set to " + desiredLocation);

            }
            else
            {

                int forground = GetForegroundWindow();

                Log(LoggingLevel.Debug, "Forground window handle is  " + forground);


                if (forground == int.Parse(handle))
                {
                    reply = SendToStickiesWithLog(String.Format("do desktop {0} rolled", noteId));
                    Log(LoggingLevel.Info, "Sticky '" + title + "' was unrolled and in the foreground");
                    Log(LoggingLevel.Info, "Sticky '" + title + "'  is now rolled ");

                }

            }

            //Bring Focus to the window
            SetForegroundWindow(wdwIntPtr);


            Log(LoggingLevel.Info, "Sticky '" + title + "' is now visible ");

            reply = SendToStickiesWithLog(String.Format("set desktop {0} size {1}", noteId, desiredSize));
            Log(LoggingLevel.Info, "Sticky '" + title + "' size is set to " + desiredSize);


        }





        //-------- Logging Setup -----------//


        private LoggingLevel loggingSetting = LoggingLevel.Info;

        public enum LoggingLevel
        {
            Low = 0,
            Info = 1,
            Debug = 2,
        };

        public void Log(LoggingLevel messageLevel, string message, string variable = "")
        {
            if (messageLevel <= loggingSetting)
            {
                messagesTextBox.AppendText(message + variable + Environment.NewLine);
            }
        }
        public void Log(LoggingLevel messageLevel)
        {
            if (messageLevel <= loggingSetting)
            {
                messagesTextBox.AppendText(Environment.NewLine);
            }
        }

        public void Log(string message)
        {
            if (LoggingLevel.Info <= loggingSetting)
            {
                messagesTextBox.AppendText(message + Environment.NewLine);
            }
        }

        public void AddCommandLines(string one, string two)
        {
            listBox1.Items.Add(one);
            listBox2.Items.Add(two);

            listBox1.TopIndex = listBox1.Items.Count - (listBox1.ClientSize.Height / listBox1.ItemHeight) + 1;
            listBox2.TopIndex = listBox2.Items.Count - (listBox2.ClientSize.Height / listBox2.ItemHeight) + 1;

        }

        public void AddEventLine(string three)
        {
            listBox3.Items.Add(three);

            listBox3.TopIndex = listBox3.Items.Count - (listBox3.ClientSize.Height / listBox3.ItemHeight) + 1;
        }



        //-------- App Config Setup -----------//

        public bool syncing = false;
        public string stickiesDB;

        public DateTime stickiesDbLastModified;
        List<Note> notesLocal = new List<Note>();
        List<Note> notesDB = new List<Note>();


        int waitDelay = 1000;
        string user = "Jack";

        Form importDialog = new Form();

        //-------- App -----------//


        public MainWindow()
        {
            InitializeComponent();

        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            RegisterHotKey(this.Handle, 1, MOD_ALT, (int)Keys.Q);
            RegisterHotKey(this.Handle, 2, MOD_ALT, (int)Keys.W);
            RegisterHotKey(this.Handle, 14, MOD_ALT, (int)Keys.E);

            //Populate dropdown comboBox

            foreach (string level in Enum.GetNames(typeof(LoggingLevel)))
            {
                loggingLevelDropdown.Items.Add(level);
            }

            loggingLevelDropdown.SelectedItem = loggingSetting.ToString();

            // Stickies Setup (Should Add each open note)
            notesLocal.Add(new Note(1, "110, 190")); // Add "To Do"
            notesLocal.Add(new Note(2, "940, 190")); // Add "Notes"
            notesLocal.Add(new Note(14, "1750, 190")); // Add "Projects"


            foreach (var note in notesLocal)
            {
                UpdateNoteObject(note);
            }

            Log(LoggingLevel.Low, "Found " + notesLocal.Count + " open stickies");
            Log(LoggingLevel.Info);

            RtfTest();

        }


        public string SendToStickiesWithLog(string message)
        {
            string reply = SendToStickies(message);

            if (reply.Length == 0)
                AddCommandLines(message, "ERROR");
            else
            {
                AddCommandLines(message, reply);
            }

            return reply;
        }


        public string SendToStickies(string str)
        {
            int hWnd = 0;
            hWnd = FindWindowEx(0, hWnd, null, "ZhornSoftwareStickiesMain");

            if (hWnd == 0)
                return "Stickies Not Found";      // stickies not found

            else
            {
                // Add API code
                str = "api " + str;

                // Generate a unique number for this command
                m_commandref++;

                // Reset flag
                m_replyreceived = false;

                // Send the message to Stickies
                IntPtr ptr = (IntPtr)Marshal.StringToHGlobalAnsi(str);
                COPYDATASTRUCT cs = new COPYDATASTRUCT();
                cs.dwData = (IntPtr)m_commandref;
                cs.lpData = ptr;
                cs.cbData = str.Length + 1;
                int ret = ZSendMessage(hWnd, WM_COPYDATA, (int)Handle, VarPtr(cs));


                // Wait for a reply for a second
                int count = 0;
                while (count < 20)
                {
                    System.Threading.Thread.Sleep(50);
                    if (m_replyreceived)
                    {
                        // Free memory
                        Marshal.FreeHGlobal(ptr);

                        return m_reply;
                    }
                    count++;
                }

                // Free memory
                Marshal.FreeHGlobal(ptr);

                // We timed out so return nothing to indicate an error
                return "Timed Out!";
            }
        }


        //-------- Button Clicks -----------//


        private void sendButton_Click(object sender, EventArgs e)
        {
            commandTextBox.SelectAll();
            String str = commandTextBox.SelectedText;

            string reply = SendToStickiesWithLog(str);

        }


        private void startSync_Click(object sender, EventArgs e)
        {
            //Start sync
            Log(LoggingLevel.Low, "Starting auto sync | " + DateTime.Now.ToString("MMM dd - h:mm tt"));
            syncing = true;


            // Get DB location
            string reply = SendToStickiesWithLog("get datafile");
            stickiesDB = reply.Substring(4);
            Log(LoggingLevel.Info, "Automatically set stickied DB to ", stickiesDB);
            Log(LoggingLevel.Info);


            stickiesDbLastModified = File.GetLastWriteTime(stickiesDB);

            //Looping Sync
            Sync();

        }

        void UpdateNoteObject(Note note)
        {



            String reply = SendToStickiesWithLog("get desktop " + note.Id + " title");
            string title = reply.Substring(4);
            Log(LoggingLevel.Info, "Note " + note.Id + " has title: " + title);
            note.Title = title;

            reply = SendToStickiesWithLog("get desktop " + note.Id + " colour");
            string colour = reply.Substring(4);
            Log(LoggingLevel.Info, "Note " + note.Id + " has colour: " + colour);
            note.Colour = colour;

            reply = SendToStickiesWithLog("get desktop " + note.Id + " modified");
            string modified = reply.Substring(4);
            Log(LoggingLevel.Debug, "Note " + note.Id + " has modified: " + modified);
            string[] dateParts = modified.Split(' ');
            DateTime modifiedDate = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            modifiedDate = modifiedDate.AddSeconds(Convert.ToInt64(dateParts[0])).ToLocalTime();
            Log(LoggingLevel.Info, "Parsed 'last modified' date for note " + note.Id + " to: " + modifiedDate);
            note.LastModified = modifiedDate;

            reply = SendToStickiesWithLog("get desktop " + note.Id + " rtf");
            string rtf = reply.Substring(4);
            Log(LoggingLevel.Debug, "Note " + note.Id + " has rtf: " + rtf);
            note.Rtf = rtf;
            Log(LoggingLevel.Info);
        }

        private void commandTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                sendButton.PerformClick();
            }
        }

        private void stopSync_Click(object sender, EventArgs e)
        {
            if (syncing)
            {
                Log(LoggingLevel.Low, "Stopping auto sync | " + DateTime.Now.ToString("MMM dd - h:mm tt"));

            }
            else
            {
                Log(LoggingLevel.Low, "Auto sync is not running");

            }

            syncing = false;
        }

        private void eventsCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (eventsCheckBox.Checked)
            {
                SendToStickiesWithLog("do register");

                Log(LoggingLevel.Info, "Started getting stickies events  | " + DateTime.Now.ToString("MMM dd - h:mm tt"));

            }
            else
            {
                SendToStickiesWithLog("do deregister");

                Log(LoggingLevel.Info, "Stopped getting stickies events  | " + DateTime.Now.ToString("MMM dd - h:mm tt"));
            }
        }


        private void clearEvents_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
        }



        private void loggingDropDown_SelectedIndexChanged(object sender, EventArgs e)
        {

            Log(LoggingLevel.Debug, "Chose: " + loggingLevelDropdown.Text);

            LoggingLevel selected = (LoggingLevel)Enum.Parse(typeof(LoggingLevel), loggingLevelDropdown.Text);


            loggingSetting = selected;
            Log(LoggingLevel.Low, "Logging level is set to: " + loggingSetting);

        }

        //-------- App Functions -----------//


        public async void Sync()
        {
            while (syncing)
            {

                DateTime currentStickiesLastModified = File.GetLastWriteTime(stickiesDB);

                if (stickiesDbLastModified != currentStickiesLastModified)
                {

                    Log(LoggingLevel.Info, "Stickies DB was recently modified: " +
                            currentStickiesLastModified.ToString("dd/MMM/yy HH:mm:ss"));

                    stickiesDbLastModified = currentStickiesLastModified;

                    foreach (var note in notesLocal)
                    {
                        string reply = SendToStickiesWithLog("get desktop " + note.Id + " modified");
                        string modified = reply.Substring(4);
                        string[] dateParts = modified.Split(' ');
                        DateTime currentmodifiedDate = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
                        currentmodifiedDate = currentmodifiedDate.AddSeconds(Convert.ToInt64(dateParts[0])).ToLocalTime();

                        if (currentmodifiedDate != note.LastModified)
                        {
                            Log(LoggingLevel.Debug);
                            Log(LoggingLevel.Debug, "note.LastModified for " + note.Title + " is: " + note.LastModified);
                            Log(LoggingLevel.Debug);


                            int seconds = Convert.ToInt32((DateTime.Now - currentmodifiedDate).TotalSeconds);

                            Log(LoggingLevel.Debug, "'" + note.Title + "' was modified " + seconds + " seconds ago");


                            Log(LoggingLevel.Info, "'" + note.Title + "' was modified at " + currentmodifiedDate.ToString("h:mm tt"));
                            Log(LoggingLevel.Info);

                            Log(LoggingLevel.Info, "Updating saved note properties for '" + note.Title + "'");

                            UpdateNoteObject(note);


                            note.LastModified = currentmodifiedDate;

                            //Update DB
                            SaveNote(note);
                        }

                    }


                }

                await Task.Delay(waitDelay);

            }
        }

        /// <summary>
        /// Save note to DB
        /// </summary>
        void SaveNote(Note note)
        {


            DBConnection connection = new DBConnection();

            var html = RtfPipe.Rtf.ToHtml(note.Rtf);
            html = html.Replace("’", "'")
                       .Replace("–", "&ndash;")
                       .Replace("'", "\\'");

            note.LastModifiedBy = "NotesSyncApp";

            string json = JsonConvert.SerializeObject(note);

            Log(LoggingLevel.Debug);
            Log(LoggingLevel.Debug, "Compiled JSON: " + json);
            Log(LoggingLevel.Debug);


            string data = "notSet";
            string info = "notSet";

            if (note.Title == "To Do")
            {
                data = "notesData1";
                info = "notesInfo1";
            }

            if (note.Title == "Notes")
            {
                data = "notesData2";
                info = "notesInfo2";
            }

            if (note.Title == "Projects")
            {
                data = "notesData3";
                info = "notesInfo3";
            }


            try
            {

                Log(LoggingLevel.Info, "Connecting to DB for upload");

                string sql = "UPDATE xMainData SET " +
                    data + " = '" + html + "'," +
                    info + " = " + "'" + json + "'" +


                    " WHERE " +
                     "User = " + "'" + user + "'";
                Log(LoggingLevel.Debug);
                Log(LoggingLevel.Debug, "--------------- SQL START ---------------");
                Log(LoggingLevel.Debug, sql);
                Log(LoggingLevel.Debug, "--------------- SQL END ---------------");
                Log(LoggingLevel.Debug);

                MySqlCommand cmd = new MySqlCommand(sql, connection.Connection);

                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Log(LoggingLevel.Low, "Failed to connect: " + ex.Message);

            }

            connection.Close();

            Log(LoggingLevel.Low, "'" + note.Title + "' was uploaded to DB sucessfully at " + DateTime.Now.ToString("h:mm tt"));

        }

        List<Note> GetNotes()
        {
            DBConnection connection = new DBConnection();
            List<Note> notesOnX = new List<Note>();

            int notesOnXCount = 3;

            Log(LoggingLevel.Info, "Connecting to DB for all notes download");


            for (int i = 1; i <= notesOnXCount; i++)
            {


                try
                {

                    //Get note properties

                    string query = "SELECT notesInfo" + i + " FROM xMainData WHERE User = '" + user + "'";

                    var cmd = new MySqlCommand(query, connection.Connection);
                    var reader = cmd.ExecuteReader();

                    Note note = new Note();


                    while (reader.Read())
                    {
                        string noteInfo = reader.GetString(0);

                        Log(LoggingLevel.Info, "Note info: ", noteInfo);

                        note = JsonConvert.DeserializeObject<Note>(noteInfo);

                        Log(LoggingLevel.Info, "Note has title ", note.Title);
                    }

                    reader.Close();


                    //Get note contents

                    query = "SELECT notesData" + i + " FROM xMainData WHERE User = '" + user + "'";

                    cmd = new MySqlCommand(query, connection.Connection);
                    reader = cmd.ExecuteReader();



                    while (reader.Read())
                    {
                        string noteRTF = reader.GetString(0);

                        Log(LoggingLevel.Info, "Found note of length: ", noteRTF.Length.ToString());
                        Log(LoggingLevel.Debug, "Found contents: ", noteRTF);
                        note.Rtf = noteRTF;
                        note.Length = noteRTF.Length;

                    }

                    reader.Close();


                    notesOnX.Add(note);




                }
                catch (Exception ex)
                {
                    messagesTextBox.AppendText("Failed to connect: " + ex + Environment.NewLine);
                }

            }

            connection.Close();


            return notesOnX;


        }

        private void downloadX_Click(object sender, EventArgs e)
        {
            // Get list of notes saved in DB
            notesDB = GetNotes();


            // Show the notes and choose whichout to download
            Panel panel = new Panel();

            DataGridView grid = new DataGridView();
            panel.Controls.Add(grid);
            grid.Dock = DockStyle.Top;


            panel.BackColor = System.Drawing.Color.FromName("Red");
            panel.Dock = DockStyle.Fill;

            importDialog.Controls.Add(panel);
            importDialog.Size = new Size(1300, 300);
            importDialog.Show(this);


            DataTable dt = new DataTable();
            dt.Columns.Add("Title", typeof(string));
            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("Length", typeof(string));
            dt.Columns.Add("Colour", typeof(string));
            dt.Columns.Add("Last Modified", typeof(DateTime));
            dt.Columns.Add("Last Modified By", typeof(string));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("Take Action", typeof(bool));



            List<List<string>> allActions = new List<List<string>>();


            // See if they already exist in stickies (by title)
            foreach (Note noteDB in notesDB)
            {
                Log("Analyzing note from DB: " + noteDB.Title);

                int foundNoteLocal = notesLocal.FindIndex(x => x.Title == noteDB.Title);

                string status = "Not Set";
                string action;
                bool download = true;

                if (foundNoteLocal > -1)
                {
                    int foundID = notesLocal[foundNoteLocal].Id;
                    action = "Replace '" + noteDB.Title + "' (ID " + foundID + ")";

                    Log("DB Last Modified Time: " + noteDB.LastModified);
                    Log("Local Last Modified Time: " + notesLocal[foundNoteLocal].LastModified);


                    if (notesLocal[foundNoteLocal].LastModified > noteDB.LastModified)
                    {
                        status = "DB is older";
                        download = false;
                    }
                    else if (notesLocal[foundNoteLocal].LastModified == noteDB.LastModified)
                    {
                        status = "DB is same";
                        download = false;
                    }
                    else
                    {
                        status = "DB is newer";
                        download = true;
                    }
                }
                else
                {
                    action = "Create new sticky '" + noteDB.Title + "'";
                }

                List<string> rowActions = new List<string>();



                rowActions.Add(action);
                rowActions.Add("Create New");
                rowActions.Add("Replace Existing");

                allActions.Add(rowActions);



                dt.Rows.Add(new object[] { noteDB.Title, noteDB.Id, noteDB.Length, noteDB.Colour, noteDB.LastModified, noteDB.LastModifiedBy, status, download });

            }

            grid.DataSource = dt;

            DataGridViewComboBoxColumn actionsColumn = new DataGridViewComboBoxColumn();
            grid.Columns.Add(actionsColumn);

            actionsColumn.HeaderText = "Action";
            actionsColumn.DataPropertyName = "Action";
            actionsColumn.FlatStyle = FlatStyle.Flat;

           




            int actionsColumnNumber = 8;
            grid.Columns.Add("Title", "Title");
            grid.Columns.Add("ID", "ID");


            int count = 0;
            foreach (List<string> rowActions in allActions)
            {
                DataGridViewComboBoxCell cell = (DataGridViewComboBoxCell)grid.Rows[count].Cells[actionsColumnNumber];

                // You might pass a boolean to determine whether to clear or not.
                //cell.Items.Clear();

                foreach (object itemToAdd in rowActions)
                {
                    cell.Items.Add(itemToAdd);
                }

                

                //White background
                grid.Rows[count].Cells[actionsColumnNumber].Style.BackColor = Color.White;

                cell.Value = rowActions[0];

                count++;
            }





            grid.DefaultValuesNeeded += (sender, EventArgs) => { OnGridDefaultValuesNeeded(sender, EventArgs); };




            void OnGridDefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
            {

                    MessageBox.Show("doin");
                    e.Row.Cells[actionsColumnNumber].Value = "DropDown";

            }


            grid.CellClick += (sender, EventArgs) => { dataGridView_CellClick(sender, EventArgs); };






            grid.AllowUserToAddRows = false;
            grid.ReadOnly = false;
            grid.Columns[1].ReadOnly = true;
            grid.Columns[2].ReadOnly = true;
            grid.Columns[4].ReadOnly = true;
            grid.Columns[5].ReadOnly = true;
            grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            grid.Columns[actionsColumnNumber].Width = 200;

            Button takeActionButton = new Button();

            takeActionButton.Location = new Point(400, 200);
            takeActionButton.Text = "Download";
            takeActionButton.AutoSize = true;
            takeActionButton.BackColor = Color.LightBlue;
            takeActionButton.Padding = new Padding(6);
            takeActionButton.Click += (sender, EventArgs) => { takeActionButton_Click(sender, EventArgs, dt); };

            panel.Controls.Add(takeActionButton);

            foreach (DataGridViewRow row in grid.Rows)
            {
                changeCell(row.Cells[actionsColumnNumber + 1], false);
                changeCell(row.Cells[actionsColumnNumber + 2], false);


            }

            grid.EditingControlShowing += (sender, EventArgs) => { Grid_EditingControlShowing(sender, EventArgs); };


            void Grid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
            {

    
                ComboBox cb = e.Control as ComboBox;
                if (cb != null && grid.CurrentCell.ColumnIndex == actionsColumnNumber)
                {
                    cb.SelectedValueChanged -= new EventHandler(cb_SelectedIndexChanged);
                    cb.SelectedValueChanged += new EventHandler(cb_SelectedIndexChanged);
                    
                }




            }

            void cb_SelectedIndexChanged(object sender, EventArgs e)
            {

                ComboBox box = (ComboBox)sender;


                //DataGridViewRow row = (DataGridViewRow)box.Name;

                if (box.SelectedIndex == 0)
                {
                    changeCell(grid.Rows[grid.CurrentCell.RowIndex].Cells[actionsColumnNumber + 1], false);
                    changeCell(grid.Rows[grid.CurrentCell.RowIndex].Cells[actionsColumnNumber + 2], false);


                }
                else if (box.SelectedIndex == 1)
                {
                    changeCell(grid.Rows[grid.CurrentCell.RowIndex].Cells[actionsColumnNumber + 1], true);
                    changeCell(grid.Rows[grid.CurrentCell.RowIndex].Cells[actionsColumnNumber + 2], false);


                }
                else if (box.SelectedIndex == 2)
                {
                    changeCell(grid.Rows[grid.CurrentCell.RowIndex].Cells[actionsColumnNumber + 1], false);
                    changeCell(grid.Rows[grid.CurrentCell.RowIndex].Cells[actionsColumnNumber + 2], true);


                }
            }



            void changeCell(DataGridViewCell dc, bool enabled)
            {
                //toggle read-only state
                dc.ReadOnly = !enabled;
                if (enabled)
                {
                    //restore cell style to the default value
                    dc.Style.BackColor = dc.OwningColumn.DefaultCellStyle.BackColor;
                    dc.Style.ForeColor = dc.OwningColumn.DefaultCellStyle.ForeColor;
                }
                else
                {
                    //gray out the cell
                    dc.Style.BackColor = Color.LightGray;
                    dc.Style.ForeColor = Color.DarkGray;
                }
            }

            void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
            {
                if (e.ColumnIndex == actionsColumnNumber)
                {
                    grid.BeginEdit(true);

                    (grid.EditingControl as DataGridViewComboBoxEditingControl).DroppedDown = true;

                }

                //        If DataGridView1.Rows(e.RowIndex).Cells(ddl.Name).Selected = True Then


                /*var editingControl = grid.EditingControl as
                    DataGridViewComboBoxEditingControl;
                if (editingControl != null)
                    editingControl.DroppedDown = true;*/
            }



        }



        private void takeActionButton_Click(object sender, EventArgs e, DataTable dt)
        {


            foreach (DataRow row in dt.Rows)
            {
                if ((bool)row["Take Action"])
                {
                    Log("Taking action for " + row["Title"]);


                    Note note = notesDB.Find(x => x.Id == int.Parse(row["Id"].ToString()));

                    string html = note.Rtf;

                    var rtfConverter = new RtfConverter();
                    var rtf = rtfConverter.Convert(html);

                    int noteId = 4;

                    string reply = SendToStickiesWithLog(String.Format("set desktop {0} rtf {1}", noteId, rtf));



                }
            }

            importDialog.Hide();
        }
        private void RtfTest()
        {
            string html = "<div style='font-size:12pt;font-family:Verdana;'>Line 1<p style='color:#FFF000;font-size:10pt;margin:0;'><strong><u>__________Line 2_______________</u></strong></p><p>Line3</p><p>Line4</p></div>"
                ;

            html = "<p style='color:#FFFFFF;font-size:10pt;background:#FFFFFF;margin:0;'><br></p><ul style='color:#F11111;font-size:10pt;background:#FFFFFF;margin:0 0 0 20px;padding-left:0;'><li>Fix Up <em>To Do</em></li></ul><ul style='color:#FFF111;font-size:10pt;background:#FFFFFF;margin:0 0 0 20px;padding-left:0;'><li>Fix Up <em>To Do</em></li><li>44</li></ul><p>Line 3</p>";
            Note note = notesLocal.Find(x => x.Id == 1);

            html = "<p>Line Before</p><ul style='font-size:10pt;background:#FFFFFF;color:#FF0000;margin:0 0 0 20px;padding-left:0;'><li>Buy <a style='text-decoration:none;color:#FF0000;' href='https://www.homehardware.ca/en/charcoal-range-hood-filter-for-model-rl-and-sm/p/3855233'>https://www.homehardware.ca/en/charcoal-range-hood-filter-for-model-rl-and-sm/p/3855233</a></li></ul><p>Line After</p>";
            //html = note.Rtf;

            html = "<p style='background:#FF1111;font-size:10pt;color:#00FFFF;font-family:Verdana, sans-serif;margin:0;'>- Repair screw hole in tire (Make appointment or DIY)</p>";

            html = "<ul style='color:#FFFFFF;font-size:10pt;background:#FFFFFF;font-family:&quot;Segoe Print&quot;;margin:0 0 0 20px;padding-left:0;'><li>Fix Up <em>To Do</em></li></ul><p style='color:#FFFFFF;font-size:10pt;background:#FFFFFF;margin:0;'><br></p><p style='color:#FFFFFF;font-size:10pt;background:#FFFFFF;margin:0;'> &nbsp;</p><ul style='color:#FFFFFF;font-size:10pt;background:#FFFFFF;font-family:&quot;Arial Black&quot;;margin:0 0 0 20px;padding-left:0;'><li>Jill (no new)</li><li style='font-family:Verdana;'>10</li>";

            html = @"<p style='color:#FFFFFF;font-size:10pt;background:#FFFFFF;margin:0;'>\</p>";

            html = "	<p><br></p>" +
                "<p></p>" +
                "<p>Hello</p>" +
                "<p>Bye</p>";


            html = @"<ul style='color:#FFFFFF;font-size:10pt;margin:0 0 0 20px;padding-left:0;'><li>One</li>";


            html = @"<ul style='color:#ffffff; font-size:10pt; margin:0 0 0 20px; padding-left:0'><li>One</li><li>Two</li><li>&nbsp;&nbsp;Three</li></ul>";


            var rtfConverter = new RtfConverter();
            var rtf = rtfConverter.Convert(html);
            System.Diagnostics.Debug.WriteLine(rtf);
            int noteId = 4;

            string reply = SendToStickiesWithLog(String.Format("set desktop {0} rtf {1}", noteId, rtf));
        }


    }
}
