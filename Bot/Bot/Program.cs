using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text;
using Microsoft.Lync.Model;
using Microsoft.Lync.Model.Extensibility;
using Microsoft.Lync.Model.Conversation;
using System.Threading;
using System.Data;
using MySql.Data.MySqlClient;

namespace Bot
{
    public partial class Form1 : Form
    {
        Contact RemoteContact;
        Boolean Start = false;
        Boolean pop = false;
        ThreadStart th;
        Thread t;
        LyncClient lyncClient;
        ContactAvailability currentAvailability = 0;

        private void _contact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            currentAvailability = (ContactAvailability)RemoteContact.GetContactInformation(ContactInformationType.Availability);
            switch (currentAvailability)
            {
                case ContactAvailability.Away:
                    messageshow("노란불", MessageBoxIcon.Question);
                    break;
                case ContactAvailability.Busy:
                    messageshow("다른용무중", MessageBoxIcon.Exclamation);
                    break;
                case ContactAvailability.BusyIdle:
                    messageshow("다른용무중", MessageBoxIcon.Exclamation);
                    break;
                case ContactAvailability.DoNotDisturb:
                    messageshow("방해금지", MessageBoxIcon.Information);
                    break;
                case ContactAvailability.Free:
                    messageshow("업무중", MessageBoxIcon.Warning);
                    break;
                case ContactAvailability.FreeIdle:
                    messageshow("업무중", MessageBoxIcon.Warning);
                    break;
                case ContactAvailability.Offline:
                    messageshow("오프라인", MessageBoxIcon.Asterisk);
                    break;
                case ContactAvailability.TemporarilyAway:
                    messageshow("곧 돌아오겠음", MessageBoxIcon.Warning);
                    break;
                default:
                    break;
            }
        }
        private void messageshow(string watch, MessageBoxIcon ico)
        {
            if (!pop)
            {
                MessageBox.Show(watch + "(으)로바뀜 삐용삐용!!",
                                    "주의!!",
                                     MessageBoxButtons.OK,
                                     ico,
                                     MessageBoxDefaultButton.Button2,
                                     MessageBoxOptions.DefaultDesktopOnly);
                pop = true;
            }
            else pop = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!Start && textBox2.Text != "")
            {
                th = new ThreadStart(flash);
                t = new Thread(th);
                t.Start();
                button1.BackColor = System.Drawing.Color.Red;
                Start = true;
            }
            else
            {
                Start = false;
                RemoteContact.ContactInformationChanged -= _contact_ContactInformationChanged;
                if (t != null) t.Abort();
                button1.BackColor = System.Drawing.Color.White;
            }
        }

        private void flash()
        {
            lyncClient = LyncClient.GetClient();
            RemoteContact = lyncClient.ContactManager.GetContactByUri(textBox2.Text);
            RemoteContact.ContactInformationChanged += _contact_ContactInformationChanged;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load_1);
            this.ResumeLayout(false);

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
    }
    static class Program
    {
        static LyncClient Client;
        static Thread t1;
        static String strConn;
        static MySqlConnection conn;
        static IDictionary<InstantMessageContentType, string> messageDictionary;
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Client = LyncClient.GetClient();
            strConn = "Server=localhost;Database=mysql;Uid=root;Pwd=;";
            conn = new MySqlConnection(strConn);
            t1 = new Thread(new ThreadStart(Run));
            t1.Start();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        static void Run()
        {
            Client.ConversationManager.ConversationAdded += ConversationManager_ConversationAdded;
            //Client.ConversationManager.ConversationRemoved += ConversationManager_ConversationRemoved;
        }
        
        private static void ConversationManager_ConversationAdded(object sender, ConversationManagerEventArgs e)
        {
            //makeMessage();
            //MessageBox.Show("첫대화");
            e.Conversation.ParticipantAdded += Conversation_ParticipantAdded;
        }

        private static void Conversation_ParticipantAdded(object sender, ParticipantCollectionChangedEventArgs e)
        {
            var instantMessageModality = e.Participant.Modalities[ModalityTypes.InstantMessage] as InstantMessageModality;
            instantMessageModality.InstantMessageReceived += InstantMessageModality_InstantMessageReceived;
        }

        private static void InstantMessageModality_InstantMessageReceived(object sender, MessageSentEventArgs e)
        {
            InstantMessageModality test1 = sender as InstantMessageModality;

            String txt = e.Text.Trim();

            int leng = txt.Length;
            if (txt == "!퇴근")
            {
                makeMessage(1,test1,null);
            }
            else if (txt == "!주사위")
            {
                makeMessage(2, test1,null);
            }
            else if (txt == "!아침")
            {
                makeMessage(4, test1, "AM");
            }
            else if (txt == "!점심")
            {
                makeMessage(4, test1, "NOON");
            }
            else if (txt == "!저녁")
            {
                makeMessage(4, test1, "PM");
            }
            else if (txt.StartsWith("!추가 "))
            {
                string add="";
                try { 
                    add = txt.Substring(4, leng-4);
                }
                catch (Exception)
                {
                    add = txt.Substring(4, leng - 5);
                }
                makeMessage(3, test1,add);
            }
            else if(txt.Substring(0,1) == "!")
            {
                makeMessage(0, test1, txt.Substring(1,leng-1));
            }
            else
            {

            }
        }

        private static void ConversationManager_ConversationRemoved(object sender, ConversationManagerEventArgs e)
        {
            MessageBox.Show("삭제");
        }

        public static void makeMessage(int flag,InstantMessageModality insM, string add)
        {
            Conversation conv = insM.Conversation;
            string Message = "";
            Random r = new Random();

            if (flag == 0)
            {
                string[] arr = add.Split(' ');
                try
                {
                    conn.Open();
                    DataSet ds = new DataSet();
                    String sql;
                    sql = "SELECT res FROM bot where req = '" + arr[0] + "'";

                    MySqlDataAdapter adpt = new MySqlDataAdapter(sql, conn);
                    adpt.Fill(ds,"bot");
                    try
                    {   
                        Message = ds.Tables[0].Rows[r.Next(ds.Tables[0].Rows.Count)].ItemArray[0].ToString(); // res 가 2개일때 처리
                        
                    }
                    catch
                    {
                        Message = "알 수 없는 명령어 입니다. \r\n\r\n EX) !추가 [가르칠 말] [반응할 말]";
                    }
                    conn.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.StackTrace);
                    conn.Close();
                }

                
            }
            if (flag == 1)
            {
                int a = 17 - Convert.ToInt32(DateTime.Now.ToString("HH"));
                int b = 59 - Convert.ToInt32(DateTime.Now.ToString("mm"));
                int c = 59 - Convert.ToInt32(DateTime.Now.ToString("ss"));
                Message = "퇴근시간까지 " + a.ToString() + "시간 " + b.ToString() + "분 " + c.ToString() + "초 남았습니다.";
            }
            
            if (flag == 2)
            {
                Message = "1~100 을 굴려 " + r.Next(1, 100) + "가 나왔습니다.";
            }
            if (flag == 3)
            {
                string[] arr = add.Split(' ');
                string Carr="";

                for (int i = 1; i < arr.Length; i++)
                {
                    Carr += arr[i] + " ";
                }

                conn.Open();
                String sql = "INSERT INTO bot (req, res) VALUES ('" + arr[0] + "', '"+Carr+"')";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
                Message = "알림 : !" + arr[0] + " 추가 되었습니다.";
            }
            if (flag == 4)
            {
                Meal SJ = new Meal();
                string meal = SJ.parse(add);
                meal = meal.Replace("한식(쌀밥)", "\r\n\r\n<DIV style = 'color: green;font:bold; font-size:20p;'> 한식(쌀밥)</DIV>\r\n\r\n");
                meal = meal.Replace("한식(잡곡밥)", "\r\n\r\n<DIV style = 'color: green;font:bold; font-size:20p;'> 한식(잡곡밥)</DIV>\r\n\r\n");
                meal = meal.Replace("간편식", "\r\n\r\n<DIV style = 'color: green;font:bold; font-size:20p;'> 간편식</DIV>\r\n\r\n");
                meal = meal.Replace("해장국", "\r\n\r\n<DIV style = 'color: green;font:bold; font-size:20p;'> 해장국</DIV>\r\n\r\n");
                meal = meal.Replace("분식", "\r\n\r\n<DIV style = 'color: green;font:bold; font-size:20p;'> 분식</DIV>\r\n\r\n");
                meal = meal.Replace("건강식", "\r\n\r\n<DIV style = 'color: green;font:bold; font-size:20p;'> 건강식</DIV>\r\n\r\n");
                meal = meal.Replace("선택코너", "\r\n\r\n<DIV style = 'color: green;font:bold; font-size:20p;'> 선택코너</DIV>\r\n\r\n");
                meal = meal.Replace("일품식", "\r\n\r\n<DIV style = 'color: green;font:bold; font-size:20p;'> 일품식</DIV>\r\n\r\n");
                messageDictionary = new Dictionary<InstantMessageContentType, string>();
                messageDictionary.Add(InstantMessageContentType.Html, meal);
                goto SENDSTART;
            }

            string FormattedMessage = "<DIV style='color: green;font:bold; font-size:20p;'>" + Message + "</DIV>";
            messageDictionary = new Dictionary<InstantMessageContentType, string>();
            messageDictionary.Add(InstantMessageContentType.Html, FormattedMessage);

            SENDSTART:
            insM.BeginSendMessage(messageDictionary, ar=>
            {
                try
                {
                    insM.EndSendMessage(ar);
                }
                catch (Exception ex) { }
            }, null);
        }
    }
}
