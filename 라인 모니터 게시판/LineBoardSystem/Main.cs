using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Media;
using Infragistics.Win.UltraWinGrid;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;


namespace LineBoardSys
{
    public partial class Form1 : Form
    {

        string strConn = @"Data Source=192.168.3.30,21778;Initial Catalog=samjin;Persist Security Info=True;User ID=sa;Password=$73J0701;";

        // [추가] 설정 저장 파일 경로
        //string settingsFilePath = "last_used.txt";
        string settingsFolderPath;
        string settingsFilePath;
        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized; //최대화
            settingsFolderPath = @"C:\SAMJIN";
            settingsFilePath = Path.Combine(settingsFolderPath, "last_used.txt");
            comboBox1.Text = "천안공장";
            comboBox2.Text = "포밍";
            comboBox3.Text = "1조";

            // [추가] 설정값 불러오기
            LoadSettings();

            // [추가] 폼 종료 시 저장 메서드 연결
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
            
        }


        // [추가] 설정 저장
        private void SaveSettings()
        {
            // 디렉토리 없으면 생성
            if (!Directory.Exists(settingsFolderPath))
            {
                Directory.CreateDirectory(settingsFolderPath);
            }

            using (StreamWriter writer = new StreamWriter(settingsFilePath))
            {
                writer.WriteLine(comboBox1.Text);
                writer.WriteLine(comboBox2.Text);
                writer.WriteLine(comboBox3.Text);
            }
        }

        // [추가] 설정 불러오기
        private void LoadSettings()
        {
            if (File.Exists(settingsFilePath))
            {
                string[] lines = File.ReadAllLines(settingsFilePath);
                if (lines.Length >= 3)
                {
                    comboBox1.Text = lines[0];
                    comboBox2.Text = lines[1];
                    comboBox3.Text = lines[2];
                }
            }
        }

        // [추가] 폼 종료 시 호출
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings(); // 설정 저장
        }




        private void Monitor()
        {
            GetData1();
            GetData2();
            GetData3();

            {

                dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                dataGridView1.Columns[0].HeaderText = "공장";
                dataGridView1.Columns[0].Visible = false;

                //dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                //dataGridView1.Columns[1].HeaderText = "라인";
                //dataGridView1.Columns[1].Visible = false;

                dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                dataGridView1.Columns[1].HeaderText = "구분";
                dataGridView1.Columns[1].Width = 90;
                dataGridView1.Columns[1].Visible = false;
                //dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                dataGridView1.Columns[2].HeaderText = "품목";
                dataGridView1.Columns[2].Width = 180;


                dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                dataGridView1.Columns[3].HeaderText = "품명";
                dataGridView1.Columns[3].Width = 260;

                dataGridView1.Columns[4].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                dataGridView1.Columns[4].HeaderText = "규격";
                dataGridView1.Columns[4].Width = 260;


                dataGridView1.Columns[5].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                dataGridView1.Columns[5].HeaderText = "사양";
                dataGridView1.Columns[5].Width = 260;


                dataGridView1.Columns[6].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                dataGridView1.Columns[6].HeaderText = "재질";
                dataGridView1.Columns[6].Width = 260;


                dataGridView1.Columns[7].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                dataGridView1.Columns[7].HeaderText = "긴급품 요청사항";
                dataGridView1.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; // 나머지 셀 채우기










                    return;
             
            }
        }
        public DataTable TransposeDataTable(DataTable dt)
        {
            DataTable newTable = new DataTable();

            newTable.Columns.Add("");

           
            for (int i = 0; i < dt.Rows.Count; i++) // 열 추가
            {
                newTable.Columns.Add("공지사항");
            }

            
            for (int i = 0; i < dt.Columns.Count; i++) // 열 -> 행으로 변환
            {
                DataRow newRow = newTable.NewRow();

                
                newRow[0] = (i + 1).ToString();  // 행 카운트 1,2,3

               
                for (int j = 0; j < dt.Rows.Count; j++) // 기존 데이터 입력
                {
                    newRow[j + 1] = dt.Rows[j][i];
                }

                newTable.Rows.Add(newRow);
            }

            return newTable;
        }

        private DataSet ds = new LineBoardSys.DataSet1();


        public class global
        {
            //public static string chk3 = "1";  

            public static string chk1;
            public static string chk2;
            public static string chk3;
            public static string chk4;
        }


        private void GetData1()
        {
            using (SqlConnection conn = new SqlConnection(strConn))
            {
                conn.Open();


                if (comboBox1.Text == "천안공장")
                { global.chk1 = "A"; }
                else if (comboBox1.Text == "울산공장")
                { global.chk1 = "B"; }
                else if (comboBox1.Text == "전주공장")
                { global.chk1 = "C"; }

                if (comboBox2.Text == "포밍")
                { global.chk2 = "F"; }
                else if (comboBox2.Text == "태핑")
                { global.chk2 = "T"; }

                if (comboBox3.Text == "1조")
                { global.chk3 = "1"; }
                else if (comboBox3.Text == "2조")
                { global.chk3 = "2"; }
                else if (comboBox3.Text == "3조")
                { global.chk3 = "3"; }
                else if (comboBox3.Text == "4조")
                { global.chk3 = "4"; }
                else if (comboBox3.Text == "5조")
                { global.chk3 = "5"; }
                else if (comboBox3.Text == "6조")
                { global.chk3 = "6"; }
                else if (comboBox3.Text == "7조")
                { global.chk3 = "7"; }
                else if (comboBox3.Text == "8조")
                { global.chk3 = "8"; }
                else if (comboBox3.Text == "9조")
                { global.chk3 = "9"; }
                else if (comboBox3.Text == "10조")
                { global.chk3 = "A0"; }



                string sql = string.Format(@"EXEC USP_PM100MA1_KO654_MONITORING '{0}','{1}','{2}'",global.chk1, global.chk2, global.chk3);
                //MessageBox.Show(sql);  // 데이터 확인
                SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);
                DataSet ds1 = new DataSet();
                adapter.Fill(ds1);

                if (ds1 == null || ds1.Tables.Count == 0)
                {
                    dataGridView1.Columns.Clear();
                }
                else
                {
                    dataGridView1.DataSource = ds1.Tables[0];




                    dataGridView1.Font = new Font("Tahoma", 20, FontStyle.Regular);  // 글꼴, 크기
                    //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // 열 크기 자동 조정


                    foreach (DataGridViewColumn col in dataGridView1.Columns)
                    {
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft; // 데이터 왼쪽 정렬
                    }


                    dataGridView1.Columns[7].HeaderCell.Style.ForeColor = Color.Red; // 헤더 색상

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells[7].Style.ForeColor = Color.Red; // 셀 색상
                    }

                }



                conn.Close();
            }
        }

        private void GetData2()
        {

            if (comboBox1.Text == "천안공장")
            { global.chk1 = "A"; }
            else if (comboBox1.Text == "울산공장")
            { global.chk1 = "B"; }
            else if (comboBox1.Text == "전주공장")
            { global.chk1 = "C"; }

            if (comboBox2.Text == "포밍")
            { global.chk2 = "F"; }
            else if (comboBox2.Text == "태핑")
            { global.chk2 = "T"; }

            if (comboBox3.Text == "1조")
            { global.chk3 = "1"; }
            else if (comboBox3.Text == "2조")
            { global.chk3 = "2"; }
            else if (comboBox3.Text == "3조")
            { global.chk3 = "3"; }
            else if (comboBox3.Text == "4조")
            { global.chk3 = "4"; }
            else if (comboBox3.Text == "5조")
            { global.chk3 = "5"; }
            else if (comboBox3.Text == "6조")
            { global.chk3 = "6"; }
            else if (comboBox3.Text == "7조")
            { global.chk3 = "7"; }
            else if (comboBox3.Text == "8조")
            { global.chk3 = "8"; }
            else if (comboBox3.Text == "9조")
            { global.chk3 = "9"; }
            else if (comboBox3.Text == "10조")
            { global.chk3 = "A0"; }


           
            using (SqlConnection conn = new SqlConnection(strConn))
            {
                conn.Open();

                string sql = string.Format(@"EXEC USP_PM100MA2_KO654_MONITORING '{0}','{1}','{2}'", global.chk1, global.chk2, global.chk3);
                SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);
                DataSet ds1 = new DataSet();
                adapter.Fill(ds1);

                if (ds1.Tables[0].Rows.Count == 0 || ds1.Tables[0].Columns.Count == 0)
                {
                    dataGridView2.DataSource = null;

                }
                else
                {
                    DataTable transposedTable = TransposeDataTable(ds1.Tables[0]);
                    dataGridView2.DataSource = transposedTable;

                    foreach (DataGridViewColumn col in dataGridView2.Columns)
                    {


                        //dataGridView2.Font = new Font("Tahoma", 20, FontStyle.Regular);
                        dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // 열 크기 자동 조정
                        dataGridView2.DefaultCellStyle.WrapMode = DataGridViewTriState.True; // 데이터 자동으로 줄 바꿈
                        dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells; // 줄바꿈으로 인한 셀크기 조정
                        dataGridView2.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 20, FontStyle.Bold);
                        dataGridView2.Columns[0].HeaderText = "";
                        dataGridView2.Columns[0].Width = 30;

                        // ✅ 1번 공지사항 색상 처리 (행 인덱스 기준 0)
                        if (dataGridView2.Rows.Count > 0)
                        {
                            DataGridViewRow firstRow = dataGridView2.Rows[0];

                            for (int i = 1; i < firstRow.Cells.Count; i++) // 0번은 인덱스 숫자니까 제외
                            {
                                firstRow.Cells[i].Style.ForeColor = Color.Red;
                                firstRow.Cells[i].Style.Font = new Font("Tahoma", 20, FontStyle.Bold); // 강조 효과
                            }
                        }

                        //col.Width = 500; // 열 크기 조정

                        //    foreach (DataGridViewRow row in dataGridView2.Rows)
                        //    {
                        //        row.Height = 50; // 행 크기 조정
                        //    }
                        //}


                    }
                }
            }
        }


        private void GetData3()
        {

            if (comboBox1.Text == "천안공장")
            { global.chk1 = "A"; }
            else if (comboBox1.Text == "울산공장")
            { global.chk1 = "B"; }
            else if (comboBox1.Text == "전주공장")
            { global.chk1 = "C"; }

            if (comboBox2.Text == "포밍")
            { global.chk2 = "F"; }
            else if (comboBox2.Text == "태핑")
            { global.chk2 = "T"; }

            if (comboBox3.Text == "1조")
            { global.chk3 = "1"; }
            else if (comboBox3.Text == "2조")
            { global.chk3 = "2"; }
            else if (comboBox3.Text == "3조")
            { global.chk3 = "3"; }
            else if (comboBox3.Text == "4조")
            { global.chk3 = "4"; }
            else if (comboBox3.Text == "5조")
            { global.chk3 = "5"; }
            else if (comboBox3.Text == "6조")
            { global.chk3 = "6"; }
            else if (comboBox3.Text == "7조")
            { global.chk3 = "7"; }
            else if (comboBox3.Text == "8조")
            { global.chk3 = "8"; }
            else if (comboBox3.Text == "9조")
            { global.chk3 = "9"; }
            else if (comboBox3.Text == "10조")
            { global.chk3 = "A0"; }




            using (SqlConnection conn = new SqlConnection(strConn))
            {
                conn.Open();

                string sql = string.Format("EXEC USP_PM100MA3_KO654_MONITORING '{0}','{1}','{2}'", global.chk1, global.chk2, global.chk3);
                SqlCommand cmd = new SqlCommand(sql, conn);

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                dataGridView3.Columns.Clear();
                dataGridView3.Rows.Clear();

                if (dt.Rows.Count > 0)
                {
                    DataGridViewImageColumn imageCol = new DataGridViewImageColumn
                    {
                        ImageLayout = DataGridViewImageCellLayout.Zoom // 이미지 셀 크기에 맞춤(AutoSize도 같이 확인해줘야함)
                    };

                    dataGridView3.Columns.Add(imageCol);

                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["IMAGE"] != DBNull.Value)
                        {
                            byte[] imgBytes = (byte[])row["IMAGE"];
                            using (MemoryStream ms = new MemoryStream(imgBytes))
                            {
                                Image img = Image.FromStream(ms);
                                dataGridView3.Rows.Add(img);
                            }
                        }
                        else
                        {
                            dataGridView3.Rows.Add((Image)null);
                        }
                    }

                    dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // 열 크기 자동
                    dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells; // 행 크기 자동
                    //dataGridView3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // 가운데 정렬

                }

                conn.Close();
            }
        }




        //private DataSet GetTeamCnt()
        //{
        //    DataSet iTemp_cnt = new DataSet();

        //    SqlConnection conn = new SqlConnection(strConn);
        //    conn.Open();

        //    string sql = string.Format(@" SELECT  COUNT(*) CNT FROM AD_MSG_CALL_HIST A (NOLOCK) WHERE STATUS = 'N' GROUP BY RECEIVER_GRP_NM");

        //    // SqlDataAdapter 초기화
        //    SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);

        //    // Fill 메서드 실행하여 결과 DataSet을 리턴받음
        //    adapter.Fill(iTemp_cnt);

        //    ds.Tables["DEPT_CNT"].Clear();
        //    ds.Tables["DEPT_CNT"].Merge(iTemp_cnt.Tables[0], false, MissingSchemaAction.Ignore);

        //    conn.Close();
        //    return iTemp_cnt;
        //}

        private void StartMonitoring()
        {
            Monitor();
            timer1.Start();
            timer2.Start();
            timer3.Start();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            StartMonitoring();
            //MessageBox.Show("새로고침");  
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            //timer1.Stop();
            //timer2.Stop();
            //timer3.Stop();
        }


        //한글깨짐방지
        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {
                Clipboard.SetText(Clipboard.GetText(), TextDataFormat.Text); //한글깨짐방지
            }
            catch
            {
            }
        }



        // 경고음 발생
        private void Sound()
        {
            SoundPlayer error = new SoundPlayer(LineBoardSys.Properties.Resources.error);
            error.PlaySync();

            SoundPlayer tag = new SoundPlayer(LineBoardSys.Properties.Resources.TAG);
            tag.PlaySync();


        }


        // 경고음 UPDATE
        public void warring_Update(string issue_req_no, string cust_barcode)
        {
            DataSet iTemp1 = new DataSet();

            SqlConnection conn = new SqlConnection(strConn);
            conn.Open();

            string sql = string.Format(@" UPDATE pda_list_check SET out_flag = 'Y' WHERE check_yn = 'N'  and ISSUE_REQ_NO ='{0}' and CUST_BARCODE = '{1}'", issue_req_no, cust_barcode);

            SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);

            adapter.Fill(iTemp1);

            conn.Close();
        }

        protected static ScrollBars GetVisibleScrollbars(ScrollableControl ctl)
        {

            if (ctl.HorizontalScroll.Visible)

                return ctl.VerticalScroll.Visible ? ScrollBars.Both : ScrollBars.Horizontal;

            else

                return ctl.VerticalScroll.Visible ? ScrollBars.Vertical : ScrollBars.None;

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            Sound();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBox1.Text == "천안공장")
                global.chk1 = "A";
            else if (comboBox1.Text == "울산공장")
                global.chk1 = "B";
            else if (comboBox1.Text == "전주공장")
                global.chk1 = "C";


        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBox2.Text == "포밍")
            { global.chk2 = "F"; }
            else if (comboBox2.Text == "태핑")
            { global.chk2 = "T"; }


        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBox3.Text == "1조")
            { global.chk3 = "1"; }
            else if (comboBox3.Text == "2조")
            { global.chk3 = "2"; }
            else if (comboBox3.Text == "3조")
            { global.chk3 = "3"; }
            else if (comboBox3.Text == "4조")
            { global.chk3 = "4"; }
            else if (comboBox3.Text == "5조")
            { global.chk3 = "5"; }
            else if (comboBox3.Text == "6조")
            { global.chk3 = "6"; }
            else if (comboBox3.Text == "7조")
            { global.chk3 = "7"; }
            else if (comboBox3.Text == "8조")
            { global.chk3 = "8"; }
            else if (comboBox3.Text == "9조")
            { global.chk3 = "9"; }
            else if (comboBox3.Text == "10조")
            { global.chk3 = "A0"; }

        }


    }
}