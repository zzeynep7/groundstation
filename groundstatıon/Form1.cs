using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using GMap.NET;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.WindowsForms.ToolTips;
using GMap.NET.WindowsForms;
using Microsoft.Office.Interop.Excel;
using offis = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using OpenTK.Graphics.OpenGL;
using OpenTK;


namespace groundstatıon
{
    public partial class Form1 : Form
    {
        public Form1()

        {
            InitializeComponent();

            CheckForIllegalCrossThreadCalls = false;  //multıthread a izin versin hata almasın diye
           
        }

        


        string data;
        string[] new_data;
        string data2;
        string[] new_data2;
        int l = 9;
        int rowc2 = 17;

        int d;
        int i = 9;

       // string x0;
       // string y0;
       // string z0;


        int x;
        int y;
        int z;
        
        

        byte[] paket = new byte[78];
        byte[] bytedElevation = new byte[4];
        byte[] bytedfloatedLat = new byte[4];
        byte[] bytedfloatedLongt = new byte[4];
        byte[] bytedHeight = new byte[4];
        byte[] bytedgpsheight = new byte[4];
        byte[] bytedpayloadlat = new byte[4];
        byte[] bytedpayloadlongt = new byte[4];
        byte[] bytedpayloadgpsheight = new byte[4];
        byte[] bytedgyrox = new byte[4];
        byte[] bytedgyroy = new byte[4];
        byte[] bytedgyroz = new byte[4];
        byte[] bytedfloatedgpsheight = new byte[4];
        byte[] intdurumr = new byte[2];


        float floatedElevation;
        float floatedLat;
        float floatedLongt;
        float floatedHeight;
        float floatedgpsheight;
        float floatedpayloadlat;
        float floatedpayloadlongt;
        float floatedpayloadgpsheight;
       // float floatedgyrox;
       // float floatedgyroy;
       // float floatedgyroz;
       // int intdurum;

        


        byte a = 0;
        int rowc = 25;
        int c ;


        private void Form1_Load(object sender, EventArgs e)
        {

            string[] ports = SerialPort.GetPortNames();
            string[] ports2 = SerialPort.GetPortNames();
            string[] ports3 = SerialPort.GetPortNames();
            guna2ComboBox1.Items.AddRange(ports);
            guna2ComboBox2.Items.AddRange(ports2);
            guna2ComboBox3.Items.AddRange(ports3);
            GL.ClearColor(Color.Yellow);
            //serialPort3.PortName = comboBox4.Text;
            //comboBox3.SelectedIndex = 1;
            glControl1.Update();


            //map.ReloadMap();
            map.DragButton = MouseButtons.Left;
            map.MapProvider = GMapProviders.GoogleSatelliteMap;
            map.MinZoom = 5;
            map.MaxZoom = 100;
            map.Zoom = 15;
            map.DragButton = MouseButtons.Left;


            map2.DragButton = MouseButtons.Left;
            map2.MinZoom = 5;
            map2.MaxZoom = 100;
            map2.Zoom = 15;
            map2.DragButton = MouseButtons.Left;



            dataGridView2.ColumnCount = 9;
            dataGridView2.RowCount = rowc2;
            dataGridView2.Columns[0].Name = "GÖREV YÜKÜ SICAKLIK";
            dataGridView2.Columns[1].Name = "GÖREV YÜKÜ NEM";
            dataGridView2.Columns[2].Name = "GÖREV YÜKÜ ENLEM";
            dataGridView2.Columns[3].Name = "GÖREV YÜKÜ BOYLAM";
            dataGridView2.Columns[4].Name = "GÖREV YÜKÜ YÜKSEKLİK";
            dataGridView2.Columns[5].Name = "GÖREV YÜKÜ GPS İRTİFA";
            dataGridView2.Columns[6].Name = "GÖREV YÜKÜ ENLEM";
            dataGridView2.Columns[7].Name = "GÖREV YÜKÜ BOYLAM";
            dataGridView2.Columns[8].Name = "PAKET SAYISI";



            //  column = sütun      row = sıra       cell = hücre


            try
            {
                dataGridView1.ColumnCount = 11;
                dataGridView1.RowCount = rowc;


                dataGridView1.Columns[0].Name = "TAKIM ID";
                dataGridView1.Columns[1].Name = "PAKET SAYISI";
                dataGridView1.Columns[2].Name = "ROKET YÜKSEKLİK";
                dataGridView1.Columns[3].Name = "ROKET ELEVATION";
                dataGridView1.Columns[4].Name = "ROKET ENLEM";
                dataGridView1.Columns[5].Name = "ROKET BOYLAM";
                dataGridView1.Columns[6].Name = "GPS İRTİFA";
                dataGridView1.Columns[7].Name = "OR X";
                dataGridView1.Columns[8].Name = "OR Y";
                dataGridView1.Columns[9].Name = "OR Z";
                dataGridView1.Columns[10].Name = "DURUM";
            }

            catch 
            {

               
            }
             


            //  dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            //  For equal width!
            // dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;  yavaşlatabilliyor.


            chart2.ChartAreas[0].AxisY.Minimum = 0;
            chart2.ChartAreas[0].AxisY.Maximum = 4000;
            chart2.ChartAreas[0].AxisY.Interval = 500;
            chart1.ChartAreas[0].AxisY.Minimum = -90;
            chart1.ChartAreas[0].AxisY.Maximum = 90;
            chart1.ChartAreas[0].AxisY.Interval = 15;


        }


     /*   private void timer1_Tick_1(object sender, EventArgs e)
        {

            try
            {
                data = serialPort1.ReadLine();

               // textBox1.Text = data.Length.ToString();
                // split metodunu if şartının içine alırsak dizin dizi sınırlarının dışında hatası kalkabilir mi...



                if (data.Length >= 30) //&& && data2.Length >= 7
                {


                    new_data = data.Split('*');

                    map.MapProvider = GMapProviders.GoogleSatelliteMap;
                    GMapOverlay markersOverlay = new GMapOverlay("markers");
                    markersOverlay.Markers.Clear();
                    map.Overlays.Clear();
                    double lat = double.Parse(new_data[2], System.Globalization.CultureInfo.InvariantCulture);
                    double longt = double.Parse(new_data[3], System.Globalization.CultureInfo.InvariantCulture);
                    GMarkerGoogle marker = new GMarkerGoogle(new PointLatLng(lat, longt), GMarkerGoogleType.red);
                    markersOverlay.Markers.Add(marker);
                    map.Overlays.Add(markersOverlay);
                    map.Position = new PointLatLng(lat, longt);






                        dataGridView1.RowCount = rowc;
                        dataGridView1.Rows[i].Cells[0].Value = 30;
                        dataGridView1.Rows[i].Cells[1].Value = c;
                        dataGridView1.Rows[i].Cells[2].Value = new_data[0];
                        dataGridView1.Rows[i].Cells[3].Value = new_data[1];
                        dataGridView1.Rows[i].Cells[4].Value = new_data[2];
                        dataGridView1.Rows[i].Cells[5].Value = new_data[3];
                        dataGridView1.Rows[i].Cells[6].Value = new_data[4];
                        dataGridView1.Rows[i].Cells[7].Value = new_data[5];
                        dataGridView1.Rows[i].Cells[8].Value = new_data[6];
                        dataGridView1.Rows[i].Cells[9].Value = new_data[7]; 
                        //dataGridView1.Rows[i].Cells[10].Value =new_data[8]; 
                       
                        rowc++;
                        c++;
                        

                        i++;

                        int f = i - 9;
                        dataGridView1.FirstDisplayedScrollingRowIndex = f;




                    chart1.Series[0].Points.AddXY(DateTime.Now.ToLongTimeString(), new_data[1]); // new_data[5]);   
                    chart2.Series[0].Points.AddXY(DateTime.Now.ToLongTimeString(), new_data[0]);//new_data[7]);
                   





                    if (c > 40)
                   {
                        dataGridView1.Rows[i - 39].Cells[0].Dispose();
                        dataGridView1.Rows[i - 39].Cells[1].Dispose();
                        dataGridView1.Rows[i - 39].Cells[2].Dispose();
                        dataGridView1.Rows[i - 39].Cells[3].Dispose();
                        dataGridView1.Rows[i - 39].Cells[4].Dispose();
                        dataGridView1.Rows[i - 39].Cells[5].Dispose();
                        dataGridView1.Rows[i - 39].Cells[6].Dispose();
                        dataGridView1.Rows[i - 39].Cells[7].Dispose();
                        dataGridView1.Rows[i - 39].Cells[8].Dispose();
                        dataGridView1.Rows[i - 39].Cells[9].Dispose();
                        //dataGridView1.Rows[i - 39].Cells[10].Dispose();

                    }


                    //  chart1.Update();
                    //  chart2.Update();
                      
                    dataGridView1.Refresh();
                      
                      // gl nesnesinin çizilmesi için çağrılan fonksiyon
                    

                    //serialPort1.DiscardInBuffer();
                }

               }


               catch 
               {
                 // MessageBox.Show(err.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
               }
           
           

           }*/



        private void copyAlltoClipboard2()
        {
            try
            {
                dataGridView2.SelectAll();
                DataObject dataObj = dataGridView2.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            catch 
            {

                
            }
            
        }



        private void copyAlltoClipboard()
        {
            try
            {
                dataGridView1.SelectAll();
                DataObject dataObj = dataGridView1.GetClipboardContent();
                if (dataObj != null)
                    Clipboard.SetDataObject(dataObj);
            }
            catch
            {

            }
           
        }





        private void timer2_Tick_1(object sender, EventArgs e)
        {
                      
            
            try
            {
                string datax = serialPort1.ReadLine();
                if (datax.Length >= 1)
                {
                    string[] data_pckg = datax.Split('*');
                    // label2.Text = new_data[4];
                    // x0 = label2.Text;
                    // label3.Text = new_data[5];
                    // y0 = label3.Text;
                    // label4.Text = new_data[6];
                    // z0 = label4.Text;
                    x = Convert.ToInt32(data_pckg[5]);
                    //  double x1 = Convert.ToDouble(x);
                    y = Convert.ToInt32(data_pckg[6]);
                    // double y1 = Convert.ToDouble(y);
                    z = Convert.ToInt32(data_pckg[7]);
                    // double z1 = Convert.ToDouble(z);



                    

                    float step = 1.0f;
                    float topla = step;
                    float radius = 2.0f; //kullanılan geometrik şekillerin yarıçapını temsil ediyor
                    float dikey1 = radius, dikey2 = -radius;
                    GL.Clear(ClearBufferMask.ColorBufferBit); //Gl nesnesinin sürekli silip yeniden çizdiriyoruz yoksa üst üste karmaşık bir şey ortaya çıkıyor
                    GL.Clear(ClearBufferMask.DepthBufferBit);

                    Matrix4 perspective = Matrix4.CreatePerspectiveFieldOfView(1.04f, 4 / 3, 1, 10000);
                    Matrix4 lookat = Matrix4.LookAt(25, 0, 0, 0, 0, 0, 0, 1, 0);
                    GL.MatrixMode(MatrixMode.Projection);
                    GL.LoadIdentity();
                    GL.LoadMatrix(ref perspective);
                    GL.MatrixMode(MatrixMode.Modelview);
                    GL.LoadIdentity();
                    GL.LoadMatrix(ref lookat);
                    GL.Viewport(0, 0, glControl1.Width, glControl1.Height);
                    GL.Enable(EnableCap.DepthTest);
                    GL.DepthFunc(DepthFunction.Less);

                   



                    GL.Rotate(y, 1.0, 0.0, 0.0);//ÖNEMLİ  x de 1 derece döndürme
                    GL.Rotate(z, 0.0, 1.0, 0.0);// y de 1 derece döndürme
                    GL.Rotate(x, 0.0, 0.0, 1.0);//z de 1 derece döndürme
                                                //Çizim için gerekli fonksiyonlar
                    silindir(step, topla, radius, 3, -13);  //5. değer silindirin boyunu uzatıyor.
                                                            //   silindir(0.01f, topla, 0.5f, 9, 9.7f); yukarıdaki küçük silindir için.
                                                            //   silindir(0.01f, topla, 0.1f, 5, dikey1 + 5);
                    koni(0.01f, 0.01f, radius, 0.05f, 3, 8); //4. değeri değiştirince koni şeklini alıyor yani ucu gittikçe daralıyor sonuçta koni şeklini elde ettim. 6. değeri değiştirince de koninin yüksekliği değişiyor.
                                                             //   koni(0.01f, 0.01f, radius, 2.0f, -5.0f, -10.0f); 5. değeri değiştirme...
                                                             //   Pervane(9.0f, 11.0f, 0.2f, 0.5f);

                    GL.Begin(BeginMode.Lines);

                    GL.Color3(Color.FromArgb(250, 0, 0));
                    GL.Vertex3(-30.0, 0.0, 0.0);
                    GL.Vertex3(30.0, 0.0, 0.0);

                    GL.Color3(Color.FromArgb(0, 0, 0));
                    GL.Vertex3(0.0, 30.0, 0.0);
                    GL.Vertex3(0.0, -30.0, 0.0);

                    GL.Color3(Color.FromArgb(0, 0, 250));
                    GL.Vertex3(0.0, 0.0, 30.0);
                    GL.Vertex3(0.0, 0.0, -30.0);

                    GL.End();



                    //GraphicsContext.CurrentContext.VSync = true;
                    glControl1.SwapBuffers();
                    glControl1.Update();
                    
                    // serialPort1.DiscardInBuffer(); // bunu yukarı çeksen ne olur glınvalıdate in altına

                }
            }
            catch 
            {

            }
           


        }

        private void silindir(float step, float topla, float radius, float dikey1, float dikey2)
        {



            float eski_step = 0.1f;
            GL.Begin(BeginMode.Quads);//Y EKSEN CIZIM DAİRENİN
            while (step <= 360)
            {
                if (step < 45)
                    GL.Color3(Color.Green); //FromArgb(255, 0, 0)
                else if (step < 90)
                    GL.Color3(Color.Green); //FromArgb(255, 255, 255)
                else if (step < 135)
                    GL.Color3(Color.Green); //FromArgb(255, 0, 0)
                else if (step < 180)
                    GL.Color3(Color.Green);  //renk kodlarını değiştirebilirsin.FromArgb(255, 255, 255)
                else if (step < 225)
                    GL.Color3(Color.Green); //FromArgb(255, 0, 0)
                else if (step < 270)
                    GL.Color3(Color.Green); //FromArgb(255, 255, 255)
                else if (step < 315)
                    GL.Color3(Color.Green); //FromArgb(255, 0, 0)
                else if (step < 360)
                    GL.Color3(Color.Green); //FromArgb(255, 255, 255)


                float ciz1_x = (float)(radius * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey1, ciz1_y);

                float ciz2_x = (float)(radius * Math.Cos((step + 2) * Math.PI / 180F));
                float ciz2_y = (float)(radius * Math.Sin((step + 2) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey1, ciz2_y);

                GL.Vertex3(ciz1_x, dikey2, ciz1_y);
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }


            GL.End();
            GL.Begin(BeginMode.Lines);
            step = eski_step;
            topla = step;
            while (step <= 180)// UST KAPAK
            {
                if (step < 45)
                    GL.Color3(Color.Goldenrod); //Color.FromArgb(255, 1, 1)
                else if (step < 90)
                    GL.Color3(Color.Goldenrod); //Color.FromArgb(250, 250, 200)
                else if (step < 135)
                    GL.Color3(Color.Goldenrod); //Color.FromArgb(255, 1, 1)
                else if (step < 180)
                    GL.Color3(Color.Goldenrod); //Color.FromArgb(250, 250, 200)
                else if (step < 225)
                    GL.Color3(Color.Goldenrod); //Color.FromArgb(255, 1, 1)
                else if (step < 270)
                    GL.Color3(Color.Goldenrod); //Color.FromArgb(250, 250, 200)
                else if (step < 315)
                    GL.Color3(Color.Goldenrod); //Color.FromArgb(255, 1, 1)
                else if (step < 360)
                    GL.Color3(Color.Goldenrod); //Color.FromArgb(250, 250, 200)


                float ciz1_x = (float)(radius * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey1, ciz1_y);

                float ciz2_x = (float)(radius * Math.Cos((step + 180) * Math.PI / 180F));
                float ciz2_y = (float)(radius * Math.Sin((step + 180) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey1, ciz2_y);

                GL.Vertex3(ciz1_x, dikey1, ciz1_y);
                GL.Vertex3(ciz2_x, dikey1, ciz2_y);
                step += topla;
            }



            step = eski_step;
            topla = step;
            while (step <= 180)//ALT KAPAK  Burada çalışan kodları kaldırırsan alt kapak görünmüyor
            {
                if (step < 45)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 90)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 135)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 180)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 225)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 270)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 315)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 360)
                    GL.Color3(Color.FromArgb(250, 250, 200));

                float ciz1_x = (float)(radius * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey2, ciz1_y);

                float ciz2_x = (float)(radius * Math.Cos((step + 180) * Math.PI / 180F));
                float ciz2_y = (float)(radius * Math.Sin((step + 180) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);

                GL.Vertex3(ciz1_x, dikey2, ciz1_y);
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }


            GL.End();
            glControl1.Update();


        }

        private void koni(float step, float topla, float radius1, float radius2, float dikey1, float dikey2)
        {


            float eski_step = 0.1f;
            GL.Begin(BeginMode.Lines);//Y EKSEN CIZIM DAİRENİN 
            while (step <= 360)
            {
                if (step < 45)
                    GL.Color3(Color.Black);
                else if (step < 90)
                    GL.Color3(Color.Black); //
                else if (step < 135)
                    GL.Color3(Color.Black); //1.0, 1.0, 1.0
                else if (step < 180)
                    GL.Color3(Color.Black);
                else if (step < 225)
                    GL.Color3(Color.Black);
                else if (step < 270)
                    GL.Color3(Color.Black);
                else if (step < 315)
                    GL.Color3(Color.Black);
                else if (step < 360)
                    GL.Color3(Color.Black);


                float ciz1_x = (float)(radius1 * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius1 * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey1, ciz1_y);

                float ciz2_x = (float)(radius2 * Math.Cos(step * Math.PI / 180F));
                float ciz2_y = (float)(radius2 * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }
            GL.End();

            GL.Begin(BeginMode.Lines); //BeginMode.Lines BeginMode.Quads
            step = eski_step;
            topla = step;
            while (step <= 180)// UST KAPAK 
            {
                if (step < 45)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 90)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 135)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 180)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 225)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 270)
                    GL.Color3(Color.FromArgb(250, 250, 200));
                else if (step < 315)
                    GL.Color3(Color.FromArgb(255, 1, 1));
                else if (step < 360)
                    GL.Color3(Color.FromArgb(250, 250, 200));


                float ciz1_x = (float)(radius2 * Math.Cos(step * Math.PI / 180F));
                float ciz1_y = (float)(radius2 * Math.Sin(step * Math.PI / 180F));
                GL.Vertex3(ciz1_x, dikey2, ciz1_y);

                float ciz2_x = (float)(radius2 * Math.Cos((step + 180) * Math.PI / 180F));
                float ciz2_y = (float)(radius2 * Math.Sin((step + 180) * Math.PI / 180F));
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);

                GL.Vertex3(ciz1_x, dikey2, ciz1_y);
                GL.Vertex3(ciz2_x, dikey2, ciz2_y);
                step += topla;
            }



            step = eski_step;
            topla = step;
            GL.End();
            glControl1.Update();



        }



      /*  private void glControl1_Load(object sender, EventArgs e)
        {
            try
            {
                GL.ClearColor(0.0f, 0.0f, 0.0f, 0.0f);
                GL.Enable(EnableCap.DepthTest);//sonradan yazdık
                glControl1.Update();
            }
            catch
            {
                //
            }
           
        }*/

      

       
        private void guna2Button1_Click(object sender, EventArgs e)
        {
           try
           {
                serialPort1.PortName = guna2ComboBox1.Text;
                serialPort1.BaudRate = 9600;
                serialPort1.DataBits = 8;
                serialPort1.Parity = Parity.None;
                serialPort1.StopBits = StopBits.One;
                serialPort1.Open();
                
                    serialPort2.PortName = guna2ComboBox2.Text;
                    serialPort2.BaudRate = 9600;
                    serialPort2.DataBits = 8;
                    serialPort2.Parity = Parity.None;
                    serialPort2.StopBits = StopBits.One;
                    serialPort2.Open();

                   
                    timer1.Start();
                // timer2.Start();
                timer4.Start();
                guna2Button1.Enabled = false;
                guna2Button2.Enabled = true;
           }
           catch 
           {
            // MessageBox.Show(err.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
           }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            try
            {
                timer1.Stop();
                if (serialPort1.IsOpen == true)
                {
                    serialPort1.Close();
                }
               // serialPort2.Close();
               /* if (backgroundWorker1.WorkerSupportsCancellation == true)
                {
                    backgroundWorker1.CancelAsync();
                }*/
                    // timer2.Stop();

                guna2Button1.Enabled = true;
                guna2Button2.Enabled = false;
            }
            catch
            {

            }
            
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen == true)
                {
                    serialPort1.Close();
                    timer1.Stop();
                }
                this.Close();
            }
            catch
            {

            }
           
            
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            try
            {
                copyAlltoClipboard();
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            catch (Exception hata)
            {
                MessageBox.Show(hata.StackTrace);
            }
        }

        


        private void guna2Button5_Click(object sender, EventArgs e)
        {
            


            try
            {
                
                serialPort3.PortName = guna2ComboBox3.Text;
                serialPort3.BaudRate = 19200;
                serialPort3.DataBits = 8;
                serialPort3.Parity = Parity.None;
                serialPort3.StopBits = StopBits.One;
                serialPort3.Open();
                backgroundWorker1.RunWorkerAsync();



            }
            catch
            {

            }
        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
           /* if ( serialPort3.IsOpen==true)
            {
                serialPort3.Close();
            }*/
           
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            try
            {
                float.TryParse(new_data[0], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedHeight);
                bytedHeight = BitConverter.GetBytes(floatedHeight);
                float.TryParse(new_data[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedElevation);
                bytedElevation = BitConverter.GetBytes(floatedElevation);
                float.TryParse(new_data[2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedLat);
                bytedfloatedLat = BitConverter.GetBytes(floatedLat);
                float.TryParse(new_data[3], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedLongt);
                bytedfloatedLongt = BitConverter.GetBytes(floatedLongt);
                float.TryParse(new_data[4], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedgpsheight);
                bytedfloatedgpsheight = BitConverter.GetBytes(floatedgpsheight);
                /*float.TryParse(x0, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedgyrox);
                  byte[] bytedgyrox = BitConverter.GetBytes(floatedgyrox);
                  float.TryParse(y0,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture, out floatedgyroy);
                  byte[] bytedgyroy = BitConverter.GetBytes(floatedgyroy);
                  float.TryParse(z0, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedgyroz);
                  byte[] bytedgyroz = BitConverter.GetBytes(floatedgyroz);*/
                float.TryParse(new_data2[7], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedpayloadlongt);
                  bytedpayloadlongt = BitConverter.GetBytes(floatedpayloadlongt);
                float.TryParse(new_data2[5], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedpayloadgpsheight);
                   bytedpayloadgpsheight = BitConverter.GetBytes(floatedpayloadgpsheight);
                float.TryParse(new_data2[6], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedpayloadlat);
                   bytedpayloadlat = BitConverter.GetBytes(floatedpayloadlat); 
                Int16.TryParse(new_data[8], out short intdurum);
                intdurumr = BitConverter.GetBytes(intdurum);
            }

            catch 
            {

            }

            try
            {
                paket[0] = 255;     //sabit
                paket[1] = 255;     //sabit
                paket[2] = 84;      //sabit
                paket[3] = 82;      //sabit
                paket[4] = 30;       //takim id
                paket[5] = a;       //paket no
                paket[6] = bytedHeight[0];        //roket irtifa
                paket[7] = bytedHeight[1];        //roket irtifa
                paket[8] = bytedHeight[2];        //roket irtifa
                paket[9] = bytedHeight[3];        //roket irtifa
                paket[10] = bytedfloatedgpsheight[0];//      //roket gps irtifa
                paket[11] = bytedfloatedgpsheight[1];//      //roket gps irtifa
                paket[12] = bytedfloatedgpsheight[2];//      //roket gps irtifa
                paket[13] = bytedfloatedgpsheight[3];//      //roket gps irtifa
                paket[14] = bytedfloatedLat[0];       //roket enlem
                paket[15] = bytedfloatedLat[1];       //roket enlem
                paket[16] = bytedfloatedLat[2];       //roket enlem
                paket[17] = bytedfloatedLat[3];       //roket enlem
                paket[18] = bytedfloatedLongt[0];       //roket boylam
                paket[19] = bytedfloatedLongt[1];       //roket boylam
                paket[20] = bytedfloatedLongt[2];       //roket boylam
                paket[21] = bytedfloatedLongt[3];       //roket boylam
                paket[22] = bytedpayloadgpsheight[0];       //gorev yuku gps  
                paket[23] = bytedpayloadgpsheight[1];       //gorev yuku gps  
                paket[24] = bytedpayloadgpsheight[2];       //gorev yuku gps  
                paket[25] = bytedpayloadgpsheight[3];       //gorev yuku gps  
                paket[26] = bytedpayloadlat[0];       //gorev yuku enlem 
                paket[27] = bytedpayloadlat[1];       //gorev yuku enlem 
                paket[28] = bytedpayloadlat[2];       //gorev yuku enlem 
                paket[29] = bytedpayloadlat[3];       //gorev yuku enlem 
                paket[30] = bytedpayloadlongt[0];       //gorev yuku boylam 
                paket[31] = bytedpayloadlongt[1];       //gorev yuku boylam 
                paket[32] = bytedpayloadlongt[2];       //gorev yuku boylam 
                paket[33] = bytedpayloadlongt[3];       //gorev yuku boylam 
                paket[34] = 0;       //kademe gps (zorlu)
                paket[35] = 0;       //kademe gps 
                paket[36] = 0;       //kademe gps
                paket[37] = 0;       //kademe gps
                paket[38] = 0;       //kademe enlem
                paket[39] = 0;       //kademe enlem
                paket[40] = 0;       //kademe enlem
                paket[41] = 0;       //kademe enlem
                paket[42] = 0;       //kademe boylam
                paket[43] = 0;       //kademe boylam
                paket[44] = 0;       //kademe boylam
                paket[45] = 0;       //kademe boylam
                paket[46] = 0;       //gyro x
                paket[47] = 0;       //gyro x
                paket[48] = 0;       //gyro x
                paket[49] = 0;       //gyro x
                paket[50] = 0;       //gyro y
                paket[51] = 0;       //gyro y
                paket[52] = 0;       //gyro y
                paket[53] = 0;       //gyro y
                paket[54] = 0;       //gyro z
                paket[55] = 0;       //gyro z
                paket[56] = 0;       //gyro z
                paket[57] = 0;       //gyro z
                paket[58] = 0;       //ivme x
                paket[59] = 0;       //ivme x
                paket[60] = 0;       //ivme x
                paket[61] = 0;       //ivme x
                paket[62] = 0;       //ivme y
                paket[63] = 0;       //ivme y
                paket[64] = 0;       //ivme y
                paket[65] = 0;       //ivme y
                paket[66] = 0;       //ivme z
                paket[67] = 0;       //ivme z
                paket[68] = 0;       //ivme z
                paket[69] = 0;       //ivme z
                paket[70] = bytedElevation[0];       //elevation
                paket[71] = bytedElevation[1];       //elevation
                paket[72] = bytedElevation[2];       //elevation
                paket[73] = bytedElevation[3];       //elevation
                paket[74] = intdurumr[0];       //durum

               
                int check_sum = 0;

                for (int j = 4; j < 75; j++)
                {

                    check_sum += paket[j];

                }


                byte moddedCheckSum = Convert.ToByte(check_sum % 256);

                paket[75] = moddedCheckSum;
                paket[76] = 13;
                paket[77] = 10;
                a++;
               // MessageBox.Show("çalışıyor");
                serialPort3.Write(paket, 0, 78);
                
            }

           catch 
            {
                //
            }

        }


        private void guna2Button7_Click(object sender, EventArgs e)
        {

            try
            {
                copyAlltoClipboard2();
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            catch (Exception hata)
            {
                MessageBox.Show(hata.StackTrace);
            }
        }

       
       

        private void guna2Button8_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();

           
         
           

        }

       

        private void timer4_Tick_1(object sender, EventArgs e)
        {

            try
            {


                data2 = serialPort2.ReadLine();
                if (data2.Length > 1)
                {
                    new_data2 = data2.Split('*');

                    /* map2.MapProvider = GMapProviders.GoogleSatelliteMap;
                     GMapOverlay markers1overlay = new GMapOverlay("markers1");
                     markers1overlay.Markers.Clear();
                     map2.Overlays.Clear();
                     double lat2 = double.Parse(new_data2[2], System.Globalization.CultureInfo.InvariantCulture);
                     double longt2 = double.Parse(new_data2[3], System.Globalization.CultureInfo.InvariantCulture);
                     GMarkerGoogle marker1 = new GMarkerGoogle(new PointLatLng(lat2, longt2), GMarkerGoogleType.red);
                     markers1overlay.Markers.Add(marker1);
                     map2.Overlays.Add(markers1overlay);
                     map2.Position = new PointLatLng(lat2, longt2);*/


                    dataGridView2.RowCount = rowc2;
                    dataGridView2.Rows[l].Cells[0].Value = new_data2[0];
                    dataGridView2.Rows[l].Cells[1].Value = new_data2[1];
                    dataGridView2.Rows[l].Cells[2].Value = new_data2[2];
                    dataGridView2.Rows[l].Cells[3].Value = new_data2[3];
                    dataGridView2.Rows[l].Cells[4].Value = new_data2[4];
                    // dataGridView2.Rows[l].Cells[5].Value = new_data2[5];
                    // dataGridView2.Rows[l].Cells[6].Value = new_data2[6];
                    // dataGridView2.Rows[l].Cells[7].Value = new_data2[7];
                    dataGridView2.Rows[l].Cells[5].Value = d;
                    d++;
                    rowc2++;
                    l++;

                    int h = l - 9;
                    dataGridView2.FirstDisplayedScrollingRowIndex = h;



                    if (d > 40)
                    {
                        dataGridView2.Rows[l - 39].Cells[0].Dispose();
                        dataGridView2.Rows[l - 39].Cells[1].Dispose();
                        dataGridView2.Rows[l - 39].Cells[2].Dispose();
                        dataGridView2.Rows[l - 39].Cells[3].Dispose();
                        dataGridView2.Rows[l - 39].Cells[4].Dispose();
                        // dataGridView2.Rows[l - 39].Cells[5].Dispose();
                        // dataGridView2.Rows[l - 39].Cells[6].Dispose();
                        // dataGridView2.Rows[l - 39].Cells[7].Dispose();
                        //  dataGridView2.Rows[l - 39].Cells[8].Dispose();
                        // dataGridView2.Rows[l - 39].Cells[9].Dispose();

                    }


                    dataGridView2.Refresh();
                }



            }
            catch

            {

            }


        }

        private void send_data(object sender, EventArgs e)
        {
            try
            {
                float.TryParse(new_data[0], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedHeight);
                bytedHeight = BitConverter.GetBytes(floatedHeight);
                float.TryParse(new_data[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedElevation);
                bytedElevation = BitConverter.GetBytes(floatedElevation);
                float.TryParse(new_data[2], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedLat);
                bytedfloatedLat = BitConverter.GetBytes(floatedLat);
                float.TryParse(new_data[3], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedLongt);
                bytedfloatedLongt = BitConverter.GetBytes(floatedLongt);
                float.TryParse(new_data[4], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedgpsheight);
                bytedfloatedgpsheight = BitConverter.GetBytes(floatedgpsheight);
                /*float.TryParse(x0, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedgyrox);
                  byte[] bytedgyrox = BitConverter.GetBytes(floatedgyrox);
                  float.TryParse(y0,System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture, out floatedgyroy);
                  byte[] bytedgyroy = BitConverter.GetBytes(floatedgyroy);
                  float.TryParse(z0, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedgyroz);
                  byte[] bytedgyroz = BitConverter.GetBytes(floatedgyroz);*/
                float.TryParse(new_data2[7], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedpayloadlongt);
                bytedpayloadlongt = BitConverter.GetBytes(floatedpayloadlongt);
                float.TryParse(new_data2[5], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedpayloadgpsheight);
                bytedpayloadgpsheight = BitConverter.GetBytes(floatedpayloadgpsheight);
                float.TryParse(new_data2[6], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out floatedpayloadlat);
                bytedpayloadlat = BitConverter.GetBytes(floatedpayloadlat);
                Int16.TryParse(new_data[8], out short intdurum);
                intdurumr = BitConverter.GetBytes(intdurum);
            }

            catch
            {

            }

            try
            {
                paket[0] = 255;     //sabit
                paket[1] = 255;     //sabit
                paket[2] = 84;      //sabit
                paket[3] = 82;      //sabit
                paket[4] = 30;       //takim id
                paket[5] = a;       //paket no
                paket[6] = bytedHeight[0];        //roket irtifa
                paket[7] = bytedHeight[1];        //roket irtifa
                paket[8] = bytedHeight[2];        //roket irtifa
                paket[9] = bytedHeight[3];        //roket irtifa
                paket[10] = bytedfloatedgpsheight[0];//      //roket gps irtifa
                paket[11] = bytedfloatedgpsheight[1];//      //roket gps irtifa
                paket[12] = bytedfloatedgpsheight[2];//      //roket gps irtifa
                paket[13] = bytedfloatedgpsheight[3];//      //roket gps irtifa
                paket[14] = bytedfloatedLat[0];       //roket enlem
                paket[15] = bytedfloatedLat[1];       //roket enlem
                paket[16] = bytedfloatedLat[2];       //roket enlem
                paket[17] = bytedfloatedLat[3];       //roket enlem
                paket[18] = bytedfloatedLongt[0];       //roket boylam
                paket[19] = bytedfloatedLongt[1];       //roket boylam
                paket[20] = bytedfloatedLongt[2];       //roket boylam
                paket[21] = bytedfloatedLongt[3];       //roket boylam
                paket[22] = bytedpayloadgpsheight[0];       //gorev yuku gps  
                paket[23] = bytedpayloadgpsheight[1];       //gorev yuku gps  
                paket[24] = bytedpayloadgpsheight[2];       //gorev yuku gps  
                paket[25] = bytedpayloadgpsheight[3];       //gorev yuku gps  
                paket[26] = bytedpayloadlat[0];       //gorev yuku enlem 
                paket[27] = bytedpayloadlat[1];       //gorev yuku enlem 
                paket[28] = bytedpayloadlat[2];       //gorev yuku enlem 
                paket[29] = bytedpayloadlat[3];       //gorev yuku enlem 
                paket[30] = bytedpayloadlongt[0];       //gorev yuku boylam 
                paket[31] = bytedpayloadlongt[1];       //gorev yuku boylam 
                paket[32] = bytedpayloadlongt[2];       //gorev yuku boylam 
                paket[33] = bytedpayloadlongt[3];       //gorev yuku boylam 
                paket[34] = 0;       //kademe gps (zorlu)
                paket[35] = 0;       //kademe gps 
                paket[36] = 0;       //kademe gps
                paket[37] = 0;       //kademe gps
                paket[38] = 0;       //kademe enlem
                paket[39] = 0;       //kademe enlem
                paket[40] = 0;       //kademe enlem
                paket[41] = 0;       //kademe enlem
                paket[42] = 0;       //kademe boylam
                paket[43] = 0;       //kademe boylam
                paket[44] = 0;       //kademe boylam
                paket[45] = 0;       //kademe boylam
                paket[46] = 0;       //gyro x
                paket[47] = 0;       //gyro x
                paket[48] = 0;       //gyro x
                paket[49] = 0;       //gyro x
                paket[50] = 0;       //gyro y
                paket[51] = 0;       //gyro y
                paket[52] = 0;       //gyro y
                paket[53] = 0;       //gyro y
                paket[54] = 0;       //gyro z
                paket[55] = 0;       //gyro z
                paket[56] = 0;       //gyro z
                paket[57] = 0;       //gyro z
                paket[58] = 0;       //ivme x
                paket[59] = 0;       //ivme x
                paket[60] = 0;       //ivme x
                paket[61] = 0;       //ivme x
                paket[62] = 0;       //ivme y
                paket[63] = 0;       //ivme y
                paket[64] = 0;       //ivme y
                paket[65] = 0;       //ivme y
                paket[66] = 0;       //ivme z
                paket[67] = 0;       //ivme z
                paket[68] = 0;       //ivme z
                paket[69] = 0;       //ivme z
                paket[70] = bytedElevation[0];       //elevation
                paket[71] = bytedElevation[1];       //elevation
                paket[72] = bytedElevation[2];       //elevation
                paket[73] = bytedElevation[3];       //elevation
                paket[74] = intdurumr[0];       //durum


                int check_sum = 0;

                for (int j = 4; j < 75; j++)
                {

                    check_sum += paket[j];

                }


                byte moddedCheckSum = Convert.ToByte(check_sum % 256);

                paket[75] = moddedCheckSum;
                paket[76] = 13;
                paket[77] = 10;
                a++;
                // MessageBox.Show("çalışıyor");
                serialPort3.Write(paket, 0, 78);

            }

            catch
            {
                //
            }
        }






        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
           
        }

        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
           data = serialPort1.ReadLine();
            this.BeginInvoke(new EventHandler(displaydataevent));

        }

        private void displaydataevent(object sender, EventArgs e)
        {
            try
            {
               // data = serialPort1.ReadLine();

                // textBox1.Text = data.Length.ToString();
                // split metodunu if şartının içine alırsak dizin dizi sınırlarının dışında hatası kalkabilir mi...



                if (data.Length >= 20) //&& && data2.Length >= 7
                {


                    new_data = data.Split('*');

                    map.MapProvider = GMapProviders.GoogleSatelliteMap;
                    GMapOverlay markersOverlay = new GMapOverlay("markers");
                    markersOverlay.Markers.Clear();
                    map.Overlays.Clear();
                    double lat = double.Parse(new_data[2], System.Globalization.CultureInfo.InvariantCulture);
                    double longt = double.Parse(new_data[3], System.Globalization.CultureInfo.InvariantCulture);
                    GMarkerGoogle marker = new GMarkerGoogle(new PointLatLng(lat, longt), GMarkerGoogleType.red);
                    markersOverlay.Markers.Add(marker);
                    map.Overlays.Add(markersOverlay);
                    map.Position = new PointLatLng(lat, longt);






                    dataGridView1.RowCount = rowc;
                    dataGridView1.Rows[i].Cells[0].Value = 30;
                    dataGridView1.Rows[i].Cells[1].Value = c;
                    dataGridView1.Rows[i].Cells[2].Value = new_data[0];
                    dataGridView1.Rows[i].Cells[3].Value = new_data[1];
                    dataGridView1.Rows[i].Cells[4].Value = new_data[2];
                    dataGridView1.Rows[i].Cells[5].Value = new_data[3];
                    dataGridView1.Rows[i].Cells[6].Value = new_data[4];
                    dataGridView1.Rows[i].Cells[7].Value = new_data[5];
                    dataGridView1.Rows[i].Cells[8].Value = new_data[6];
                    dataGridView1.Rows[i].Cells[9].Value = new_data[7];
                    //dataGridView1.Rows[i].Cells[10].Value =new_data[8]; 

                    rowc++;
                    c++;


                    i++;

                    int f = i - 9;
                    dataGridView1.FirstDisplayedScrollingRowIndex = f;




                    chart1.Series[0].Points.AddXY(DateTime.Now.ToLongTimeString(), new_data[1]); // new_data[5]);   
                    chart2.Series[0].Points.AddXY(DateTime.Now.ToLongTimeString(), new_data[0]);//new_data[7]);






                    if (c > 40)
                    {
                        dataGridView1.Rows[i - 39].Cells[0].Dispose();
                        dataGridView1.Rows[i - 39].Cells[1].Dispose();
                        dataGridView1.Rows[i - 39].Cells[2].Dispose();
                        dataGridView1.Rows[i - 39].Cells[3].Dispose();
                        dataGridView1.Rows[i - 39].Cells[4].Dispose();
                        dataGridView1.Rows[i - 39].Cells[5].Dispose();
                        dataGridView1.Rows[i - 39].Cells[6].Dispose();
                        dataGridView1.Rows[i - 39].Cells[7].Dispose();
                        dataGridView1.Rows[i - 39].Cells[8].Dispose();
                        dataGridView1.Rows[i - 39].Cells[9].Dispose();
                        //dataGridView1.Rows[i - 39].Cells[10].Dispose();

                    }


                    //  chart1.Update();
                    //  chart2.Update();

                    dataGridView1.Refresh();

                    // gl nesnesinin çizilmesi için çağrılan fonksiyon


                    //serialPort1.DiscardInBuffer();
                }

            }


            catch
            {
                // MessageBox.Show(err.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void backgroundWorker1_DoWork_1(object sender, DoWorkEventArgs e)
        {
            do
            {
                send_data(sender, e);
                System.Threading.Thread.Sleep(100);
            } while (serialPort3.IsOpen == true);
        }
    }



}
    

