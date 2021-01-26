using System;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Dynamic;
using System.Data.SqlClient;
using Newtonsoft.Json;

namespace ExceltoJSonSpecs
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            var pathToExcel = @"~/../../../../DotNet452.Tools_marka_model/ExceltoJson.xlsx";
            var sheetName = "PhoneSpecs";
            

            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(@"
                Provider=Microsoft.ACE.OLEDB.12.0;
                Data Source={0};
                Extended Properties=""Excel 12.0 Xml;HDR=YES""
            ", pathToExcel);
            var J = 0;

            
            

            string connectionStringDB = "data source=.\\SQLEXPRESS;initial catalog=DotNet452DB;integrated security=True;MultipleActiveResultSets=True;App=EntityFramework";
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName
                    );

                using (var rdr = cmd.ExecuteReader())
                {
                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        (from DbDataRecord row in rdr
                         select row).Select(x =>
                         {


                             dynamic itemdynamic = new ExpandoObject();
     
                             if (J != 0)
                             {
                                 for (int i = 0; i < x.FieldCount; i++)
                                 {
                                     // 'Başlık': 'Apple iPhone 6',
                                     // 'Genel': { 'Üretici': 'Apple'},
                                     //  'İşlemci':{ 'Model': 'Apple A8','Hız': '1.4 GHz'},
                                     //itemdynamic.Başlık = x[0];
                                     itemdynamic.Genel = new { Model = x[1].ToString() };
                                     itemdynamic.İşlemci = new { Model = x[2].ToString(), Hız = x[3].ToString() };
                                     itemdynamic.İşletim_Sistemi = x[4].ToString();
                                     itemdynamic.Hafıza_Bellek = new { RAM = x[5].ToString(), Hafıza = x[6].ToString(), MicroSD = x[7].ToString() };
                                     itemdynamic.Ekran = new { Açıklama = x[8].ToString(), Çözünürlük = x[9].ToString(), Ekran_Boyutu = x[10].ToString(), Özellikler = x[11].ToString() };
                                     itemdynamic.Fiziksel = new { Boyutlar = x[12].ToString(), Ağırlıkk = x[13].ToString(), Renkler = x[14].ToString(), Suya_Dayanıklılık = x[15].ToString() };
                                     itemdynamic.Kamera = new { Açıklama = x[16].ToString(), Megapixel = x[17].ToString(), Özellikler = x[18].ToString(), Ön_Kamera_Megapixel = x[19].ToString() };
                                     itemdynamic.Bağlantı = new { Wifi = x[20].ToString(), Bluetooth = x[21].ToString(), GPS = x[22].ToString(), Bağlantı_Türü = x[23].ToString(), Micro_USB = x[24].ToString(), Sim_Boyutu = x[25].ToString(), Sim_Sayısı = x[26].ToString(), ÜçG_Görüntülü_Konuşma = x[27].ToString() };
                                     itemdynamic.Batarya = new { Lityum_Ion = x[28].ToString(), Değiştirilebilir = x[29].ToString() };
                                     itemdynamic.Diğer = new { SAR = x[30].ToString(), Parmak_İzi_Okuyucu = x[31].ToString(), Radyo = x[32].ToString() };
                                 }
                             }

                             
                             var jsonitem = JsonConvert.SerializeObject(itemdynamic);

                             string jsonspec= Convert.ToString(jsonitem);
                             
                             if (J != 0)
                             {
                                 using (SqlConnection sqlConnection = new SqlConnection(connectionStringDB))
                                 {                         
                                     using (SqlCommand sqlCommand = new SqlCommand("update Models set Specs=@Specs where Name=@Name",sqlConnection))
                                     {
                                         sqlCommand.Parameters.AddWithValue("@Name", x[0].ToString());
                                         sqlCommand.Parameters.AddWithValue("@Specs", jsonspec.Replace("_"," ").Replace("Üç","3"));
                                         sqlConnection.Open();
                                         sqlCommand.ExecuteNonQuery();
                                         sqlConnection.Close();
                                     }
                                 }
                             }
                             J++;
                             return jsonitem;
                         });
                    //Generates JSON from the LINQ query
                    var json = JsonConvert.SerializeObject(query);
                    //return json;
                }
            }

        }

        private void InsertColor_Click(object sender, EventArgs e)
        {
            var pathToExcel = @"~/../../../../DotNet452.Tools_marka_model/ExceltoJson.xlsx";
            var sheetName = "PhoneSpecs";
            //Use this connection string if you have Office 2007+ drivers installed and 
            //your data is saved in a .xlsx file
            var connectionString = String.Format(@"
                Provider=Microsoft.ACE.OLEDB.12.0;
                Data Source={0};
                Extended Properties=""Excel 12.0 Xml;HDR=YES""
            ", pathToExcel);
            var J = 0;
            string[] colors;
            string[] storages;
            string modelId = "";
            string connectionStringDB = "data source=.\\SQLEXPRESS;initial catalog=DotNet452DB;integrated security=True;MultipleActiveResultSets=True;App=EntityFramework";
            using (var conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                var cmd = conn.CreateCommand();
                cmd.CommandText = String.Format(
                    @"SELECT * FROM [{0}$]",
                    sheetName
                    );
                using (var rdr = cmd.ExecuteReader())
                {
                    //LINQ query - when executed will create anonymous objects for each row
                    var query =
                        (from DbDataRecord row in rdr
                         select row).Select(x =>
                         {
                             if (J != 0)
                             {
                                 colors = x[14].ToString().Split(',');
                                 storages = x[6].ToString().Split(',');

                                 using (SqlConnection sqlConnection = new SqlConnection(connectionStringDB))
                                 {
                                     using (SqlCommand sqlCommand = new SqlCommand("select Id from Models where Name=@Name", sqlConnection))
                                     {
                                         sqlCommand.Parameters.AddWithValue("@Name", x[0].ToString());
                                         sqlConnection.Open();
                                         SqlDataReader read = sqlCommand.ExecuteReader();
                                         while (read.Read())
                                         {
                                             modelId = (read["Id"].ToString());
                                         }
                                         read.Close();
                                         sqlConnection.Close();
                                     }
                                     foreach (string items in colors)
                                     {
                                         using (SqlCommand sqlCommand = new SqlCommand("insert into ModelColorOptions(Id, ModelId, Name, Description) VALUES(@Id, @ModelId, @Name, @Description)", sqlConnection))
                                         {
                                             sqlCommand.Parameters.AddWithValue("@Id", Guid.NewGuid());
                                             sqlCommand.Parameters.AddWithValue("@ModelId", new Guid(modelId));
                                             sqlCommand.Parameters.AddWithValue("@Name", items.Replace("Black", "Siyah").Replace("Red", "Kırmızı").Replace("Blue", "Mavi").Replace("Goldenrod", "Altın").Replace("White", "Beyaz").Replace("Silver", "Gümüş").Replace("Grey", "Gri").Replace("Gray", "Gri").Replace("Green", "Yeşil").Replace("Pink", "Pembe").Replace("Rose Gold", "Altın").Replace("Jet Black", "Siyah").Replace("Gold", "Altın").Replace("Cyan", "Cam Göbeği").Replace("Purple", "Mor").Trim());
                                             sqlCommand.Parameters.AddWithValue("@Description", items.Replace("Rose Gold", "Goldenrod").Replace("Jet Black", "Black").Replace("Gray","Grey").Trim());
                                             sqlConnection.Open();
                                             sqlCommand.ExecuteNonQuery();
                                             sqlConnection.Close();
                                         }
                                     }

                                     foreach (string items in storages)
                                     {
                                         using (SqlCommand sqlCommand = new SqlCommand("insert into ModelStorageOptions(Id, ModelId, Name, Description) VALUES(@Id, @ModelId, @Name, @Description)", sqlConnection))
                                         {
                                             sqlCommand.Parameters.AddWithValue("@Id", Guid.NewGuid());
                                             sqlCommand.Parameters.AddWithValue("@ModelId", new Guid(modelId));
                                             sqlCommand.Parameters.AddWithValue("@Name", items.Trim());
                                             sqlCommand.Parameters.AddWithValue("@Description", items.Trim().Split(' ')[0].ToString()+ " GigaByte");
                                             sqlConnection.Open();
                                             sqlCommand.ExecuteNonQuery();
                                             sqlConnection.Close();
                                         }
                                     }
                                 }
                             }
                             J++;
                             return "";
                         });
                    //Generates JSON from the LINQ query
                    var json = JsonConvert.SerializeObject(query);
                    //return json;
                }
            }
        }

        private void InsertModel_Click(object sender, EventArgs e)
        {

        }
    }

}
