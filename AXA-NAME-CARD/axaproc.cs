using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Security;
using System.Security.Cryptography;

namespace AXA_NAME_CARD
{
    class axaproc
    {
        const int Keysize = 256;
        const int DerivationIterations = 1221;
        OleDbConnection conn = new OleDbConnection();
        SqlConnection con;
        SqlDataAdapter adapter;
        static string constr = Properties.Settings.Default.ConSql.Trim();
        static string connectdb = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Properties.Settings.Default.ConMDB.Trim() + ";User Id=admin;Password=;";
   
        DataTable dt;
        public void connectmdb()
        {
            conn = new OleDbConnection();
            conn.ConnectionString = connectdb;
            try
            {
                conn.Open();
                // Insert code to process data.
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to connect to data source:" + ex.ToString());
                if (conn.State != ConnectionState.Closed)
                {
                    conn.Close();
                }
            }
        }
        public void connection()
        {
            con = new SqlConnection(constr);

            if (con.State != ConnectionState.Closed)
            {
                con.Close();
            }
        }
        public SqlDataReader GetData(string query)
        {
            SqlDataReader reader = null;
            SqlConnection connection = new SqlConnection(constr);
            try
            {
                SqlCommand command = new SqlCommand(query, connection);
                command.Connection = connection;
                connection.Open();
                reader = command.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch
            {

                connection.Close();
            }
            return reader;
        }
        public void executeQuery(string query)
        {
            con = new SqlConnection(constr);
            if (con.State != ConnectionState.Closed)
            {
                con.Close();
            }

            con.Open();

            SqlCommand ObjCmd = new SqlCommand(query, con);
            ObjCmd.CommandTimeout = 10000;
            ObjCmd.ExecuteNonQuery();
            con.Close();

        }

        public DataTable bukatabel(string query)
        {
            conn = new OleDbConnection();
            conn.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Properties.Settings.Default.ConMDB.Trim() + ";User Id=admin;Password=;";
            try
            { conn.Open();}
            catch (Exception ex)
            {MessageBox.Show(ex.Message);}
            OleDbCommand cmd = conn.CreateCommand();
            cmd.CommandText = query;
            OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection); // close conn after complete
            dt.Load(reader);
            return dt;
        }
        public DataTable openTable(string query)
        {
            con = new SqlConnection(constr); 
            if (con.State != ConnectionState.Closed)
            {con.Close();}
            con.Open();
            dt = new DataTable();
            adapter = new SqlDataAdapter(query, constr);
            adapter.SelectCommand.CommandTimeout = 6000;
            adapter.Fill(dt);
            con.Close();
            return dt;
        }   
        public void embed(StreamWriter fhandle, string finput)
        {
            string[] bacafile = File.ReadAllLines(finput);
            foreach (string perbaris in bacafile)
            {
                fhandle.Write(perbaris + "\n");  // amnbil isi txt string[] isinya = perbaris.Split();
            }
        }
        public void addtmpl(StreamWriter fhandle, string dirlist, string flogo, string nmback)
        {
            if (flogo == "Y")
            {
                foreach (string isifile in Directory.GetFiles(dirlist, "*.eps"))
                {
                    string vdata = "logo_" + Path.GetFileName(isifile).Replace(".eps", "").Replace(".EPS", "");
                    string tmpl = "/temp_" + Path.GetFileName(isifile).Replace(".eps", "").Replace(".EPS", "");
                    //Kasih Logo
                    fhandle.WriteLine("/" + vdata);
                    fhandle.WriteLine("currentfile");
                    fhandle.WriteLine("<< /Filter /SubFileDecode");
                    fhandle.WriteLine("/DecodeParms << /EODCount 0 /EODString (*EOD*) >>");
                    fhandle.WriteLine(">> /ReusableStreamDecode filter");
                    fhandle.WriteLine("%%");
                    fhandle.WriteLine("%% ************************ Copy paste File Form Eps disini ******************************");
                    //embed logo
                    embed(fhandle, isifile);
                    fhandle.WriteLine("%%");
                    fhandle.WriteLine("%%*EOD*");
                    fhandle.WriteLine("def");
                    fhandle.WriteLine(tmpl);
                    fhandle.WriteLine("<< /FormType 1");
                    fhandle.WriteLine("   /BBox [0 0 162 162]");
                    fhandle.WriteLine("   /Matrix [ 1 0 0 1 0 0]");
                    fhandle.WriteLine("   /PaintProc");
                    fhandle.WriteLine("   { pop");
                    fhandle.WriteLine("       /ostate save def");
                    fhandle.WriteLine("         /showpage {} def");
                    fhandle.WriteLine("         /setpagedevice /pop load def");
                    fhandle.WriteLine("         " + vdata + " 0 setfileposition " + vdata + " cvx exec");
                    fhandle.WriteLine("       ostate restore");
                    fhandle.WriteLine("   } bind");
                    fhandle.WriteLine(">> def");
                }
            }
            else
            {
                string vdata = "logo_" + nmback;
                string tmpl = "/temp_" + nmback;
                string isifile = dirlist + nmback+".eps";
                //Kasih Logo
                fhandle.WriteLine("/" + vdata);
                fhandle.WriteLine("currentfile");
                fhandle.WriteLine("<< /Filter /SubFileDecode");
                fhandle.WriteLine("/DecodeParms << /EODCount 0 /EODString (*EOD*) >>");
                fhandle.WriteLine(">> /ReusableStreamDecode filter");
                fhandle.WriteLine("%%");
                fhandle.WriteLine("%% ************************ Copy paste File Form Eps disini ******************************");
                //embed logo
                embed(fhandle, isifile);
                fhandle.WriteLine("%%");
                fhandle.WriteLine("%%*EOD*");
                fhandle.WriteLine("def");
                fhandle.WriteLine(tmpl);
                fhandle.WriteLine("<< /FormType 1");
                fhandle.WriteLine("   /BBox [0 0 906.9 1276.05]");
                fhandle.WriteLine("   /Matrix [ 1 0 0 1 0 0]");
                fhandle.WriteLine("   /PaintProc");
                fhandle.WriteLine("   { pop");
                fhandle.WriteLine("       /ostate save def");
                fhandle.WriteLine("         /showpage {} def");
                fhandle.WriteLine("         /setpagedevice /pop load def");
                fhandle.WriteLine("         " + vdata + " 0 setfileposition " + vdata + " cvx exec");
                fhandle.WriteLine("       ostate restore");
                fhandle.WriteLine("   } bind");
                fhandle.WriteLine(">> def");
            }
        }
        public void buatdir(string pathdir)
        {
            if (!Directory.Exists(pathdir))
            {
                Directory.CreateDirectory(pathdir);
            }
        }
        public string Encrypt(string plainText, string passPhrase)
        {
            // Salt and IV is randomly generated each time, but is preprended to encrypted cipher text
            // so that the same Salt and IV values can be used when decrypting.  
            var saltStringBytes = Generate256BitsOfRandomEntropy();
            var ivStringBytes = Generate256BitsOfRandomEntropy();
            var plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            using (var password = new Rfc2898DeriveBytes(passPhrase, saltStringBytes, DerivationIterations))
            {
                var keyBytes = password.GetBytes(Keysize / 8);
                using (var symmetricKey = new RijndaelManaged())
                {
                    symmetricKey.BlockSize = 256;
                    symmetricKey.Mode = CipherMode.CBC;
                    symmetricKey.Padding = PaddingMode.PKCS7;
                    using (var encryptor = symmetricKey.CreateEncryptor(keyBytes, ivStringBytes))
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                            {
                                cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                                cryptoStream.FlushFinalBlock();
                                // Create the final bytes as a concatenation of the random salt bytes, the random iv bytes and the cipher bytes.
                                var cipherTextBytes = saltStringBytes;
                                cipherTextBytes = cipherTextBytes.Concat(ivStringBytes).ToArray();
                                cipherTextBytes = cipherTextBytes.Concat(memoryStream.ToArray()).ToArray();
                                memoryStream.Close();
                                cryptoStream.Close();
                                return Convert.ToBase64String(cipherTextBytes);
                            }
                        }
                    }
                }
            }
        }

        public string Decrypt(string cipherText, string passPhrase)
        {
            // Get the complete stream of bytes that represent:
            // [32 bytes of Salt] + [32 bytes of IV] + [n bytes of CipherText]
            var cipherTextBytesWithSaltAndIv = Convert.FromBase64String(cipherText);
            // Get the saltbytes by extracting the first 32 bytes from the supplied cipherText bytes.
            var saltStringBytes = cipherTextBytesWithSaltAndIv.Take(Keysize / 8).ToArray();
            // Get the IV bytes by extracting the next 32 bytes from the supplied cipherText bytes.
            var ivStringBytes = cipherTextBytesWithSaltAndIv.Skip(Keysize / 8).Take(Keysize / 8).ToArray();
            // Get the actual cipher text bytes by removing the first 64 bytes from the cipherText string.
            var cipherTextBytes = cipherTextBytesWithSaltAndIv.Skip((Keysize / 8) * 2).Take(cipherTextBytesWithSaltAndIv.Length - ((Keysize / 8) * 2)).ToArray();

            using (var password = new Rfc2898DeriveBytes(passPhrase, saltStringBytes, DerivationIterations))
            {
                var keyBytes = password.GetBytes(Keysize / 8);
                using (var symmetricKey = new RijndaelManaged())
                {
                    symmetricKey.BlockSize = Keysize;
                    symmetricKey.Mode = CipherMode.CBC;
                    symmetricKey.Padding = PaddingMode.PKCS7;
                    using (var decryptor = symmetricKey.CreateDecryptor(keyBytes, ivStringBytes))
                    {
                        using (var memoryStream = new MemoryStream(cipherTextBytes))
                        {
                            using (var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read))
                            {
                                var plainTextBytes = new byte[cipherTextBytes.Length];
                                var decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
                                memoryStream.Close();
                                cryptoStream.Close();
                                return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount);
                            }
                        }
                    }
                }
            }
        }

        private static byte[] Generate256BitsOfRandomEntropy()
        {
            var randomBytes = new byte[32]; // 32 Bytes will give us 256 bits.
            using (var rngCsp = new RNGCryptoServiceProvider())
            {
                // Fill the array with cryptographically secure random bytes.
                rngCsp.GetBytes(randomBytes);
            }
            return randomBytes;
        }
        public string[] perkata(string ftxt, int brs, int maxline)
        {
            string[] addr = new string[brs];
            for (int jh = 0; jh < brs; jh++)
            {
                addr[jh] = "";
            }
            ftxt = ftxt.Trim();
            int pjg = ftxt.Length;
            int pjx = 0;
            string upword = ftxt.Trim();
            string mh = "";
            int loop = 0;
            while (loop < brs) 
            {
                if (pjg > maxline)
                {
                    for (int k = pjg - 1; k > 0; k--)
                    {
                        string mt = upword.Trim().Substring(k, 1);
                        if (mt == " " || mt == ":" || mt == ",")
                        {
                            if (pjg - mh.Length > maxline)
                            {
                                mh = mh + mt;
                            }
                            else
                            { break; }
                        }
                        else
                        { mh = mh + mt; }
                    }
                    pjx = mh.Length;
                    addr[loop] = upword.Substring(0, pjg - pjx);
                    int prg = upword.Substring(0, pjg - pjx).Length;
                    upword = upword.Substring(prg);

                    pjg = upword.Trim().Length;
                    mh = "";
                    loop++;
                }
                else
                {
                    addr[loop] = upword.Trim();
                    break;
                }
            }
            return (addr);
        }
        public DataTable excelapprove(string fullname)
        {

            DataTable approvedxls = new DataTable();
            OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + fullname + ";Extended Properties=Excel 8.0");
            //OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fullname + ";Extended Properties=Excel 12.0");
            connection.Open();
            DataTable Sheets = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            connection.Close();
            foreach (DataRow dr in Sheets.Rows)
            {
                string sht = dr[2].ToString().Replace("'", "");
                DataTable dt2 = new DataTable();
                DataSet ds2 = new DataSet();
                var connectionString = "";
                connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; data source={0}; Extended Properties=Excel 12.0;", fullname);
                connectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fullname);
                try
                {
                    var adapter2 = new OleDbDataAdapter("SELECT * FROM [" + sht + "]", connectionString); //colon name
                    adapter2.Fill(ds2, "NameCard");
                    dt2 = ds2.Tables["NameCard"];
                    approvedxls = dt2;
                    //DataView dv = new DataView();
                    //DataTable sortedDT = new DataTable();
                    //dv = dt2.DefaultView;
                    //dv.Sort = "NIK"; ///urutkan data berdasarkan collcode,prodcode,Policy No
                    //sortedDT = dv.ToTable();
                    //ListXls = sortedDT;
                    break;
                }
                catch (Exception)
                {
                    ds2 = null;
                    dt2 = null;
                }
            }
            return approvedxls;
        }
        void cetakbackground(StreamWriter fs, int start, int jmldok, DataTable ListXls, string fpage)
        {
            int flogo = 0;
            int col = 0, brs = 1;
            for (int jj = start; jj < jmldok; jj++)
            {
                col++;
                //lembar Background
                int clogo = Convert.ToInt32(ListXls.Rows[jj]["JmlLogo"].ToString().Trim());
                string l1 = "temp_" + ListXls.Rows[jj]["LOGO1"].ToString().Trim().ToLower();
                string l2 = "temp_" + ListXls.Rows[jj]["LOGO2"].ToString().Trim().ToLower();
                string l3 = "temp_" + ListXls.Rows[jj]["LOGO3"].ToString().Trim().ToLower();
                string l4 = "temp_" + ListXls.Rows[jj]["LOGO4"].ToString().Trim().ToLower();
                if (l1 == "temp_")
                { l1 = "%%%temp_"; }
                if (l2 == "temp_")
                { l2 = "%%%temp_"; }
                if (l3 == "temp_")
                { l3 = "%%%temp_"; }
                if (l4 == "temp_")
                { l4 = "%%%temp_"; }

                int jp = (jj + 1) % 21;
                if (jp == 1 || flogo != clogo)
                {
                    if (jj > start)
                    { fs.WriteLine("showpage"); break; }
                    fs.WriteLine("clear"); fs.WriteLine(fpage);
                }
                if (brs == 1 && col == 1)
                {
                    fs.WriteLine("34 0 translate");
                    if (clogo == 1)
                    { fs.WriteLine("130 965 translate"); }
                    else if (clogo == 2)
                    { fs.WriteLine("77 965 translate"); }
                    else if (clogo == 3)
                    { fs.WriteLine("52 965 translate"); }
                    else if (clogo == 4)
                    { fs.WriteLine("20 965 translate"); }
                }
                else if (col == 1)
                {
                    if (clogo == 1)
                    { fs.WriteLine("-520 -165 translate"); }
                    else if (clogo == 2)
                    { fs.WriteLine("-630 -165 translate"); }
                    else if (clogo == 3)
                    { fs.WriteLine("-680 -165 translate"); }
                    else if (clogo == 4)
                    { fs.WriteLine("-740 -165 translate"); }
                }
                if (col == 1)
                {
                    fs.WriteLine("gsave");
                    fs.WriteLine(l1 + " execform");
                    fs.WriteLine("grestore");
                    if (clogo == 2)
                    {
                        fs.WriteLine("70 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l2 + " execform");
                        fs.WriteLine("grestore");
                    }
                    if (clogo == 3)
                    {
                        fs.WriteLine("70 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l2 + " execform");
                        fs.WriteLine("grestore");
                        fs.WriteLine("70 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l3 + " execform");
                        fs.WriteLine("grestore");
                    }
                    if (clogo == 4)
                    {
                        fs.WriteLine("60 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l2 + " execform");
                        fs.WriteLine("grestore");
                        fs.WriteLine("60 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l3 + " execform");
                        fs.WriteLine("grestore");
                        fs.WriteLine("60 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l4 + " execform");
                        fs.WriteLine("grestore");
                    }
                }
                else
                {
                    if (clogo == 1)
                    { fs.WriteLine("260 0 translate"); }
                    else if (clogo == 2)
                    { fs.WriteLine("210 0 translate"); }
                    else if (clogo == 3)
                    { fs.WriteLine("130 0 translate"); }
                    else if (clogo == 4)
                    { fs.WriteLine("100 0 translate"); }
                    fs.WriteLine("gsave");
                    fs.WriteLine(l1 + " execform");
                    fs.WriteLine("grestore");
                    if (clogo == 2)
                    {
                        fs.WriteLine("70 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l2 + " execform");
                        fs.WriteLine("grestore");
                    }
                    if (clogo == 3)
                    {
                        fs.WriteLine("70 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l2 + " execform");
                        fs.WriteLine("grestore");
                        fs.WriteLine("70 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l3 + " execform");
                        fs.WriteLine("grestore");
                    }
                    if (clogo == 4)
                    {
                        fs.WriteLine("60 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l2 + " execform");
                        fs.WriteLine("grestore");
                        fs.WriteLine("60 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l3 + " execform");
                        fs.WriteLine("grestore");
                        fs.WriteLine("60 0 translate");
                        fs.WriteLine("gsave");
                        fs.WriteLine(l4 + " execform");
                        fs.WriteLine("grestore");
                    }
                }
                flogo = clogo;
                if (col == 3)
                {
                    col = 0; brs++;
                }
                if (brs == 21)
                { brs = 1; }
            }  //end for 2

        }
        void cetakHeader(StreamWriter fs, int start, DataTable ListLogo, DataTable ListXls, int jmldok, string fpage)
        {
            int flogo = 0;
            int col = 0, brs = 1;
            double posx = 10, posy = 900;
            ListLogo.Clear();
            int recs = 0;
            for (int jk = start; jk < jmldok; jk++)
            {
                col++;
                //backgroundWorker1.ReportProgress(jk + 1);
                //Data Card
                double hitbrs = 0;
                string vNama = ListXls.Rows[jk]["NAMA"].ToString().Trim();
                string vNmPT = ListXls.Rows[jk]["NMPT"].ToString().Trim();
                string vTitle1 = ListXls.Rows[jk]["TITLE"].ToString().Trim();
                string vTitle2 = ListXls.Rows[jk]["TITLE2"].ToString().Trim();
                string vPhone = ListXls.Rows[jk]["PHONE"].ToString().Trim();
                string vFax = ListXls.Rows[jk]["FAX"].ToString().Trim();
                string vHP = ListXls.Rows[jk]["HP"].ToString().Trim();
                string vEmail = ListXls.Rows[jk]["EMAIL"].ToString().Trim();
                if (vPhone != string.Empty)
                { hitbrs = hitbrs + 8; }
                if (vFax != string.Empty)
                { hitbrs = hitbrs + 8; }
                if (vHP != string.Empty)
                { hitbrs = hitbrs + 8; }
                if (vHP.Length > 55)
                { hitbrs = hitbrs + 8; }
                if (vEmail != string.Empty)
                { hitbrs = hitbrs + 8; }

                string vBuild = ListXls.Rows[jk]["BUILDING"].ToString().Trim();
                string vAddress = ListXls.Rows[jk]["ADDRESS"].ToString().Trim();
                string vWeb = ListXls.Rows[jk]["WEB"].ToString().Trim();
                int clogo = Convert.ToInt32(ListXls.Rows[jk]["JmlLogo"].ToString().Trim());


                int jp = (jk + 1) % 21;
                if (jp == 1)
                {
                    if (jk > 0)
                    {
                        fs.WriteLine("showpage");
                        //cetakMultiBG(fs, recs);
                    }
                    recs = jk;
                    fs.WriteLine("clear"); fs.WriteLine(fpage);
                    posx = 10; posy = 900;
                }
                if (brs == 1)
                { posy = 1150; }
                else if (brs == 2)
                { posy = 985; }
                else if (brs == 3)
                { posy = 820; }
                else if (brs == 4)
                { posy = 655; }
                else if (brs == 5)
                { posy = 490; }
                else if (brs == 6)
                { posy = 325; }
                else if (brs == 7)
                { posy = 160; }
                if (col == 1)
                { posx = 30; }
                else if (col == 2)
                { posx = 300; }
                else if (col == 3)
                { posx = 570; }

                double brinf = (posy - 69) + hitbrs, brbu = posy - 83, brwb = posy - 105, colnm = posx + 110;
                fs.WriteLine("FB09 (" + vNama + ") B " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                posy = posy - 15;
                fs.WriteLine("FB07 (" + vNmPT + ") B " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                posy = posy - 8;
                fs.WriteLine("FB06 (" + vTitle1 + ") B " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                if (vTitle2 != string.Empty)
                {
                    posy = posy - 8;
                    fs.WriteLine("FB06 (" + vTitle2 + ") B " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                }
                brinf = brinf - 8;
                fs.WriteLine("FA06 (T " + vPhone + ") K " + posx.ToString("###.##") + " " + brinf.ToString("###.##") + " SL");
                if (vFax != string.Empty)
                {
                    brinf = brinf - 8;
                    fs.WriteLine("FA06 (F " + vFax + ") K " + posx.ToString("###.##") + " " + brinf.ToString("###.##") + " SL");
                }
                if (vHP != string.Empty)
                {
                    vHP = "M " + vHP;
                    brinf = brinf - 8;
                    fs.WriteLine("FA06 (" + vHP + ") K " + posx.ToString("###.##") + " " + brinf.ToString("###.##") + " SL");
                }
                fs.WriteLine("FA06 (" + vBuild + ") K " + posx.ToString("###.##") + " " + brbu.ToString("###.##") + " SL");
                brbu = brbu - 8;
                fs.WriteLine("FA06 (" + vAddress + ") K " + posx.ToString("###.##") + " " + brbu.ToString("###.##") + " SL");
                fs.WriteLine("FA06 (" + vWeb + ") K " + posx.ToString("###.##") + " " + brwb.ToString("###.##") + " SL");
                flogo = clogo;
                if (col == 3)
                { col = 0; brs++; }
                if (brs == 21)
                { brs = 1; }
            } // endfor 1
            fs.WriteLine("showpage");
            fs.Close();
        }
        public void ceklist(string fhandle, DataTable tblceklist, string kdcycle, string nmfile, string fgrp)
        {
            int hitnopol = 0, hal = 0;
            double y1 = 718.5;
            double y2 = 731;
            double y3 = 722;
            using (StreamWriter sw = new StreamWriter(fhandle))
            {
                string[] bacafile = File.ReadAllLines(Directory.GetCurrentDirectory() + "\\HEAD-PS.HDR");
                foreach (string perbaris in bacafile)
                {
                    string bagus = perbaris;
                    sw.WriteLine(bagus);
                }
                int jmlpol = tblceklist.Rows.Count;
                Decimal jhal = 0;
                string cnoseq = "", cnopol = "", ccycle = "", cbarcode = "";
                string cnamatt = "", ccabang = "";
                double bar1 = 0, bar2 = 0, bar3 = 3;
                for (int rec = 0; rec < jmlpol; rec++)
                {
                    ccycle = kdcycle;
                    cnoseq = (rec + 1).ToString("D4");
                    cnopol = cnoseq;
                    cbarcode = cnoseq;
                    cnamatt = tblceklist.Rows[rec]["NAMA"].ToString().Trim();
                    int pjj = cnamatt.Length;
                    if (pjj > 35)
                    { cnamatt = cnamatt.Substring(0, 35); }
                    ccabang = "";


                    hitnopol++;
                    if (hitnopol % 50 == 1)
                    {
                        if (hitnopol > 50)
                        {
                            sw.WriteLine("0 550 -0.5  20  15.00 BO");
                            sw.WriteLine("0 550 -0.5  20  35.00 BO");
                            sw.WriteLine("0 550 -0.5  20  55.00 BO");
                            sw.WriteLine("0 550 -0.5  20  75.00 BO");
                            sw.WriteLine("0 550 -0.5  20  95.00 BO");
                            sw.WriteLine("0 550 -0.5  20  115.00 BO");
                            sw.WriteLine("0 0.5 100  20  15.00 BO");
                            sw.WriteLine("0 0.5 100  270 15.00 BO");
                            sw.WriteLine("0 0.5 100  420 15.00 BO");
                            sw.WriteLine("0 0.5 100  570 15.00 BO");
                            sw.WriteLine("FA07 (NAMA PETUGAS PREPARE :) K 25.00 102.50 SL");
                            sw.WriteLine("FA07 (NAMA PETUGAS SCANNING :) K 25.00 82.50 SL");
                            sw.WriteLine("FA07 (NAMA PETUGAS BINDING :) K 25.00 62.50 SL");
                            sw.WriteLine("FA07 (NAMA PETUGAS QC POLIS :) K 25.00 42.50 SL");
                            sw.WriteLine("FA07 (NAMA PETUGAS FINISHING :) K 25.00 22.50 SL");
                            sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 102.50 SL");
                            sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 82.50 SL");
                            sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 62.50 SL");
                            sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 42.50 SL");
                            sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 22.50 SL");
                            sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 102.50 SL");
                            sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 82.50 SL");
                            sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 62.50 SL");
                            sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 42.50 SL");
                            sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 22.50 SL");
                            sw.WriteLine("showpage");
                            hitnopol = 1;
                            y1 = 718;
                            y2 = 731;
                            y3 = 722;
                        }
                        hal++;
                        sw.WriteLine("%%Page:" + hal.ToString("###"));
                        sw.WriteLine("clear");
                        sw.WriteLine("<< /PageSize [596 840] /MediaColor (white) /MediaWeight 80.000000 /MediaType (Plain) /Duplex false >> setpagedevice << /OutputType () >> setpagedevice");
                        sw.WriteLine("FA07 (CHK/PRO/02/00) K 525.00 825.00 SL");
                        sw.WriteLine("FA10 (CHECKLIST KARTU NAMA "+fgrp+") K 305.00 810.00 SC");
                        sw.WriteLine("(CHECKLIST KARTU NAMA " + fgrp + ") UL");
                        sw.WriteLine("FA07 (CYCLE : " + kdcycle + ") K 40.00 790.00 SL");
                        sw.WriteLine("FA07 (HAL  : " + hal.ToString("###") + " dari " + jhal.ToString("###") + ") K 40.00 780.00 SL");
                        sw.WriteLine("FA07 (FILE : " + nmfile + ") K 40.00 765.00 SL");
                        sw.WriteLine("FA07 (TOTAL SEMUA DATA : " + jmlpol.ToString("###,###") + ") K 400.00 790.00 SL");

                        sw.WriteLine("0 550 -0.5  20  755.00 BO");
                        sw.WriteLine("0 255 -0.5  315 742.50 BO");
                        sw.WriteLine("0 0.5 -25   20  755.00 BO");
                        sw.WriteLine("0 0.5 -25   45  755.00 BO");
                        sw.WriteLine("0 0.5 -25   75  755.00 BO");
                        //sw.WriteLine("0 0.5 -25   115 755.00 BO");
                        sw.WriteLine("0 0.5 -25   208 755.00 BO");
                        sw.WriteLine("0 0.5 -25   315 755.00 BO");
                        sw.WriteLine("0 0.5 -12.5 340 742.00 BO");
                        sw.WriteLine("0 0.5 -12.5 370 742.00 BO");
                        sw.WriteLine("0 0.5 -25   398 755.00 BO");
                        sw.WriteLine("0 0.5 -12.5 417 742.50 BO");
                        sw.WriteLine("0 0.5 -12.5 438 742.50 BO");
                        sw.WriteLine("0 0.5 -12.5 454 742.50 BO");
                        sw.WriteLine("0 0.5 -12.5 473 742.00 BO");
                        sw.WriteLine("0 0.5 -12.5 487 742.50 BO");
                        sw.WriteLine("0 0.5 -25   505 755.00 BO");
                        sw.WriteLine("0 0.5 -25   570 755.00 BO");
                        sw.WriteLine("0 550 -0.5  20  730.00 BO");
                        sw.WriteLine("FA07 (NO) K 32.50 745.00 SC");
                        sw.WriteLine("FA07 (URUT) K 32.50 735.00 SC");
                        sw.WriteLine("FA07 (NO) K 61.00 745.00 SC");
                        sw.WriteLine("FA07 (IDT) K 61.00 735.00 SC");
                        //sw.WriteLine("FA07 (NO) K 95.00 745.00 SC");
                        //sw.WriteLine("FA07 (IDT) K 95.00 735.00 SC");
                        sw.WriteLine("FA07 (N A M A) K 155 745.00 SC");
                        sw.WriteLine("FA07 (NAMA SO) K 265.00 745.00 SC");
                        sw.WriteLine("FA07 (STATUS PENGERJAAN) K 358.50 745.00 SC");
                        sw.WriteLine("FA07 (DETAIL CARD PAGES) K 455.50 745.00 SC");
                        sw.WriteLine("FA07 (KETERANGAN) K 535.50 745.00 SC");
                        sw.WriteLine("FA07 (SCAN) K 328.50 735.00 SC");
                        sw.WriteLine("FA07 (BINDING) K 355.50 735.00 SC");
                        sw.WriteLine("FA07 (QC POL) K 384.50 735.00 SC");
                        sw.WriteLine("FA07 (BND) K 407.50 735.00 SC");
                        sw.WriteLine("FA07 (RP) K 428.50 735.00 SC");
                        sw.WriteLine("FA07 (WL) K 447.50 735.00 SC");
                        sw.WriteLine("FA07 (TS) K 463.50 735.00 SC");
                        sw.WriteLine("FA07 (SSP) K 480.50 735.00 SC");
                        sw.WriteLine("FA07 (MAP) K 496.50 735.00 SC");
                    }
                    bar1 = y1 - ((hitnopol - 1) * 12);
                    bar2 = y2 - ((hitnopol - 1) * 12);
                    bar3 = y3 - ((hitnopol - 1) * 12);
                    sw.WriteLine("0 550 -0.5 20 " + bar1.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12  20 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12  45 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12  75 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    //sw.WriteLine("0 0.5 -12 115 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 208 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 315 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 340 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 370 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 398 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 417 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 438 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 454 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 473 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 487 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 505 " + bar2.ToString("###.##").Replace(",", ".") + " BO");
                    sw.WriteLine("0 0.5 -12 570 " + bar2.ToString("###.##").Replace(",", ".") + " BO");

                    sw.WriteLine("FA07 (" + cnoseq + ") K 41.00 " + bar3.ToString("###.##").Replace(",", ".") + " SR");
                    sw.WriteLine("FA06 (" + cnopol + ") K 48.00 " + bar3.ToString("###.##").Replace(",", ".") + " SL");
                    sw.WriteLine("FA06 (" + cnamatt + ") K 81 " + bar3.ToString("###.##").Replace(",", ".") + " SL");
                    sw.WriteLine("FA06 (" + ccabang + ") K 210.00 " + bar3.ToString("###.##").Replace(",", ".") + " SL");
                }
                sw.WriteLine("0 550 -0.5  20  15.00 BO");
                sw.WriteLine("0 550 -0.5  20  35.00 BO");
                sw.WriteLine("0 550 -0.5  20  55.00 BO");
                sw.WriteLine("0 550 -0.5  20  75.00 BO");
                sw.WriteLine("0 550 -0.5  20  95.00 BO");
                //sw.WriteLine("0 550 -0.5  20  115.00 BO");
                sw.WriteLine("0 0.5 100  20  15.00 BO");
                sw.WriteLine("0 0.5 100  270 15.00 BO");
                sw.WriteLine("0 0.5 100  420 15.00 BO");
                sw.WriteLine("0 0.5 100  570 15.00 BO");
                sw.WriteLine("FA07 (NAMA PETUGAS PREPARE :) K 25.00 102.50 SL");
                sw.WriteLine("FA07 (NAMA PETUGAS SCANNING :) K 25.00 82.50 SL");
                sw.WriteLine("FA07 (NAMA PETUGAS BINDING :) K 25.00 62.50 SL");
                sw.WriteLine("FA07 (NAMA PETUGAS QC POLIS :) K 25.00 42.50 SL");
                sw.WriteLine("FA07 (NAMA PETUGAS FINISHING :) K 25.00 22.50 SL");
                sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 102.50 SL");
                sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 82.50 SL");
                sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 62.50 SL");
                sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 42.50 SL");
                sw.WriteLine("FA07 (TGL PENGERJAAN :) K 275.00 22.50 SL");
                sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 102.50 SL");
                sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 82.50 SL");
                sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 62.50 SL");
                sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 42.50 SL");
                sw.WriteLine("FA07 (JML DOKUMEN YG ADA :) K 425.00 22.50 SL");
                sw.WriteLine("showpage");
            }
        }

    }
}
