using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Configuration.Assemblies;
using System.Collections;
using System.Management;
using System.Management.Instrumentation;
using System.Security.Cryptography;

namespace AXA_NAME_CARD
{
    public partial class Form1 : Form
    {
        OleDbConnection kon = new OleDbConnection();
        DataTable dt = new DataTable();
        DataTable dtmst = new DataTable();
        int jmldok = 0; DataTable ListXls = new DataTable(), ListLogo = new DataTable(); string tglcyc ="", pathdir = "",nmfile="";
        //string fpageA3 = "<< /PageSize [842 1191] /MediaColor (blue) /MediaWeight 80.000000 /MediaType (Plain) /Duplex false >> setpagedevice << /OutputType () >> setpagedevice";
        string fpage = "<< /PageSize [906.9 1276.05] /MediaColor (blue) /MediaWeight 80.000000 /MediaType (Plain) /Duplex false >> setpagedevice << /OutputType () >> setpagedevice";
        axaproc proc = new axaproc();
        public Form1()
        {
            InitializeComponent();
            tglcyc = DateTime.Now.ToString("yyyyMMdd");
            pathdir = Directory.GetCurrentDirectory() + "\\";
            ListLogo.Columns.Add("JmlLogo", typeof(int));
            ListLogo.Columns.Add("LOGO1");
            ListLogo.Columns.Add("LOGO2");
            ListLogo.Columns.Add("LOGO3");
            ListLogo.Columns.Add("LOGO4");
            button1.Enabled = false;
            if (File.Exists("SerialNum.txt"))
            {
                string sn = GetMotherBoardID();
                string serial = proc.Encrypt(sn, "-");
                string[] bacaserial = File.ReadAllLines("SerialNum.txt");
                int pjg = bacaserial.Length;
                string isitxt = "";
                for (int jx=0;jx<pjg;jx++)
                {
                    isitxt = isitxt+bacaserial[jx].ToString().Trim();
                }
                string decserial = proc.Decrypt(serial, "-");
                if (decserial == sn)
                {
                    button1.Enabled = true;
                    label5.Text = "Registered";
                    button3.Enabled = false;
                }
            }
        }
 
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            ListXls = proc.excelapprove(txtInput.Text.Trim());
            jmldok = ListXls.Rows.Count;
            button1.Enabled = false;
            button2.Enabled = false;
            if (!this.backgroundWorker1.IsBusy)
            {
                this.backgroundWorker1.RunWorkerAsync();
            }
            button1.Enabled = true;
            button2.Enabled = true;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //Delete rows empty
            foreach (DataRow dr in ListXls.Rows)
            {
                if (dr[1].ToString() == string.Empty)
                    dr.Delete();
            }
            ListXls.AcceptChanges();
            jmldok = ListXls.Rows.Count;
            string fproc = "", fbacktemp = "", fback = "", fbacklogo = "",fbackfront="",vWeb="", fbuild="", vPT="",vjns="",vgrp="";
            int jmst = dtmst.Rows.Count;
            for (int fm = 0; fm < jmst; fm++)
            {
                vgrp = dtmst.Rows[fm]["GRP_CARD"].ToString().Trim();
                vjns = dtmst.Rows[fm]["JNS_CARD"].ToString().Trim();
                fproc = dtmst.Rows[fm]["PROC_TYPE"].ToString().Trim();
                fbacktemp = dtmst.Rows[fm]["BACKTEMP"].ToString().Trim();
                fback = dtmst.Rows[fm]["FBACK"].ToString().Trim();
                fbacklogo = dtmst.Rows[fm]["BACKLOGO"].ToString().Trim();
                fbackfront = dtmst.Rows[fm]["BACKFRONT"].ToString().Trim();
                vWeb = dtmst.Rows[fm]["WEBNAME"].ToString().Trim();
                fbuild = dtmst.Rows[fm]["FBUILD"].ToString().Trim();
                vPT = dtmst.Rows[fm]["GRP_NAME"].ToString().Trim();
            }
            jmldok = ListXls.Rows.Count;
            nmfile = Path.GetFileName(txtInput.Text.Replace(".xlsx", "").Replace(".XLSX", ""));
            string cetak = pathdir + "CETAK-CARD\\" + vgrp + "\\" + tglcyc + "\\" + vjns + "\\";
            proc.buatdir(cetak);
            //Buat Ceklist
            string flcek = cetak  +"CEKLIST-CARD-"+ vgrp + "-" + tglcyc + "-" + vjns +".ps";
            proc.ceklist(flcek, ListXls, tglcyc, nmfile.Replace(".XLS","").Replace(".xls",""), vgrp);


            if (fproc == "F01")
            {
                if (fbacklogo == "True")
                {
                    ListXls.Columns.Add("JmlLogo", typeof(int));
                    for (int jj = 0; jj < jmldok; jj++)
                    {
                        int jml = 0;
                        string l1 = ListXls.Rows[jj]["LOGO1"].ToString().Trim();
                        string l2 = ListXls.Rows[jj]["LOGO2"].ToString().Trim();
                        string l3 = ListXls.Rows[jj]["LOGO3"].ToString().Trim();
                        string l4 = ListXls.Rows[jj]["LOGO4"].ToString().Trim();
                        if (l1 != string.Empty)
                        { jml++; }
                        if (l2 != string.Empty)
                        { jml++; }
                        if (l3 != string.Empty)
                        { jml++; }
                        if (l4 != string.Empty)
                        { jml++; }
                        ListXls.Rows[jj]["JmlLogo"] = jml;
                    }
                    ListXls.AcceptChanges();
                    if (ckMulti.Checked == false)
                    {
                        DataView dv = new DataView();
                        dv = ListXls.DefaultView;
                        dv.Sort = "JmlLogo"; ///urutkan data by Jmlogo
                        ListXls = dv.ToTable();
                    }
                }
                string vhead = pathdir + "HEAD-" + vgrp + ".HDR";
                string pathlogo = pathdir + "TEMPLATES\\LOGO\\";
                string pathtmpl = pathdir + "TEMPLATES\\";
                string pstmp = Directory.GetCurrentDirectory() + "ftmp.ps";
                if (jmldok > 0)
                {
                    string pscetak = cetak + "NC-" +vjns+"-" + tglcyc + "_" + nmfile + ".PS";
                    using (System.IO.StreamWriter fs = new System.IO.StreamWriter(pscetak, false))
                    {
                        proc.embed(fs, vhead);
                        proc.addtmpl(fs, pathtmpl, "N", fbackfront);
                        if (fbacklogo == "True")
                        {
                            proc.addtmpl(fs, pathlogo,"Y","");
                        }
                        fs.WriteLine("%%%END-TMPL");
                        int flogo = 0;
                        int col = 0, brs = 1;
                        double posx = 10, posy = 900, cekpos=0;
                        ListLogo.Clear();
                        int recs = 0;
                        for (int jk = 0; jk < jmldok; jk++)
                        {
                            col++;
                            backgroundWorker1.ReportProgress(jk + 1);
                            //Data Card
                            double hitbrs = 0;
                            string vNama = ListXls.Rows[jk]["NAMA"].ToString().Trim();
                            string vNmPT = ListXls.Rows[jk]["Nama Perusahaan"].ToString().Trim();
                            string vktit = ListXls.Rows[jk]["KODE JABATAN"].ToString().Trim();
                            string vstate = ListXls.Rows[jk]["KODE NEGARA"].ToString().Trim();
                            string vaddr1 = ListXls.Rows[jk]["ALAMAT 1"].ToString().Trim();
                            string vaddr2 = ListXls.Rows[jk]["ALAMAT 2"].ToString().Trim();
                            string vaddr3 = ListXls.Rows[jk]["ALAMAT 3"].ToString().Trim();
                            string vTitle1 = ListXls.Rows[jk]["JABATAN 1"].ToString().Trim();
                            string vTitle2 = ListXls.Rows[jk]["JABATAN 2"].ToString().Trim();
                            string vPhone = ListXls.Rows[jk]["TELEPON 1"].ToString().Trim();
                            string vPhone2 = ListXls.Rows[jk]["TELEPON 2"].ToString().Trim();
                            string vFax = ListXls.Rows[jk]["FAX 1"].ToString().Trim();
                            string vFax2 = ListXls.Rows[jk]["FAX 2"].ToString().Trim();
                            string vHP = ListXls.Rows[jk]["HP 1"].ToString().Trim();
                            string vHP2 = ListXls.Rows[jk]["HP 2"].ToString().Trim();
                            string vSo = ListXls.Rows[jk]["Kode So"].ToString().Trim();
                            string vAgen = ListXls.Rows[jk]["Agent Code"].ToString().Trim();
                            string vEmail = ListXls.Rows[jk]["EMAIL 1"].ToString().Trim();
                            string vEmail2 = ListXls.Rows[jk]["EMAIL 2"].ToString().Trim();
                            string vBuild = vaddr1;
                            if (vFax == "0")
                            { vFax = ""; }
                            if (vFax2 == "0")
                            { vFax = ""; }
                            if (vPhone2 == "0")
                            { vPhone2 = ""; }
                            string vAlamat = "";
                            if (fbuild == "True")
                            { vAlamat = vaddr2 + " " + vaddr3; }
                            else
                            { vAlamat = vaddr1 + " " + vaddr2 + " " + vaddr3; vBuild = ""; }
                            vAlamat = vAlamat.Replace("  ", " ");
                            if (vEmail2 != string.Empty)
                            { vEmail = vEmail + "/" + vEmail2; }
                            if (vHP2 != string.Empty)
                            { vHP = "|M." + vHP + "-" + vHP2; }
                            else
                            { vHP = "|M." + vHP; }
                            if (vPhone2 != string.Empty)
                            {
                                vPhone = "T." + vPhone + "-" + vPhone2;
                                hitbrs = hitbrs + 8;
                            }
                            else
                            { vPhone = "T." + vPhone; }
                            if (vFax != string.Empty)
                            {
                                vFax = "|F." + vFax;
                                if (vFax2 != string.Empty)
                                { vFax = vFax + "-" + vFax2; }
                                hitbrs = hitbrs + 8; 
                            }
                            int clogo = 0;
                            if (fbacklogo == "True")
                            {
                               clogo =  Convert.ToInt32(ListXls.Rows[jk]["JmlLogo"].ToString().Trim());
                            }
                            int jp = (jk + 1) % 21;
                            if (ckMulti.Checked)
                            {
                                if (jp == 1)
                                {
                                    if (jk > 0)
                                    {
                                        fs.WriteLine("showpage");
                                        if (fbacklogo == "True")
                                        {cetakMultiBG(fs, recs);}
                                    }
                                    recs = jk;
                                    fs.WriteLine("clear"); fs.WriteLine(fpage);
                                    fs.WriteLine("gsave");
                                    fs.WriteLine("temp_"+fbackfront+ " execform");
                                    fs.WriteLine("grestore");
                                    fs.WriteLine("61 0 translate");
                                }
                            }
                            if (brs == 1)
                            { posy = 1160; cekpos = posy; }
                            else
                            {posy = cekpos;}
                            if (col == 1)
                            { posx = 30; }
                            else if (col == 2)
                            { posx = 285; }
                            else if (col == 3)
                            { posx = 541; }
                            double brinf = (posy - 69) + hitbrs, brbu = posy - 83, brwb = posy - 105, colnm = posx + 110;
                            fs.WriteLine("F10PB (" + vNama + ") BR " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            posy = posy - 10;
                            fs.WriteLine("F07SB (" + vTitle1 + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (vTitle2 != string.Empty)
                            {
                                posy = posy - 8;
                                fs.WriteLine("F07SB (" + vTitle2 + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                                posx = posy + 8;
                            }
                            string vKomunikasi = vPhone + vFax + vHP;
                            //posy = posy - 30;
                            string[] tAddr = proc.perkata(vAlamat, 2, 70);
                            posy = posy - 9;
                            int pjt = tAddr[0].Length;
                            posy = cekpos - 82;  //start from bottom
                            fs.WriteLine("F07SS (" + vWeb + ") BR " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            posy = posy + 9;
                            fs.WriteLine("F07SR (" + vEmail + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            posy = posy + 8;
                            fs.WriteLine("F07SR (" + vKomunikasi + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (tAddr[1].Trim() != string.Empty)
                            {
                                posy = posy + 8;
                                fs.WriteLine("F07SR (" + tAddr[1] + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            }
                            posy = posy + 8;
                            fs.WriteLine("F07SR (" + tAddr[0] + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (fbuild == "True")
                            {
                                posy = posy + 8;
                                fs.WriteLine("F07SR (" + vBuild + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            }
                            posy = posy + 9;
                            fs.WriteLine("F07SB (" + vNmPT + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");

                            flogo = clogo;
                            if (col == 3)
                            { col = 0; brs++; posy = cekpos - 167; cekpos = posy; }
                            if (brs==8)
                            {
                                brs = 1; col = 0; posx = 10; posy = 900; 
                            }
                        } // endfor 1
                        fs.WriteLine("showpage");
                        col = 0; brs = 1;
                        if (fbacklogo == "True")
                        {
                            if (ckMulti.Checked)
                            { cetakMultiBG(fs, recs); }
                        }
                        fs.Close();
                    }
                }
            }
            else if (fproc == "S01")
            {
                string vhead = pathdir + "HEAD-" + vgrp + ".HDR";
                string pathtmpl = pathdir + "TEMPLATES\\";
                string pstmp = Directory.GetCurrentDirectory() + "ftmp.ps";
                if (jmldok > 0)
                {
                    string pscetak = cetak + "NC-" + vjns + "-" + tglcyc + "_" + nmfile + ".PS";
                    using (System.IO.StreamWriter fs = new System.IO.StreamWriter(pscetak, false))
                    {
                        proc.embed(fs, vhead);
                        proc.addtmpl(fs, pathtmpl, "N", fbackfront);
                        fs.WriteLine("%%%END-TMPL");
                        int col = 0, brs = 1;
                        double posx = 10, posy = 900, cekpos=0;
                        int recs = 0;
                        for (int jk = 0; jk < jmldok; jk++)
                        {
                            col++;
                            backgroundWorker1.ReportProgress(jk + 1);
                            //Data Card
                            double hitbrs = 0;
                            string vNama = ListXls.Rows[jk]["NAMA"].ToString().Trim();
                            string vNmPT = ListXls.Rows[jk]["Nama Perusahaan"].ToString().Trim();
                            string vstate = ListXls.Rows[jk]["KODE NEGARA"].ToString().Trim();
                            string vaddr1 = ListXls.Rows[jk]["ALAMAT 1"].ToString().Trim();
                            string vaddr2 = ListXls.Rows[jk]["ALAMAT 2"].ToString().Trim();
                            string vTitle1 = ListXls.Rows[jk]["JABATAN 1"].ToString().Trim();
                            string vPhone = ListXls.Rows[jk]["TELEPON 1"].ToString().Trim();
                            string vSO = ListXls.Rows[jk]["Kode SO"].ToString().Trim();
                            string vExt = ListXls.Rows[jk]["EXT"].ToString().Trim();
                            string vFax = ListXls.Rows[jk]["FAX 1"].ToString().Trim();
                            string vFax2 = ListXls.Rows[jk]["FAX 2"].ToString().Trim();
                            string vHP = ListXls.Rows[jk]["HP 1"].ToString().Trim();
                            string vHP2 = ListXls.Rows[jk]["HP 2"].ToString().Trim();
                            string vEmail = ListXls.Rows[jk]["EMAIL 1"].ToString().Trim();
                            string vEmail2 = ListXls.Rows[jk]["EMAIL 2"].ToString().Trim();
                            string vAlamat = "";
                            string vBuild = vaddr1;
                            { vAlamat = vaddr2; }
                            vAlamat = vAlamat.Replace("  ", " ");
                            vHP = "M " + vHP; 
                            vPhone = "T " + vPhone;
                            if (vExt != string.Empty)
                            { vPhone = vPhone + " Ext. " + vExt; }
                            vFax = "F " + vFax;
                            int jp = (jk + 1) % 21;
                            if (jp == 1)
                            {
                                if (jk > 0)
                                {
                                    fs.WriteLine("showpage");
                                    if (fbacklogo == "True")
                                    { cetakMultiBG(fs, recs); }
                                }
                                recs = jk;
                                fs.WriteLine("clear"); fs.WriteLine(fpage);
                                fs.WriteLine("gsave");
                                fs.WriteLine("temp_" + fbackfront + " execform");
                                fs.WriteLine("grestore");
                                fs.WriteLine("61 0 translate");
                            }
                            if (brs == 1)
                            { posy = 1180; cekpos = posy; }
                            else
                            {
                                posy = cekpos;
                            }
                            if (col == 1)
                            { posx = 30; }
                            else if (col == 2)
                            { posx = 285; }
                            else if (col == 3)
                            { posx = 541; }
                            double brinf = (posy - 69) + hitbrs, brbu = posy - 83, brwb = posy - 105, colnm = posx + 90;
                            fs.WriteLine("F10PB (" + vNama + ") BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            posy = posy - 16;
                            fs.WriteLine("F07SB (" + vNmPT + ") BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            posy = posy - 10;
                            fs.WriteLine("F07SB (" + vTitle1 + ") BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            string[] tAddr = proc.perkata(vAlamat, 2, 70);
                            int pjt = tAddr[0].Length;
                            posy = cekpos - 110;
                            fs.WriteLine("F07SR (" + vWeb + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (tAddr[1].Trim() != string.Empty)
                            {
                                posy = posy + 9;
                                fs.WriteLine("F07SR (" + tAddr[1] + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                                posy = posy - 1;
                            }
                            posy = posy + 9;
                            fs.WriteLine("F07SR (" + tAddr[0] + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (fbuild == "True")
                            {
                                posy = posy + 8;
                                fs.WriteLine("F07SR (" + vBuild + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            }
                            posy = posy + 15;
                            fs.WriteLine("F07SR (" + vEmail + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            posy = posy + 8;
                            fs.WriteLine("F07SR (" + vHP + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (vFax != string.Empty)
                            {
                                posy = posy + 8;
                                fs.WriteLine("F07SR (" + vFax + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            }
                            posy = posy + 8;
                            fs.WriteLine("F07SR (" + vPhone + ") K " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (col == 3)
                            { col = 0; brs++; posy = cekpos - 160; cekpos = posy; }
                            if (brs == 8)
                            {
                                brs = 1; col = 0; posx = 10; posy = 900;
                            }
                        } // endfor 1
                        fs.WriteLine("showpage");
                        col = 0; brs = 1;
                        if (fbacklogo == "True")
                        {
                            if (ckMulti.Checked)
                            { cetakMultiBG(fs, recs); }
                        }
                        fs.Close();
                    }
                }
            }
            else if (fproc == "A01" || fproc =="M01")
            {
                string vhead = pathdir + "HEAD-"+vgrp+".HDR";
                string pathtmpl = pathdir + "TEMPLATES\\";
                string pstmp = Directory.GetCurrentDirectory() + "ftmp.ps";
                if (jmldok > 0)
                {
                    string pscetak = cetak + "NC-" + vjns + "-" + tglcyc + "_" + nmfile + ".PS";
                    using (System.IO.StreamWriter fs = new System.IO.StreamWriter(pscetak, false))
                    {
                        proc.embed(fs, vhead);
                        proc.addtmpl(fs, pathtmpl, "N", fbackfront);
                        fs.WriteLine("%%%END-TMPL");
                        int col = 0, brs = 1;
                        double posx = 83, posy = 1129, cekpos = 0 ;
                        int recs = 0;
                        for (int jk = 0; jk < jmldok; jk++)
                        {
                            col++;
                            backgroundWorker1.ReportProgress(jk + 1);
                            //Data Card
                            double hitbrs = 0;
                            string vNama = ListXls.Rows[jk]["NAMA"].ToString().Trim();
                            string vNmPT = vPT;
                            string vstate = ListXls.Rows[jk]["KODE NEGARA"].ToString().Trim();
                            string vaddr1 = ListXls.Rows[jk]["ALAMAT 1"].ToString().Trim();
                            string vaddr2 = ListXls.Rows[jk]["ALAMAT 2"].ToString().Trim();
                            string vaddr3 = ListXls.Rows[jk]["ALAMAT 3"].ToString().Trim();                            
                            string vTitle1 = ListXls.Rows[jk]["JABATAN 1"].ToString().Trim();
                            string vTitle2 = ListXls.Rows[jk]["JABATAN 2"].ToString().Trim();
                            string vPhone = ListXls.Rows[jk]["TELEPON 1"].ToString().Trim();
                            string vPhone2 = ListXls.Rows[jk]["TELEPON 2"].ToString().Trim();
                            //string vSO = "";
                            string vExt = ListXls.Rows[jk]["EXT"].ToString().Trim();
                            string vFax = ListXls.Rows[jk]["FAX 1"].ToString().Trim();
                            string vFax2 = ListXls.Rows[jk]["FAX 2"].ToString().Trim();
                            string vHP = ListXls.Rows[jk]["HP 1"].ToString().Trim();
                            string vHP2 = ListXls.Rows[jk]["HP 2"].ToString().Trim();
                            string vEmail = ListXls.Rows[jk]["EMAIL 1"].ToString().Trim();
                            string vEmail2 = ListXls.Rows[jk]["EMAIL 2"].ToString().Trim();
                            string vAlamat = "";
                            string vBuild = vaddr1;
                            { vAlamat = vaddr2; }
                            vAlamat = vAlamat.Replace("  ", " ");
                            if (vExt != string.Empty)
                            { vPhone = vPhone + "Ext. " + vExt; }
                            int jp = (jk + 1) % 21;
                            if (jp == 1)
                            {
                                if (jk > 0)
                                {
                                    fs.WriteLine("showpage");
                                    if (fbacklogo == "True")
                                    { cetakMultiBG(fs, recs); }
                                }
                                recs = jk;
                                fs.WriteLine("clear"); fs.WriteLine(fpage);
                                fs.WriteLine("gsave");
                                fs.WriteLine("temp_" + fbackfront + " execform");
                                fs.WriteLine("grestore");
                                fs.WriteLine("%%0 0 translate"); //60 0
                            }
                            if (brs == 1)
                            { posy = 1129;cekpos = posy;}
                            else 
                            {
                                posy = cekpos; 
                            }
                            if (col == 1)
                            { posx = 83; }  //30
                            else if (col == 2)
                            { posx = 338; } //285
                            else if (col == 3)
                            { posx = 593; } //541
                            double brinf = (posy - 69) + hitbrs, brbu = posy - 83, brwb = posy - 105, colnm = posx + 110;
                            fs.WriteLine("F8MR (" + vNama + ") BR " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            posy = posy - 8;
                            fs.WriteLine("F8MR (" + vTitle1 + ") BR " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (vTitle2 != string.Empty)
                            {
                                posy = posy - 8;
                                fs.WriteLine("F8MR (" + vTitle2 + ") BR " + posx.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            }
                            int tmax=40;
                            if (fproc == "M01")
                            { tmax = 50; }
                            string[] tAddr = proc.perkata(vaddr2, 2, tmax);
                            int pjt = tAddr[0].Length;
                            posy = cekpos - 91;
                            fs.WriteLine((colnm).ToString("###.##") + " " + posy.ToString("###.##") + " MV");
                            fs.WriteLine("basefont [7 0 0 8 0 0] makefont setfont");
                            fs.WriteLine("("+vEmail + ") show");
                            posy = posy + 8;
                            fs.WriteLine("F8MR (Mobile.) BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            fs.WriteLine("F8MR (" + vHP + ") BR " + (colnm + 31).ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (vFax != string.Empty)
                            {
                                posy = posy + 8;
                                fs.WriteLine("F8MR (Fax.) BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                                fs.WriteLine("F8MR (" + vFax + ") BR " + (colnm + 31).ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            }
                            if (vPhone2 != string.Empty)
                            { posy = posy + 8; fs.WriteLine("F8MR (" + vPhone2 + ") BR " + (colnm + 31).ToString("###.##") + " " + posy.ToString("###.##") + " SL"); }
                            posy = posy + 8;
                            fs.WriteLine("F8MR (Tel.) BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            fs.WriteLine("F8MR (" + vPhone + ") BR " + (colnm + 31).ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (vaddr3.Trim() != string.Empty)
                            {
                                posy = posy + 8;
                                fs.WriteLine("F8MR (" + vaddr3 + ") BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            }
                            if (tAddr[1].Trim() != string.Empty)
                            {
                                posy = posy + 8;
                                fs.WriteLine("F8MR (" + tAddr[1] + ") BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            }
                            if (tAddr[0].Trim() != string.Empty)
                            {
                                posy = posy + 8;
                                fs.WriteLine("F8MR (" + tAddr[0] + ") BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            }
                            posy = posy + 8;
                            fs.WriteLine("F8MR (" + vBuild + ") BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            posy = posy + 16;
                            fs.WriteLine("F8MR (" + vNmPT + ") BR " + colnm.ToString("###.##") + " " + posy.ToString("###.##") + " SL");
                            if (col == 3)
                            { col = 0; brs++; posy = cekpos - 156; cekpos = posy; }
                            if (brs == 8)
                            { brs = 1; col = 0; posx = 83; posy = 1129;}
                            //if (col + 1 == 3)
                            //{ posy = cekpos - 140; cekpos = posy; }
                        } // endfor 1
                        fs.WriteLine("showpage");
                        col = 0; brs = 1;
                        if (fbacklogo == "True")
                        {
                            if (ckMulti.Checked)
                            { cetakMultiBG(fs, recs); }
                        }
                        fs.Close();
                    }
                }
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int persen = (e.ProgressPercentage * 100) / jmldok;
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Maximum = Convert.ToInt32(jmldok);
            progressBar1.Value = e.ProgressPercentage;
            label2.Text = Convert.ToString(persen) + "%";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                button1.Enabled = true;
                MessageBox.Show("Creating NameCard Finished", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information); 
            }
        }
 

        private void txtInput_Click(object sender, EventArgs e)
        {
            string dirasli = pathdir  + "DATA";
            string cektext = "", cekfilter = "";
            cektext = "xlsx"; cekfilter = "xlsx files (*.xlsx;*.xls)|*.xlsx;*.xls";

            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = dirasli,
                Title = "Browse Xls Name Card List",

                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = cektext,
                Filter = cekfilter,
                FilterIndex = 2,
                RestoreDirectory = true,

                ReadOnlyChecked = false,
                ShowReadOnly = false
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtInput.Text = openFileDialog1.FileName;
            }  
        }

        void cetakMultiBG(StreamWriter fs, int start)
        {
            int flogo = 0;
            int col = 0, brs = 1, tx = 0; 
            int dfx = 0;
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
                if (jp == 1)
                {
                    if (jj > start)
                    { fs.WriteLine("showpage"); break; }
                    fs.WriteLine("clear"); fs.WriteLine(fpage);
                }
                dfx = 60;
                dfx = (dfx * (4 - clogo)) / 2;
                if (col == 1)
                {
                    tx = 20 + dfx;
                }
                string ftrans = dfx.ToString("###") + " 0 translate";
                if (ftrans == " 0 translate")
                { ftrans = ""; }
                if (brs == 1 && col == 1)
                {
                    fs.WriteLine("34 0 translate");
                    fs.WriteLine("%%% Baris 1"); 
                    fs.WriteLine(tx.ToString("###")+" 965 translate"); 
                }
                else if (col == 1)
                {
                    fs.WriteLine("%%% Baris "+brs.ToString());
                    fs.WriteLine("-740 -165 translate");
                    fs.WriteLine(ftrans);
                }
                if (col != 1)
                {
                    fs.WriteLine(ftrans);
                    fs.WriteLine("100 0 translate");
                }
                fs.WriteLine("%%% "+col.ToString());
                if (clogo == 1)
                {
                    fs.WriteLine("gsave");
                    fs.WriteLine(l1 + " execform");
                    fs.WriteLine("grestore");
                    fs.WriteLine(ftrans);
                }
                if (clogo == 2)
                {
                    fs.WriteLine("gsave");
                    fs.WriteLine(l1 + " execform");
                    fs.WriteLine("grestore");
                    fs.WriteLine("60 0 translate");
                    fs.WriteLine("gsave");
                    fs.WriteLine(l2 + " execform");
                    fs.WriteLine("grestore");
                    fs.WriteLine(ftrans);
                }
                if (clogo == 3)
                {
                    fs.WriteLine("gsave");
                    fs.WriteLine(l1 + " execform");
                    fs.WriteLine("grestore");
                    fs.WriteLine("60 0 translate");
                    fs.WriteLine("gsave");
                    fs.WriteLine(l2 + " execform");
                    fs.WriteLine("grestore");
                    fs.WriteLine("60 0 translate");
                    fs.WriteLine("gsave");
                    fs.WriteLine(l3 + " execform");
                    fs.WriteLine("grestore");
                    fs.WriteLine(ftrans);
                }
                if (clogo == 4)
                {
                    fs.WriteLine("gsave");
                    fs.WriteLine(l1 + " execform");
                    fs.WriteLine("grestore");
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
                flogo = clogo;
                if (col == 3)
                {
                    col = 0; brs++;
                }
                if (brs == 21)
                { brs = 1; }
            }  //end for 2
        }

        private void dateTimePicker1_CloseUp(object sender, EventArgs e)
        {
            tglcyc = dateTimePicker1.Value.ToString("yyyyMMdd");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string sn = GetMotherBoardID();
            string serial = proc.Encrypt(sn,"-");
            MessageBox.Show("Serial Info :" + sn, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //MessageBox.Show("Serial Info :" + sn+"\n"+"SN:"+serial, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            string fout = Directory.GetCurrentDirectory() + "\\InfoSn.TXT";
            using (System.IO.StreamWriter fs = new System.IO.StreamWriter(fout, false))
            {
                fs.WriteLine(sn);
                fs.Close();
            }
            fout = Directory.GetCurrentDirectory() + "\\SerialNum.TXT";
            using (System.IO.StreamWriter fs = new System.IO.StreamWriter(fout, false))
            {
                fs.WriteLine(serial);
                fs.Close();
            }
            //string dec = proc.Decrypt(serial, "-");
            //MessageBox.Show("Result Info :" + dec, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
        public static string GetMotherBoardID()
        {
            String SerialNumber = "";
            ManagementObjectSearcher  mbs = new ManagementObjectSearcher("Select * from Win32_BaseBoard");
            foreach (ManagementObject mo in mbs.Get())
            {
                SerialNumber = mo["SerialNumber"].ToString().Trim();
            }
            return SerialNumber;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataTable listgrp = opendt("Select grp_card from MSTPROD Group BY grp_card");
            int jmlrec = listgrp.Rows.Count;
            comboBox1.Items.Clear();
            for (int gr = 0; gr < jmlrec; gr++)
            {
                comboBox1.Items.Add(listgrp.Rows[gr][0].ToString());
            }

            DataTable listprod = opendt("Select jns_card from MSTPROD ORDER BY grp_card,jns_card");
            jmlrec = listprod.Rows.Count;
            comboBox2.Items.Clear();
            for (int gr = 0; gr < jmlrec; gr++)
            {
                comboBox2.Items.Add(listprod.Rows[gr][0].ToString());
            }

        }
        void connectMDB()
        {
            kon = new OleDbConnection();
            kon.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Properties.Settings.Default.ConMDB.Trim() + ";User Id=admin;Password=;";
            try
            {
                kon.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private DataTable opendt(string strqry)
        {
            DataTable ft = new DataTable();
            connectMDB();
            OleDbCommand cmd = kon.CreateCommand();
            cmd.CommandText = strqry;
            OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection); // close conn after complete
            ft.Load(reader);
            return (ft);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.Items.Clear();
            comboBox2.Text = "";

            string fgrp = comboBox1.Text.Trim();
            DataTable listprod = opendt("Select jns_card from MSTPROD where grp_card='" + fgrp + "' ORDER BY jns_card");
            int jmlrec = listprod.Rows.Count;
            for (int gr = 0; gr < jmlrec; gr++)
            {
                comboBox2.Items.Add(listprod.Rows[gr][0].ToString());
                comboBox2.Text = listprod.Rows[0][0].ToString();
            }
            showtbl();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            showtbl();
            int jmlrec = dtmst.Rows.Count;
            if (jmlrec > 0)
            { txtInput.Enabled = true; }
        }
        void showtbl()
        {
            string fgrp = comboBox1.Text.Trim();
            string fjns = comboBox2.Text.Trim();
            dtmst = opendt("Select * from MSTPROD where grp_card='" + fgrp + "' and jns_card='" + fjns + "'");
        }
    }
}
