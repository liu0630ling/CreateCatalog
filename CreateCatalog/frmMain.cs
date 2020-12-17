using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;



namespace CreateCatalog
{
    public partial class frmMain : Form
    {
        public static string sProjectNo = System.Configuration.ConfigurationManager.AppSettings["ProjectNo"];
        string gProjectName;
        string gSubDiscipline;
        public static string sConstr = "Data Source=" + System.Configuration.ConfigurationManager.AppSettings["ProjectDataBaseServer"] + ";Initial Catalog=" + System.Configuration.ConfigurationManager.AppSettings["ProjectDataBaseName"] + ";User ID=PE-ADMIN;Password=PE-ADMIN";
        public static string sConstrBase = "Data Source=" + System.Configuration.ConfigurationManager.AppSettings["dbPEDMSServer"] + ";Initial Catalog=db_PEDMS;User ID=PE-ADMIN;Password=PE-ADMIN";
        string gRootDir;//项目的根目录
        string gsTemp;//临时文件目录

        int scu;
        string gUser;

        string gsVuser;//虚拟访问用户
        string gsVpwd;//虚拟访问密码

        bool gbUseProjCatalog;//是否使用项目自定义的目录格式
        int giCatalogDocNo;//目录中填写成达编号还是自定义编号
        int giCatalogTitle;//目录中填写成达主题还是自定义主题

        int iPageNum;
        string sDocTitle26;
        string sISOSoftWare;//抽出空视图的软件

        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            this.Hide();
            gUser = Environment.UserName;  //取当前用户名

            gsTemp = System.Environment.GetEnvironmentVariable("TEMP");
            if (gsTemp == null)
            {
                if (Directory.Exists(@"C:\Temp"))
                {
                    gsTemp = @"C:\Temp";
                }
                else
                {
                    MessageBox.Show(this, "无法找到系统临时目录，目录文件生成失败！请联系项目IT人员处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                    //Application.Exit();
                    return;
                }
            }

            //求虚拟用户名和密码
            DataSet dsV = new DataSet();
            dsV = PIMS_ClassLib.DataAccess.GetDataSet(sConstrBase, "select * from tb_access_user", out scu);
            if (scu == -1)
            {
                MessageBox.Show(this, "数据库连接失败！目录文件生成失败！请联系项目IT人员处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                //Application.Exit();
                this.Close();
                return;
            }
            else if (dsV.Tables[0].Rows.Count == 0)
            {
                MessageBox.Show(this, "数据库错误！目录文件生成失败！请联系项目IT人员处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);

                //Application.Exit();
                this.Close();
                return;
            }
            else
            {
                gsVuser = "chengda\\" + Crack(dsV.Tables[0].Rows[0][0].ToString());
                gsVpwd = Crack(dsV.Tables[0].Rows[0][1].ToString());
            }
            getProjectInfo();

            DataSet ds = new DataSet();
            ds = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select catalog_id,catalog_type,sub_discipline_no from tb_create_catalog where user_code = '" + gUser + "'", out scu);
            if (ds.Tables[0].Rows.Count > 0)
            {
                gSubDiscipline = ds.Tables[0].Rows[0][2].ToString();
                if (ds.Tables[0].Rows[0][1].ToString() == "ISO")
                {
                    getExcelISO(ds.Tables[0].Rows[0][0].ToString());
                    GC.Collect();
                    DataSet dsUpdateFormat = new DataSet();
                    dsUpdateFormat = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "update tb_design_doc set A4_qty = " + iPageNum.ToString() + ",doc_title = '" + sDocTitle26 + "' where design_doc_id = " + ds.Tables[0].Rows[0][0].ToString(), out scu);
                }
                else
                {
                    getExcel(ds.Tables[0].Rows[0][0].ToString());
                    GC.Collect();
                    DataSet dsUpdateFormat = new DataSet();
                    dsUpdateFormat = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "update tb_design_doc set source_format_suffix = '.xls',A4_qty = " + iPageNum.ToString() + " where design_doc_id = " + ds.Tables[0].Rows[0][0].ToString(), out scu);
                }
                ds = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "delete from tb_create_catalog where user_code = '" + gUser + "'", out scu);
                
            }
            this.Close();
            
        }
        private void getProjectInfo()
        {
            DataSet ds = new DataSet();
            ds = PIMS_ClassLib.DataAccess.GetDataSet(sConstrBase, "select project_name,use_proj_catalog,catalog_doc_no,catalog_title,iso_software from tb_V4_project_info where project_no = '" + sProjectNo + "'", out scu);
            gbUseProjCatalog = Convert.ToBoolean(ds.Tables[0].Rows[0]["use_proj_catalog"].ToString());
            giCatalogDocNo = (int)ds.Tables[0].Rows[0]["catalog_doc_no"];
            giCatalogTitle = (int)ds.Tables[0].Rows[0]["catalog_title"];
            sISOSoftWare = ds.Tables[0].Rows[0]["iso_software"].ToString();

            DataSet dsOther = PIMS_ClassLib.DataAccess.GetDataSet(sConstrBase, "select * from tb_other_project_name where project_no = '" + sProjectNo + "'", out scu);
            if (dsOther.Tables[0].Rows.Count == 0)
            {
                gProjectName = ds.Tables[0].Rows[0]["project_name"].ToString();
            }
            else
            {
                frmSelectProjectName fm = new frmSelectProjectName();
                fm.ShowDialog();
                gProjectName = frmSelectProjectName.sOtherProjectName;
            }
            //求根目录
            DataSet dsRootDir = new DataSet();
            dsRootDir = PIMS_ClassLib.DataAccess.GetDataSet(sConstrBase, "select doc_server,doc_root from tb_project_doc_path where project_no = '" + sProjectNo + "'", out scu);
            if (dsRootDir.Tables[0].Rows.Count != 1)
            {
                MessageBox.Show(this, "取项目文件根目录时出错！目录文件生成失败！请联系项目IT处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Application.Exit();
                return;
            }
            else if (dsRootDir.Tables[0].Rows[0][0] == DBNull.Value)
            {
                MessageBox.Show(this, "取项目文件根目录时出错！目录文件生成失败！请联系项目IT处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Application.Exit();
                return;
            }
            else if (dsRootDir.Tables[0].Rows[0][1] == DBNull.Value)
            {
                MessageBox.Show(this, "取项目文件根目录时出错！目录文件生成失败！请联系项目IT处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //Application.Exit();
                return;
            }
            else
            {
                gRootDir = @"\\" + dsRootDir.Tables[0].Rows[0][0].ToString() + @"\" + dsRootDir.Tables[0].Rows[0][1].ToString() + @"$\";
            }

        }
        private string Crack(string str) //解密
        {
            int i = str.Length;
            string CrackStr = "";
            for (int j = 0; j < i; j++)
            {
                CrackStr = CrackStr + (char)((int)str[j] - 18);
            }
            return CrackStr;
        }
        private void getExcelISO(string sCatalogId)
        {
            DataSet dsId = new DataSet();
            dsId = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select * from tb_design_doc where design_doc_id = " + sCatalogId, out scu);

            string sCdDocNo = dsId.Tables[0].Rows[0]["cd_doc_no"].ToString();
            string sCdDocNoWithNoNum = sCdDocNo.Substring(0, sCdDocNo.LastIndexOf('-') + 1); //不带流水号的文件编号
            string sCdDocNoNum = sCdDocNo.Substring(sCdDocNo.LastIndexOf('-') + 1, sCdDocNo.Length - sCdDocNo.LastIndexOf('-') - 1); //流水号

            string gPlant = dsId.Tables[0].Rows[0]["plant_no"].ToString();
            string gSelectedMainUnitNo = dsId.Tables[0].Rows[0]["main_unit_no"].ToString();
            string gSelectedSubUnitNo = dsId.Tables[0].Rows[0]["sub_unit_no"].ToString();
            string strVer = dsId.Tables[0].Rows[0]["verison"].ToString();
            string strVerDis = dsId.Tables[0].Rows[0]["verison_discription"].ToString();
            string sDesigner26 = dsId.Tables[0].Rows[0]["designer"].ToString();
            string sChecker26 = dsId.Tables[0].Rows[0]["checker"].ToString();
            int iStage = (int)dsId.Tables[0].Rows[0]["stage_id"];
            string gSelectedDesignDate;
            string gSelectedCheckDate;
            if (dsId.Tables[0].Rows[0]["IFC_refer"] == DBNull.Value)
            {
                gSelectedDesignDate = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString();
            }
            else
            {
                gSelectedDesignDate = Convert.ToDateTime(dsId.Tables[0].Rows[0]["IFC_refer"].ToString()).Year.ToString() + "." + Convert.ToDateTime(dsId.Tables[0].Rows[0]["IFC_refer"].ToString()).Month.ToString() + "." + Convert.ToDateTime(dsId.Tables[0].Rows[0]["IFC_refer"].ToString()).Day.ToString();
            }
            if (dsId.Tables[0].Rows[0]["check_refer"] == DBNull.Value)
            {
                gSelectedCheckDate = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString();
            }
            else
            {
                gSelectedCheckDate = Convert.ToDateTime(dsId.Tables[0].Rows[0]["check_refer"].ToString()).Year.ToString() + "." + Convert.ToDateTime(dsId.Tables[0].Rows[0]["check_refer"].ToString()).Month.ToString() + "." + Convert.ToDateTime(dsId.Tables[0].Rows[0]["check_refer"].ToString()).Day.ToString();
            }

            //求此目录是新增还是升版
            string sDocRemark;
            DataSet dsRemark = new DataSet();
            dsRemark = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select * from tb_design_doc where cd_doc_no = '" + sCdDocNo + "'", out scu);
            if (dsRemark.Tables[0].Rows.Count == 1)
            {
                sDocRemark = "新增";
            }
            else
            {
                sDocRemark = "已升版";
            }

            int iDocCount = 0; //数据库总共有多少个文件
            int iNeedRow = 0; //excel中总共需要多少行来显示
            DataSet dsDocCount = new DataSet();
            dsDocCount = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select count(relation_id) from tbr_catalog_relation_iso where catalog_id = " + sCatalogId, out scu);
            iDocCount = Convert.ToInt32(dsDocCount.Tables[0].Rows[0][0].ToString());
            iNeedRow = iDocCount + 1 + 5;
            iPageNum = iNeedRow / 43 + 2;//这个是需要的页数

            int iFileNum = 1; //空视图序号
            int iCurrentRow = 3;//当前行
            int iA3 = 0;//记录A3的数量

            //打开工作薄
            Microsoft.Office.Interop.Excel.Application app;
            Workbooks wbs;
            Workbook wb;
            Worksheet ws; //目录页
            Worksheet ws1; //首页
            Range raMerge;


            app = new Microsoft.Office.Interop.Excel.Application();
            app.DisplayAlerts = false;
            wbs = app.Workbooks;

            //将服务器上的模板拷贝至temp目录
            try
            {
                ImpersonationHelper helper = new ImpersonationHelper();
                if (helper.Impersonate(gsVuser, gsVpwd))
                {

                    //拷贝模板
                    File.Copy(gRootDir + sProjectNo + @"\目录模板\model_iso_catalog.xls", gsTemp + "\\model_iso_catalog.xls", true);

                    helper.EndImpersonate();//结束身份模拟。
                }
                else
                {
                    //如果模拟失败在这里处理。
                    MessageBox.Show(this, "拷贝项目自定义模板失败！请通知项目IT处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch
            {
                MessageBox.Show(this, "拷贝项目自定义模板失败！可能是后台权限问题，请通知项目IT处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            wb = wbs.Open(gsTemp + "\\model_iso_catalog.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            ws1 = (Worksheet)wb.Worksheets["首页"];

            //求目录文件信息
            DataSet dsDoc = new DataSet();
            dsDoc = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select * from tb_design_doc where design_doc_id = " + sCatalogId, out scu);



            //求业主名称
            DataSet dsOwner = new DataSet();
            dsOwner = PIMS_ClassLib.DataAccess.GetDataSet(sConstrBase, "select owner_name from tb_V4_project_info where project_no = '" + sProjectNo + "'", out scu);

            //写首页信息
            ws1.Cells[8, 1] = dsOwner.Tables[0].Rows[0][0].ToString();
            ws1.Cells[10, 1] = gProjectName;
            ws1.Cells[12, 1] = Message.PlantName(gPlant);
            ws1.Cells[34, 1] = strVer;
            ws1.Cells[34, 3] = strVerDis;
            ws1.Cells[34, 11] = sDesigner26.Substring(sDesigner26.IndexOf(' ') + 1, sDesigner26.Length - sDesigner26.IndexOf(' ') - 1);
            ws1.Cells[34, 16] = gSelectedDesignDate;
            ws1.Cells[34, 18] = sChecker26.Substring(sChecker26.IndexOf(' ') + 1, sChecker26.Length - sChecker26.IndexOf(' ') - 1);
            ws1.Cells[34, 23] = gSelectedCheckDate;
            ws1.Cells[37, 11] = Message.MainUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString()) + (Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString()).Trim().IndexOf("总体") >= 0 || Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString()).Trim().IndexOf("General") >= 0 ? "" : Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString()));
            ws1.Cells[37, 25] = gProjectName;
            //ws1.Cells[39, 25] = Message.PlantName(gPlant) + "/" + Message.MainUnitName(gPlant, gSelectedMainUnitNo) + "/" + Message.SubUnitName(gPlant, gSelectedMainUnitNo, gSelectedSubUnitNo); ;
            ws1.Cells[39, 25] = Message.PlantName(gPlant) + "/" + Message.MainUnitName(gPlant, gSelectedMainUnitNo);

            ws1.Cells[41, 25] = sCdDocNo;
            ws1.Cells[43, 14] = sProjectNo;
            ws1.Cells[43, 18] = Message.SubDisciplineName(gSubDiscipline);
            ws1.Cells[43, 24] = Message.StageName(iStage.ToString());
            ws1.Cells[43, 27] = "第1张 共" + iPageNum.ToString() + "张";
            ws1.Cells[44, 27] = "SHEET 1 OF " + iPageNum.ToString();

            sDocTitle26 = Message.MainUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString()) + (Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString()).Trim().IndexOf("总体") >= 0 || Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString()).Trim().IndexOf("General") >= 0 ? "" : Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString())) + "管道空视图图纸目录";

            //写主页信息
            ws = (Worksheet)wb.Worksheets["目录"];

            //第一行写目录本身
            ws.Cells[iCurrentRow, 1] = iFileNum.ToString();
            ws.Cells[iCurrentRow, 2] = Message.MainUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString()) + (Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString()).Trim() == "总体" || Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString()).Trim() == "General" || Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString()).Trim() == "总体 General" ? "" : Message.SubUnitName(gPlant, dsDoc.Tables[0].Rows[0]["main_unit_no"].ToString(), dsDoc.Tables[0].Rows[0]["sub_unit_no"].ToString())) + (giCatalogTitle == 1 ? "Isometric Drawing List" : "管道空视图目录");
            if (sISOSoftWare == "SP3D" || sISOSoftWare == "PDMS")
            {
                raMerge = ws.get_Range("D" + iCurrentRow.ToString(), "G" + iCurrentRow.ToString());
                raMerge.Merge();
                raMerge.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                ws.Cells[iCurrentRow, 4] = sCdDocNo;
            }
            else
            {
                ws.Cells[iCurrentRow, 4] = sCdDocNoWithNoNum;
                ws.Cells[iCurrentRow, 5] = sCdDocNoNum;
            }
            ws.Cells[iCurrentRow, 8] = "A4";
            ws.Cells[iCurrentRow, 9] = iPageNum.ToString();
            ws.Cells[iCurrentRow, 10] = strVer;
            ws.Cells[iCurrentRow, 11] = sDocRemark;
            iCurrentRow++;
            iFileNum++;

            //第二行检查是否有封面文件
            DataSet dsCover = new DataSet();
            dsCover = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select * from tbr_catalog_relation_iso where iso_name like 'cover%' and catalog_id = " + sCatalogId, out scu);
            if (dsCover.Tables[0].Rows.Count > 0)
            {
                ws.Cells[iCurrentRow, 1] = iFileNum.ToString();
                ws.Cells[iCurrentRow, 2] = dsCover.Tables[0].Rows[0]["iso_name"].ToString().Substring(dsCover.Tables[0].Rows[0]["iso_name"].ToString().IndexOf(':') + 1, dsCover.Tables[0].Rows[0]["iso_name"].ToString().Length - dsCover.Tables[0].Rows[0]["iso_name"].ToString().IndexOf(':') - 1);
                ws.Cells[iCurrentRow, 4] = sCdDocNoWithNoNum;
                ws.Cells[iCurrentRow, 5] = "Cover";
                ws.Cells[iCurrentRow, 8] = "A3";
                ws.Cells[iCurrentRow, 9] = dsCover.Tables[0].Rows[0]["qty"].ToString();
                ws.Cells[iCurrentRow, 10] = dsCover.Tables[0].Rows[0]["verison"].ToString();
                ws.Cells[iCurrentRow, 11] = dsCover.Tables[0].Rows[0]["remark"].ToString();
                iCurrentRow++;
                iFileNum++;
                if (dsCover.Tables[0].Rows[0]["remark"].ToString().Trim() != "作废") iA3 = iA3 + Convert.ToInt32(dsCover.Tables[0].Rows[0]["qty"].ToString());
            }

            //填写空视图文件
            DataSet dsISO = new DataSet();
            dsISO = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select * from tbr_catalog_relation_iso where iso_name not like 'cover%' and catalog_id = " + sCatalogId + " order by iso_name", out scu);
            foreach (DataRow row in dsISO.Tables[0].Rows)
            {
                ws.Cells[iCurrentRow, 1] = iFileNum.ToString();
                if (sISOSoftWare == "SP3D" || sISOSoftWare == "PDMS")
                {
                    raMerge = ws.get_Range("B" + iCurrentRow.ToString(), "C" + iCurrentRow.ToString());
                    raMerge.Merge();
                    raMerge.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ws.Cells[iCurrentRow, 2] = row["iso_title"].ToString();
                    raMerge = ws.get_Range("D" + iCurrentRow.ToString(), "G" + iCurrentRow.ToString());
                    raMerge.Merge();
                    raMerge.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    ws.Cells[iCurrentRow, 4] = row["iso_name"].ToString();
                }
                else
                {
                    ws.Cells[iCurrentRow, 2] = row["iso_name"].ToString().Substring(0, row["iso_name"].ToString().Length - 2);
                    ws.Cells[iCurrentRow, 3] = giCatalogTitle == 1 ? "Isometric" : "空视图";
                    ws.Cells[iCurrentRow, 4] = sCdDocNoWithNoNum;
                    ws.Cells[iCurrentRow, 5] = row["iso_name"].ToString().Substring(0, row["iso_name"].ToString().Length - 2);
                    ws.Cells[iCurrentRow, 6] = "-";
                    ws.Cells[iCurrentRow, 7] = row["iso_name"].ToString().Substring(row["iso_name"].ToString().Length - 2, 2);
                }
                ws.Cells[iCurrentRow, 8] = "A3";
                ws.Cells[iCurrentRow, 9] = row["qty"].ToString();
                ws.Cells[iCurrentRow, 10] = row["verison"].ToString();
                ws.Cells[iCurrentRow, 11] = row["remark"].ToString();
                iCurrentRow++;
                iFileNum++;
                //if (row["remark"].ToString().Trim() != "作废") iA3++;
                if (row["remark"].ToString().Trim() != "作废") iA3 = iA3 + Convert.ToInt32(row["qty"].ToString());
            }

            //填写总计信息
            iCurrentRow = iCurrentRow + 2;
            if (giCatalogTitle == 1)
            {
                ws.Cells[iCurrentRow, 5] = "sheet:";
                ws.Cells[iCurrentRow, 8] = "A4";
                ws.Cells[iCurrentRow, 9] = iPageNum.ToString();
                //ws.Cells[iCurrentRow, 10] = "张";
                iCurrentRow++;
                ws.Cells[iCurrentRow, 8] = "A3";
                ws.Cells[iCurrentRow, 9] = iA3.ToString();
                //ws.Cells[iCurrentRow, 10] = "张";
                iCurrentRow++;
                ws.Cells[iCurrentRow, 5] = "Equivalent to:";
                ws.Cells[iCurrentRow, 8] = "A1";
                ws.Cells[iCurrentRow, 9] = iPageNum / 8.0 + iA3 / 4.0;
                //ws.Cells[iCurrentRow, 10] = "张";
            }
            else
            {
                ws.Cells[iCurrentRow, 7] = "共";
                ws.Cells[iCurrentRow, 8] = "A4";
                ws.Cells[iCurrentRow, 9] = iPageNum.ToString();
                ws.Cells[iCurrentRow, 10] = "张";
                iCurrentRow++;
                ws.Cells[iCurrentRow, 8] = "A3";
                ws.Cells[iCurrentRow, 9] = iA3.ToString();
                ws.Cells[iCurrentRow, 10] = "张";
                iCurrentRow++;
                ws.Cells[iCurrentRow, 6] = "折合";
                ws.Cells[iCurrentRow, 8] = "A1";
                ws.Cells[iCurrentRow, 9] = iPageNum / 8.0 + iA3 / 4.0;
                ws.Cells[iCurrentRow, 10] = "张";
            }

            //删除多余的页
            int iDelBegin = 2 + 43 * (iPageNum - 1) + 1;
            if (iDelBegin < 10000)
            {
                Range ra = ws.get_Range(ws.Cells[iDelBegin, 1], ws.Cells[10000, 11]);
                ra.Delete(XlDeleteShiftDirection.xlShiftUp);
            }

            //将数量计入tb_paper_putin_qty_record
            DataSet dsExist = new DataSet();
            dsExist = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select * from tb_paper_putin_qty_record where catalog_id = " + sCatalogId, out scu);
            if (dsExist.Tables[0].Rows.Count > 0)
            {
                DataSet dsUpdate = new DataSet();
                dsUpdate = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "update tb_paper_putin_qty_record set A0_qty = 0,A1_qty = 0,A2_qty = 0,A3_qty = " + iA3.ToString() + ",A4_qty = " + iPageNum.ToString() + " where catalog_id = " + sCatalogId, out scu);
            }
            else
            {
                DataSet dsInsert = new DataSet();
                dsInsert = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "insert into tb_paper_putin_qty_record(catalog_id,A0_qty,A1_qty,A2_qty,A3_qty,A4_qty)values(" + sCatalogId + ",0,0,0," + iA3.ToString() + "," + iPageNum.ToString() + ")", out scu);
            }

            try
            {

                wb.SaveAs(gsTemp + "\\" + iStage.ToString() + " " + sCdDocNo + "-Rev" + strVer + ".xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            }
            catch
            {
                MessageBox.Show(this, "可能你要替换的文件正在使用中，无法进行替换，本次导出操作失败！", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                wb.Close(false, false, Type.Missing);
                wbs.Close();
                ws = null;
                wb = null;
                wbs = null;
                app.Quit();
                app = null;

                return;
            }


            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            ws = null;
            wb = null;
            wbs = null;
            app.Quit();
            app = null;
        }
        private void getExcelNew(string sCatalogId)
        {

        }
        private void getExcel(string sCatalogId)
        {
            DataSet dsId = new DataSet();
            dsId = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select * from tb_design_doc where design_doc_id = " + sCatalogId, out scu);

            string sCdDocNo = dsId.Tables[0].Rows[0]["cd_doc_no"].ToString();
            int iStage = (int)dsId.Tables[0].Rows[0]["stage_id"];
            string gPlant = dsId.Tables[0].Rows[0]["plant_no"].ToString();
            string gSelectedMainUnitNo = dsId.Tables[0].Rows[0]["main_unit_no"].ToString();
            string gSelectedSubUnitNo = dsId.Tables[0].Rows[0]["sub_unit_no"].ToString();
            string gSelectedVerison = dsId.Tables[0].Rows[0]["verison"].ToString();
            string gSelectedVerisonDis = dsId.Tables[0].Rows[0]["verison_discription"].ToString();
            string gSelectedDesigner = dsId.Tables[0].Rows[0]["designer"].ToString();
            string gSelectedChecker = dsId.Tables[0].Rows[0]["checker"].ToString();
            string gSelectedDocTitle = dsId.Tables[0].Rows[0]["doc_title"].ToString();
            string gSelectedUserDocNo = dsId.Tables[0].Rows[0]["user_doc_no"].ToString();
            string gSelectedUserTitle = dsId.Tables[0].Rows[0]["user_title"].ToString();
            string gSelectedDesignDate;
            string gSelectedCheckDate;
            if (dsId.Tables[0].Rows[0]["IFC_refer"] == DBNull.Value)
            {
                gSelectedDesignDate = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString();
            }
            else
            {
                gSelectedDesignDate = Convert.ToDateTime(dsId.Tables[0].Rows[0]["IFC_refer"].ToString()).Year.ToString() + "." + Convert.ToDateTime(dsId.Tables[0].Rows[0]["IFC_refer"].ToString()).Month.ToString() + "." + Convert.ToDateTime(dsId.Tables[0].Rows[0]["IFC_refer"].ToString()).Day.ToString();
            }
            if (dsId.Tables[0].Rows[0]["check_refer"] == DBNull.Value)
            {
                gSelectedCheckDate = DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString();
            }
            else
            {
                gSelectedCheckDate = Convert.ToDateTime(dsId.Tables[0].Rows[0]["check_refer"].ToString()).Year.ToString() + "." + Convert.ToDateTime(dsId.Tables[0].Rows[0]["check_refer"].ToString()).Month.ToString() + "." + Convert.ToDateTime(dsId.Tables[0].Rows[0]["check_refer"].ToString()).Day.ToString();
            }

            //先求是不是第一个版本
            bool bFirst = true;
            DataSet dsFirst = new DataSet();
            dsFirst = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select design_doc_id from tb_design_doc where stage_id = " + iStage.ToString() + " and cd_doc_no = '" + sCdDocNo + "'", out scu);
            if (dsFirst.Tables[0].Rows.Count > 1)
            {
                bFirst = false;
            }

            int iDocCount = 0;
            DataSet dsDocCount = new DataSet();
            dsDocCount = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select count(design_doc_id) from tbr_catalog_relation where catalog_id = " + sCatalogId, out scu);
            iDocCount = Convert.ToInt32(dsDocCount.Tables[0].Rows[0][0].ToString());
            //计算需要多少页

            if (iDocCount % 18 > 15)
            {
                iPageNum = iDocCount / 18 + 2;
            }
            else
            {
                iPageNum = iDocCount / 18 + 1;
            }
            //int iFileNum = 1;

            //打开工作薄
            Microsoft.Office.Interop.Excel.Application app;
            Workbooks wbs;
            Workbook wb;
            Worksheet ws;



            app = new Microsoft.Office.Interop.Excel.Application();
            app.DisplayAlerts = false;
            wbs = app.Workbooks;

                       
            //将服务器上的模板拷贝至temp目录
            try
            {
                ImpersonationHelper helper = new ImpersonationHelper();
                if (helper.Impersonate(gsVuser, gsVpwd))
                {

                    //拷贝模板
                    File.Copy(gRootDir + sProjectNo + @"\目录模板\model_catalog.xls", gsTemp + "\\model_catalog.xls", true);

                    helper.EndImpersonate();//结束身份模拟。
                }
                else
                {
                    //如果模拟失败在这里处理。
                    MessageBox.Show(this, "拷贝项目自定义模板失败！请通知项目IT处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch
            {
                MessageBox.Show(this, "拷贝项目自定义模板失败！可能是后台权限问题，请通知项目IT处理。", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            wb = wbs.Open(gsTemp + "\\model_catalog.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            
            ws = (Worksheet)wb.Worksheets["图纸目录"];
            string sFXMC = "";
            //if (Message.SubUnitName(gPlant, gSelectedMainUnitNo, gSelectedSubUnitNo).Trim() == "总体" || Message.SubUnitName(gPlant, gSelectedMainUnitNo, gSelectedSubUnitNo).Trim().ToLower() == "general")
            //{
            //    if (Message.MainUnitName(gPlant, gSelectedMainUnitNo).Trim() == "总体" || Message.MainUnitName(gPlant, gSelectedMainUnitNo).Trim().ToLower() == "general")
            //    {
            //        sFXMC = Message.PlantName(gPlant);
            //    }
            //    else
            //    {
            //        sFXMC = Message.MainUnitName(gPlant, gSelectedMainUnitNo);
            //    }
            //}
            //else
            //{
            //    sFXMC = Message.SubUnitName(gPlant, gSelectedMainUnitNo, gSelectedSubUnitNo);
            //}
            //sFXMC = Message.PlantName(gPlant) + "/" + Message.MainUnitName(gPlant, gSelectedMainUnitNo) + "/" + Message.SubUnitName(gPlant, gSelectedMainUnitNo, gSelectedSubUnitNo);
            sFXMC = Message.PlantName(gPlant) + "/" + Message.MainUnitName(gPlant, gSelectedMainUnitNo);
            //先填写其他信息
            for (int i = 0; i <= iPageNum - 1; i++)
            {
                ws.Cells[i * 30 + 25, 2] = gSelectedVerison;//版本号
                ws.Cells[i * 30 + 25, 4] = gSelectedVerisonDis;//版本说明
                ws.Cells[i * 30 + 25, 14] = gSelectedDesigner.Substring(gSelectedDesigner.IndexOf(' ') + 1, gSelectedDesigner.Length - gSelectedDesigner.IndexOf(' ') - 1);//设计人
                ws.Cells[i * 30 + 25, 21] = gSelectedChecker.Substring(gSelectedChecker.IndexOf(' ') + 1, gSelectedChecker.Length - gSelectedChecker.IndexOf(' ') - 1); ;//校核人
                ws.Cells[i * 30 + 25, 19] = gSelectedDesignDate;
                ws.Cells[i * 30 + 25, 26] = gSelectedCheckDate;
                //if (cboAuditor.Enabled)
                //{
                //    ws.Cells[i * 30 + 45, 28] = cboAuditor.Text.Substring(cboAuditor.Text.IndexOf(' ') + 1, cboAuditor.Text.Length - cboAuditor.Text.IndexOf(' ') - 1);//审核人
                //}
                //ws.Cells[i * 30 + 27, 11] = gSelectedDocTitle;//主题
                ws.Cells[i * 30 + 27, 11] = Message.SubUnitName(gPlant, gSelectedMainUnitNo, gSelectedSubUnitNo);//子项名称
                ws.Cells[i * 30 + 28, 11] = gSelectedDocTitle;//主题
                ws.Cells[i * 30 + 27, 27] = gProjectName;//项目名称
                ws.Cells[i * 30 + 28, 27] = sFXMC;//装置名称
                if (giCatalogDocNo == 0)
                {
                    ws.Cells[i * 30 + 29, 27] = sCdDocNo;//图号
                }
                else if (giCatalogDocNo == 1)
                {
                    ws.Cells[i * 30 + 29, 27] = gSelectedUserDocNo;//图号
                }
                else if (giCatalogDocNo == 2)
                {
                    ws.Cells[i * 30 + 29, 27] = sCdDocNo + " " + gSelectedUserDocNo;
                }
                ws.Cells[i * 30 + 30, 14] = sProjectNo;//项目代码
                ws.Cells[i * 30 + 30, 20] = Message.SubDisciplineName(gSubDiscipline);//小专业
                ws.Cells[i * 30 + 30, 26] = Message.StageName(iStage.ToString());//设计阶段
                ws.Cells[i * 30 + 30, 32] = "共 " + iPageNum + " 页";


            }
            //填图纸内容
            int iCurrentRow = 5;
            int iCurrentPage = 0;//当前页

            int iNatureQty = 0;//自然张数
            double iToA1Qty = 0;//折合A1

            string sQty = "";//图幅和张数变量

            //先填写自己
            iNatureQty = iNatureQty + iPageNum;
            iToA1Qty = iToA1Qty + Convert.ToDouble(iPageNum) / 8;
            ws.Cells[iCurrentRow + iCurrentPage * 30, 2] = "1";
            //iFileNum++;
            if (giCatalogDocNo == 0)
            {
                ws.Cells[iCurrentRow + iCurrentPage * 30, 4] = sCdDocNo;
            }
            else if (giCatalogDocNo == 1)
            {
                ws.Cells[iCurrentRow + iCurrentPage * 30, 4] = gSelectedUserDocNo;
            }
            else if (giCatalogDocNo == 2)
            {
                ws.Cells[iCurrentRow + iCurrentPage * 30, 4] = sCdDocNo + " " + gSelectedUserDocNo;
            }
            ws.Cells[iCurrentRow + iCurrentPage * 30, 13] = gSelectedVerison;
            if (giCatalogTitle == 0)
            {
                ws.Cells[iCurrentRow + iCurrentPage * 30, 16] = gSelectedDocTitle;
            }
            else if (giCatalogTitle == 1)
            {
                ws.Cells[iCurrentRow + iCurrentPage * 30, 16] = gSelectedUserTitle;
            }
            else if (giCatalogTitle == 2)
            {
                ws.Cells[iCurrentRow + iCurrentPage * 30, 16] = gSelectedDocTitle + " " + gSelectedUserTitle;
            }
            ws.Cells[iCurrentRow + iCurrentPage * 30, 25] = "A4：" + iPageNum.ToString();
            iCurrentRow++;

            DataSet dsDoc = new DataSet();
            dsDoc = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select b.cd_doc_no,b.user_doc_no,b.doc_title,b.user_title,b.verison,b.A0_qty,b.A1_qty,b.A2_qty,b.A3_qty,b.A4_qty,a.remark,a.no from tbr_catalog_relation a left outer join tb_design_doc b on a.design_doc_id = b.design_doc_id where a.catalog_id = " + sCatalogId + " and a.remark <> '作废' order by a.no", out scu);

            double dA0 = 0;
            double dA1 = 0;
            double dA2 = 0;
            double dA3 = 0;
            double dA4 = 0;

            foreach (DataRow row in dsDoc.Tables[0].Rows)
            {

                sQty = "";
                if (row[5].ToString() != "0")
                {
                    dA0 = dA0 + Convert.ToDouble(row[5].ToString());
                    sQty = sQty + "A0：" + row[5].ToString() + " ";
                    iNatureQty = iNatureQty + (int)Convert.ToDouble(row[5].ToString());
                    iToA1Qty = iToA1Qty + Convert.ToDouble(row[5].ToString()) * 2;
                }
                if (row[6].ToString() != "0")
                {
                    dA1 = dA1 + Convert.ToDouble(row[6].ToString());
                    sQty = sQty + "A1：" + row[6].ToString() + " ";
                    iNatureQty = iNatureQty + (int)Convert.ToDouble(row[6].ToString());
                    iToA1Qty = iToA1Qty + Convert.ToDouble(row[6].ToString());
                }
                if (row[7].ToString() != "0")
                {
                    dA2 = dA2 + Convert.ToDouble(row[7].ToString());
                    sQty = sQty + "A2：" + row[7].ToString() + " ";
                    iNatureQty = iNatureQty + (int)Convert.ToDouble(row[7].ToString());
                    iToA1Qty = iToA1Qty + Convert.ToDouble(row[7].ToString()) / 2;
                }
                if (row[8].ToString() != "0")
                {
                    dA3 = dA3 + Convert.ToDouble(row[8].ToString());
                    sQty = sQty + "A3：" + row[8].ToString() + " ";
                    iNatureQty = iNatureQty + (int)Convert.ToDouble(row[8].ToString());
                    iToA1Qty = iToA1Qty + Convert.ToDouble(row[8].ToString()) / 4;
                }
                if (row[9].ToString() != "0")
                {
                    dA4 = dA4 + Convert.ToDouble(row[9].ToString());
                    sQty = sQty + "A4：" + row[9].ToString() + " ";
                    iNatureQty = iNatureQty + (int)Convert.ToDouble(row[9].ToString());
                    iToA1Qty = iToA1Qty + Convert.ToDouble(row[9].ToString()) / 8;
                }
                ws.Cells[iCurrentRow + iCurrentPage * 30, 2] = row[11].ToString();
                //iFileNum++;
                if (giCatalogDocNo == 0)
                {
                    ws.Cells[iCurrentRow + iCurrentPage * 30, 4] = row[0].ToString();
                }
                else if (giCatalogDocNo == 1)
                {
                    ws.Cells[iCurrentRow + iCurrentPage * 30, 4] = row[1].ToString();
                }
                else if (giCatalogDocNo == 2)
                {
                    ws.Cells[iCurrentRow + iCurrentPage * 30, 4] = row[0].ToString() + " " + row[1].ToString();
                }
                ws.Cells[iCurrentRow + iCurrentPage * 30, 13] = row[4].ToString();
                if (giCatalogTitle == 0)
                {
                    ws.Cells[iCurrentRow + iCurrentPage * 30, 16] = row[2].ToString();
                }
                else if (giCatalogTitle == 1)
                {
                    ws.Cells[iCurrentRow + iCurrentPage * 30, 16] = row[3].ToString();
                }
                else if (giCatalogTitle == 2)
                {
                    ws.Cells[iCurrentRow + iCurrentPage * 30, 16] = row[2].ToString() + " " + row[3].ToString();
                }
                ws.Cells[iCurrentRow + iCurrentPage * 30, 25] = sQty.Substring(0, sQty.Length - 1);
                //ws.Cells[iCurrentRow + iCurrentPage * 30, 31] = row[10].ToString();
                if (!bFirst && row[10].ToString() != "未升版")
                {
                    ws.Cells[iCurrentRow + iCurrentPage * 30, 31] = row[10].ToString();
                }
                iCurrentRow++;
                if (iCurrentRow > 22)
                {
                    iCurrentPage++;
                    iCurrentRow = 5;
                }

            }

            dA4 = dA4 + iPageNum;

            //将数量计入tb_paper_putin_qty_record
            DataSet dsExist = new DataSet();
            dsExist = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "select * from tb_paper_putin_qty_record where catalog_id = " + sCatalogId, out scu);
            if (dsExist.Tables[0].Rows.Count > 0)
            {
                DataSet dsUpdate = new DataSet();
                dsUpdate = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "update tb_paper_putin_qty_record set A0_qty = " + dA0.ToString() + ",A1_qty = " + dA1.ToString() + ",A2_qty = " + dA2.ToString() + ",A3_qty = " + dA3.ToString() + ",A4_qty = " + dA4.ToString() + " where catalog_id = " + sCatalogId, out scu);
            }
            else
            {
                DataSet dsInsert = new DataSet();
                dsInsert = PIMS_ClassLib.DataAccess.GetDataSet(sConstr, "insert into tb_paper_putin_qty_record(catalog_id,A0_qty,A1_qty,A2_qty,A3_qty,A4_qty)values(" + sCatalogId + "," + dA0.ToString() + "," + dA1.ToString() + "," + dA2.ToString() + "," + dA3.ToString() + "," + dA4.ToString() + ")", out scu);
            }

            //添加自然张数和折合A1
            if (iCurrentRow + iCurrentPage * 30 == 21 || iCurrentRow + iCurrentPage * 30 == 22)
            {
                ws.Cells[35, 16] = "自然张数：";
                ws.Cells[35, 25] = iNatureQty.ToString();
                ws.Cells[36, 16] = "折合A1：";
                ws.Cells[36, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 51 || iCurrentRow + iCurrentPage * 30 == 52)
            {
                ws.Cells[65, 16] = "自然张数：";
                ws.Cells[65, 25] = iNatureQty.ToString();
                ws.Cells[66, 16] = "折合A1：";
                ws.Cells[66, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 81 || iCurrentRow + iCurrentPage * 30 == 82)
            {
                ws.Cells[95, 16] = "自然张数：";
                ws.Cells[95, 25] = iNatureQty.ToString();
                ws.Cells[96, 16] = "折合A1：";
                ws.Cells[96, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 111 || iCurrentRow + iCurrentPage * 30 == 112)
            {
                ws.Cells[125, 16] = "自然张数：";
                ws.Cells[125, 25] = iNatureQty.ToString();
                ws.Cells[126, 16] = "折合A1：";
                ws.Cells[126, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 141 || iCurrentRow + iCurrentPage * 30 == 142)
            {
                ws.Cells[155, 16] = "自然张数：";
                ws.Cells[155, 25] = iNatureQty.ToString();
                ws.Cells[156, 16] = "折合A1：";
                ws.Cells[156, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 171 || iCurrentRow + iCurrentPage * 30 == 172)
            {
                ws.Cells[185, 16] = "自然张数：";
                ws.Cells[185, 25] = iNatureQty.ToString();
                ws.Cells[186, 16] = "折合A1：";
                ws.Cells[186, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 201 || iCurrentRow + iCurrentPage * 30 == 202)
            {
                ws.Cells[215, 16] = "自然张数：";
                ws.Cells[215, 25] = iNatureQty.ToString();
                ws.Cells[216, 16] = "折合A1：";
                ws.Cells[216, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 231 || iCurrentRow + iCurrentPage * 30 == 232)
            {
                ws.Cells[245, 16] = "自然张数：";
                ws.Cells[245, 25] = iNatureQty.ToString();
                ws.Cells[246, 16] = "折合A1：";
                ws.Cells[246, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 261 || iCurrentRow + iCurrentPage * 30 == 262)
            {
                ws.Cells[275, 16] = "自然张数：";
                ws.Cells[275, 25] = iNatureQty.ToString();
                ws.Cells[276, 16] = "折合A1：";
                ws.Cells[276, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 291 || iCurrentRow + iCurrentPage * 30 == 292)
            {
                ws.Cells[305, 16] = "自然张数：";
                ws.Cells[305, 25] = iNatureQty.ToString();
                ws.Cells[306, 16] = "折合A1：";
                ws.Cells[306, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 321 || iCurrentRow + iCurrentPage * 30 == 322)
            {
                ws.Cells[335, 16] = "自然张数：";
                ws.Cells[335, 25] = iNatureQty.ToString();
                ws.Cells[336, 16] = "折合A1：";
                ws.Cells[336, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 351 || iCurrentRow + iCurrentPage * 30 == 352)
            {
                ws.Cells[365, 16] = "自然张数：";
                ws.Cells[365, 25] = iNatureQty.ToString();
                ws.Cells[366, 16] = "折合A1：";
                ws.Cells[366, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 381 || iCurrentRow + iCurrentPage * 30 == 382)
            {
                ws.Cells[395, 16] = "自然张数：";
                ws.Cells[395, 25] = iNatureQty.ToString();
                ws.Cells[396, 16] = "折合A1：";
                ws.Cells[396, 25] = iToA1Qty.ToString();
            }
            else if (iCurrentRow + iCurrentPage * 30 == 411 || iCurrentRow + iCurrentPage * 30 == 412)
            {
                ws.Cells[425, 16] = "自然张数：";
                ws.Cells[425, 25] = iNatureQty.ToString();
                ws.Cells[426, 16] = "折合A1：";
                ws.Cells[426, 25] = iToA1Qty.ToString();
            }
            else
            {
                ws.Cells[iCurrentRow + iCurrentPage * 30 + 1, 16] = "自然张数：";
                ws.Cells[iCurrentRow + iCurrentPage * 30 + 1, 25] = iNatureQty.ToString();
                ws.Cells[iCurrentRow + iCurrentPage * 30 + 2, 16] = "折合A1：";
                ws.Cells[iCurrentRow + iCurrentPage * 30 + 2, 25] = iToA1Qty.ToString();
            }

            //删除多余页
            if (iPageNum != 15)
            {

                Range ra = ws.get_Range(ws.Cells[iPageNum * 30 + 1, 1], ws.Cells[450, 34]);
                ra.Delete(XlDeleteShiftDirection.xlShiftUp);

                //for (int i = iPageNum + 1; i <= 8; i++)
                //{
                //    ws.Shapes.Item(i).Delete();
                //}

            }


            try
            {

                wb.SaveAs(gsTemp + "\\" + iStage.ToString() + " " + sCdDocNo + "-Rev" + gSelectedVerison + ".xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            }
            catch
            {
                MessageBox.Show(this, "可能你要替换的文件正在使用中，无法进行替换，本次导出操作失败！", "PEDMS 错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                wb.Close(false, false, Type.Missing);
                wbs.Close();
                ws = null;
                wb = null;
                wbs = null;
                app.Quit();
                app = null;

                return;
            }


            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            ws = null;
            wb = null;
            wbs = null;
            app.Quit();
            app = null;
        }
    }
}
