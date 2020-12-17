using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using PedmsDataAccess;

namespace CreateCatalog
{
    class Message
    {
        public static string SubUnitName(string PlantCode, string MainUnitCode, string SubUnitCode)
        {
            string myKey = PlantCode + "*" + MainUnitCode + "*" + SubUnitCode;
            if (PedmsDataAccess.myCache.mySubUnit.ContainsKey(myKey))
            {
                return PedmsDataAccess.myCache.mySubUnit[myKey].ToString();
            }
            else
            {
                int scu;
                DataSet ds = new DataSet();
                ds = PIMS_ClassLib.DataAccess.GetDataSet(frmMain.sConstrBase, "select sub_unit_name from vw_wbs_v2 where plant_no = '" + PlantCode + "' and main_unit_no = '" + MainUnitCode + "' and sub_unit_no = '" + SubUnitCode + "' and project_no = '" + frmMain.sProjectNo + "'", out scu);
                if (scu == 1)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        string temp = ds.Tables[0].Rows[0][0].ToString();
                        PedmsDataAccess.myCache.mySubUnit.Add(myKey, temp);
                        return temp;
                    }
                    else
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
            }
        }
        public static string MainUnitName(string PlantCode, string MainUnitCode)
        {
            string myKey = PlantCode + "*" + MainUnitCode;
            if (PedmsDataAccess.myCache.myMainUnit.ContainsKey(myKey))
            {
                return PedmsDataAccess.myCache.myMainUnit[myKey].ToString();
            }
            else
            {
                int scu;
                DataSet ds = new DataSet();
                ds = PIMS_ClassLib.DataAccess.GetDataSet(frmMain.sConstrBase, "select main_unit_name from vw_wbs_v2 where plant_no = '" + PlantCode + "' and main_unit_no = '" + MainUnitCode + "' and project_no = '" + frmMain.sProjectNo + "'", out scu);
                if (scu == 1)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        string temp = ds.Tables[0].Rows[0][0].ToString();
                        PedmsDataAccess.myCache.myMainUnit.Add(myKey, temp);
                        return temp;
                    }
                    else
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
            }
        }
        public static string PlantName(string PlantCode)
        {
            if (PedmsDataAccess.myCache.myPlant.ContainsKey(PlantCode))
            {
                return PedmsDataAccess.myCache.myPlant[PlantCode].ToString();
            }
            else
            {
                int scu;
                DataSet ds = new DataSet();
                ds = PIMS_ClassLib.DataAccess.GetDataSet(frmMain.sConstrBase, "select plant_name from vw_wbs_v2 where plant_no = '" + PlantCode + "' and project_no = '" + frmMain.sProjectNo + "'", out scu);
                if (scu == 1)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        string temp = ds.Tables[0].Rows[0][0].ToString();
                        PedmsDataAccess.myCache.myPlant.Add(PlantCode, temp);
                        return temp;
                    }
                    else
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
            }
        }
        public static string SubDisciplineName(string SubDisciplineCode)
        {
            int scu;
            DataSet ds = new DataSet();
            ds = PIMS_ClassLib.DataAccess.GetDataSet(frmMain.sConstrBase, "select sub_discipline_name from tb_sub_discipline_v2 where sub_discipline_no = '" + SubDisciplineCode + "'", out scu);
            if (scu == 1)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    return ds.Tables[0].Rows[0][0].ToString();
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }
        public static string StageName(string StageId)
        {
            int scu;
            DataSet ds = new DataSet();
            ds = PIMS_ClassLib.DataAccess.GetDataSet(frmMain.sConstrBase, "select stage_name from tb_stage where stage_id = '" + StageId + "'", out scu);
            if (scu == 1)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    return ds.Tables[0].Rows[0][0].ToString();
                }
                else
                {
                    return "";
                }

            }
            else
            {
                return "";
            }
        }
    }
}
