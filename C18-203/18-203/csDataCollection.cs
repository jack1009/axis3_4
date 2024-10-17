using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaigeneDataCollection
{
    public class csDataCollection
    {
        #region 上端命令
        /// <summary>
        /// 人員操作資格 D40007
        /// 0001:許可
        /// 0011:不許可
        /// </summary>
        public short EnableOperator { get; set; }//
        /// <summary>
        /// 設備端機種/材料庫存 D40008
        /// 1001 : 機種/材料-設備可操作
        /// 1011 : 機種材料不匹配-設備禁止操做
        /// 1111 : 材料庫存數不足-設備禁止操做
        /// </summary>
        public short EnableWorkTypesOrStockOut { get; set; }
        /// <summary>
        /// TCS 材料控管 D40009
        /// 2001 : 材料批號 OK-設備可操作
        /// 2011 : 材料批號異常-設備禁止操做
        /// 2111 : 材料過期確認-設備禁止操做
        /// </summary>
        public short EnableParts { get; set; }
        /// <summary>
        /// TCS 生產管控 D40010
        /// 3001 : 當下工件 ID 在前製程判定為 OK 品-設備可操作
        /// 3011 : 當下工件 ID 在前製程判定為 NG 品-設備禁止操做
        /// 3111 : 當下設備設定參數與原始設定不同-設備禁止操做
        /// 3100 : 當下設備管制範圍與原始設定不同-設備禁止操做
        /// 3101 : 當下工件連續三次連續NG設備警示顯示(不停機)
        /// 3110 : 當下工件2HR達10次NG設備警示顯示(不停機)
        /// </summary>
        public short EnableMachineRun { get; set; }
        /// <summary>
        /// 機台運轉總時間-時(自動程式) D40578
        /// 24 hr
        /// </summary>
        public short OeeMachineRunHour { get; set; }
        /// <summary>
        /// 機台運轉總時間-分(自動程式) D40579
        /// 59 min
        /// </summary>
        public short OeeMachineRunMin { get; set; }
        /// <summary>
        /// 機台運轉總時間-秒(自動程式) D40580
        /// 59 sec
        /// </summary>
        public short OeeMachineRunSec { get; set; }
        #endregion

        #region 通訊
        /// <summary>
        /// 設備動作 D40003
        /// 0:設備運轉停止
        /// 1:設備運轉中
        /// </summary>
        public short MachineStatus { get; set; }
        /// <summary>
        /// PLC 機種 ID D40011
        /// 1001 : PLC 機種 ID 送出
        /// 1000 : PLC 無機種 ID
        /// </summary>
        public short WorkID { get; set; }
        /// <summary>
        /// PLC 材料 ID(Read)設備=>TCS D40012
        /// 2001 : PLC 材料 ID 送出
        /// 2000 : PLC 無材料 ID->TCS資料讀取後寫回2000，設備將材料清除
        /// </summary>
        public short FlagPLCPartsID { get; set; }
        /// <summary>
        /// PLC 作業後狀態資料(Read)設備=>TCS D40013
        /// 3001 : PLC 作業後狀態資料送出 ->  通知 TCS 進行資料收集動作 
        /// 3000 : PLC 作業後狀態資料取走 -> TCS 資料取出後複寫( 3001 -> 3000 )
        /// </summary>
        public short FlagPLCRunFinish { get; set; }
        /// <summary>
        /// PLC 作業後狀態資料(Write)設備<=TCS D40014
        /// 7001 : PLC 作業後狀態資料送出 ->  通知 TCS 進行資料收集動作 
        /// 7000 : PLC 作業後狀態資料取走 -> TCS 資料取出後複寫( 3001 -> 3000 )
        public short FlagMainIDChanged { get; set; }
        /// <summary>
        /// PLC 材料 ID(Write)設備<=TCS D40015
        /// 2001 : PLC 材料 ID 送出
        /// 2000 : PLC 無材料 ID->TCS資料讀取後寫回2000，設備將材料清除
        /// </summary>
        public short FlagWorkTypeChanged { get; set; }
        #endregion

        #region 基本設定
        /// <summary>
        /// LINE No. D40025
        /// 線別： 1,2,3….
        /// </summary>
        public short BasicSettingLineNo { get; set; }
        /// <summary>
        /// 機種(佔用兩個D值）ASCII D40026,D40027
        /// 197,209...
        /// </summary>
        public Int16[] BasicSettingWorkType { get; set; }
        /// <summary>
        /// 設備即時-年 D40028
        /// 2021 year
        /// </summary>
        public short BasicSettingYear { get; set; }
        /// <summary>
        /// 設備即時-月 D40029
        /// 12 month
        /// </summary>
        public short BasicSettingMonth { get; set; }
        /// <summary>
        /// 設備即時-日 D40030
        /// 31 day
        /// </summary>
        public short BasicSettingDay { get; set; }
        /// <summary>
        /// 設備即時-時 D40031
        /// 24 hr
        /// </summary>
        public short BasicSettingHour { get; set; }
        /// <summary>
        /// 設備即時-分 D40032
        /// 59 min
        /// </summary>
        public short BasicSettingMin { get; set; }
        /// <summary>
        /// 設備即時-秒 D40033
        /// 59 sec
        /// </summary>
        public short BasicSettingSec { get; set; }
        /// <summary>
        /// 設備工位 D40034
        ///  1 or 2 ,3…..
        /// </summary>
        public short BasicSettingEquipmentNo { get; set; }
        #endregion

        #region 稼動率
        /// <summary>
        /// 電源投入總時間-時 D40575
        /// 24 hr
        /// </summary>
        public short OeePowerOnHour { get; set; }
        /// <summary>
        /// 電源投入總時間-分 D40576
        /// 59 min
        /// </summary>
        public short OeePowerOnMin { get; set; }
        /// <summary>
        /// 電源投入總時間-秒 D40577
        /// 59 sec
        /// </summary>
        public short OeePowerOnSec { get; set; }
        /// <summary>
        /// 循環時間 D40581
        /// 1234.5sec
        /// </summary>
        public short OeeCycleTime { get; set; }
        /// <summary>
        /// 生產數0~65535 D40582
        /// </summary>
        public short OeeTotalCount { get; set; }
        /// <summary>
        /// 良品數0~65535 D40583
        /// </summary>
        public short OeeOkCount { get; set; }
        /// <summary>
        /// 不良品數0~65535 D40584
        /// </summary>
        public short OeeNgCount { get; set; }
        /// <summary>
        /// 生產數歸零的設定時間-時 D40585
        /// </summary>
        public short OeeResetCountHour { get; set; }
        /// <summary>
        /// 生產數歸零的設定時間-分 D40586
        /// </summary>
        public short OeeResetCountMin { get; set; }
        /// <summary>
        /// 生產數歸零的設定時間-秒 D40587
        /// </summary>
        public short OeeResetCountSec { get; set; }
        #endregion

        #region 製程ID
        /// <summary>
        /// ID1-本體 ASCII 2097010218125000001      長度19字請保留24 D40675~D40686
        /// </summary>
        public Int16[] MainID { get; set; }
        public void setMainID(string id)
        {
            string s = id;
            byte[] bts = Encoding.ASCII.GetBytes(s);
            for (int i = 0; i < bts.Length; i += 2)
            {
                if (i == bts.Length - 1)
                {
                    int i1 = bts[i];
                    MainID[i / 2] = (short)i1;
                }
                else
                {
                    int i1 = bts[i + 1] << 8;
                    i1 += bts[i];
                    MainID[i / 2] = (short)i1;
                }
            }
        }
        public string getMainID()
        {
            string s = "";
            List<byte> bts = new List<byte>();
            foreach (var x in MainID)
            {
                int i = x;
                int i1 = 0xFF;
                i = i & i1;
                bts.Add((byte)i);
                int i2 = x;
                i1 = 0xFF00;
                i2 = i2 & i1;
                i2 = i2 >> 8;
                bts.Add((byte)i2);
            }
            byte[] bt = new byte[bts.Count];
            for (int i = 0; i < bt.Length; i++)
            {
                bt[i] = bts[i];
            }
            s = Encoding.ASCII.GetString(bt);
            return s;
        }
        #endregion
        #region 製程ID2(RS,TMR)
        /// <summary>
        /// ID2-RS 2097010218125000001      長度19字請保留24
        /// </summary>
        public Int16[] MainID2 { get; set; }
        public void setMainID2(string id)
        {
            string s = id;
            byte[] bts = Encoding.ASCII.GetBytes(s);
            for (int i = 0; i < bts.Length; i += 2)
            {
                if (i == bts.Length - 1)
                {
                    int i1 = bts[i];
                    MainID2[i / 2] = (short)i1;
                }
                else
                {
                    int i1 = bts[i + 1] << 8;
                    i1 += bts[i];
                    MainID2[i / 2] = (short)i1;
                }
            }
        }
        public string getMainID2()
        {
            string s = "";
            List<byte> bts = new List<byte>();
            foreach (var x in MainID2)
            {
                int i = x;
                int i1 = 0xFF;
                i = i & i1;
                bts.Add((byte)i);
                int i2 = x;
                i1 = 0xFF00;
                i2 = i2 & i1;
                i2 = i2 >> 8;
                bts.Add((byte)i2);
            }
            byte[] bt = new byte[bts.Count];
            for (int i = 0; i < bt.Length; i++)
            {
                bt[i] = bts[i];
            }
            s = Encoding.ASCII.GetString(bt);
            return s;
        }
        #endregion
        #region 材料ID
        /// <summary>
        /// 材料ID #*TP100025R09-A*9999*2019070912345*20207031*11TG00*0523   長度55字請保留70 材料1D40075~D40109 材料2D40120~D40154
        /// 材料數量 材料1D40115 材料2D40160
        /// </summary>
        public csPartID[] PartIDs { get; set; }
        public class csPartID
        {
            public short[] ID { get; set; }
            /// <summary>
            /// 數量(材料數量顯示)由系統（TCS）端寫入材料剩餘數量
            /// </summary>
            public short PartsCount { get; set; }
            public csPartID()
            {
                ID = new short[35];
            }
            public void setPartID(string id)
            {
                string s = id;
                byte[] bts = Encoding.ASCII.GetBytes(s);
                if (bts.Length==0)
                {
                    for (int j = 0; j < ID.Length; j++)
                    {
                        ID[j] = 0;
                    }
                }
                for (int i = 0; i < bts.Length; i+=2)
                {
                    if (i == bts.Length - 1)
                    {
                        int i1 = bts[i];
                        ID[i / 2] = (short)i1;
                    }
                    else
                    {
                        int i1 = bts[i + 1] << 8;
                        i1 += bts[i];
                        ID[i / 2] = (short)i1;
                    }
                }
            }
            public string getPartID()
            {
                string s = "";
                List<byte> bts = new List<byte>();
                foreach (var x in ID)
                {
                    int i = x;
                    int i1 = 0xFF;
                    i = i & i1;
                    bts.Add((byte)i);
                    int i2 = x;
                    i1 = 0xFF00;
                    i2 = i2 & i1;
                    i2 = i2 >> 8;
                    bts.Add((byte)i2);
                }
                byte[] bt = new byte[bts.Count];
                for (int i = 0; i < bt.Length; i++)
                {
                    bt[i] = bts[i];
                }
                s = Encoding.ASCII.GetString(bt);
                return s;
            }
        }
        #endregion

        #region 測試數據
        /// <summary>
        /// 總合判定結果 (產出結果) D414374
        /// OK:NG
        /// or 1:OK;0:NG
        /// </summary>
        public short OverrallJudgement { get; set; }
        /// <summary>
        /// 鎖付數據
        /// 扭力數據結果1 D41075 扭力判定結果1 D41076 扭力數據結果2 D41077 扭力判定結果2 D41078 扭力數據結果3 D41079 扭力判定結果3 D41080
        /// 角度數據結果1 D41087 角度判定結果1 D41088 角度數據結果2 D41089 角度判定結果2 D41090 角度數據結果3 D41091 角度判定結果3 D41092
        /// </summary>
        public csScrewRow[] ScrewRows { get; set; }
        public class csScrewRow
        {
            /// <summary>
            /// 數據結果-扭力值
            /// </summary>
            public short TorqueValue { get; set; }
            /// <summary>
            /// Set扭力值
            /// </summary>
            /// <param name="torque">輸入扭力值</param>
            public void SetTorqueValue(string torque)
            {
                double d;
                bool b = double.TryParse(torque, out d);
                if (b)
                {
                    d = d * 100;
                    short i = (short)d;
                    TorqueValue = i;
                }
            }
            /// <summary>
            /// 判定結果-扭力值
            /// </summary>
            public short TorqueJudgement { get; set; }
            /// <summary>
            /// Set扭力判定結果
            /// </summary>
            /// <param name="judgement">SDB回應字串-A:OK,H:high,L:low</param>
            public void SetTorqueJudgement(string judgement)
            {
                if (judgement.Equals("A"))
                {
                    TorqueJudgement = 1;
                }
                else
                {
                    TorqueJudgement = 0;
                }
            }
            /// <summary>
            /// 數據結果-鎖緊角度
            /// </summary>
            public short AngleValue { get; set; }
            /// <summary>
            /// Set角度值
            /// </summary>
            /// <param name="Angle">輸入角度值</param>
            public void SetAngleValue(string Angle)
            {
                double d;
                bool b = double.TryParse(Angle, out d);
                if (b)
                {
                    d = d * 10;
                    short i = (short)d;
                    AngleValue = i;
                }
            }
            /// <summary>
            /// 判定結果-鎖緊角度
            /// </summary>
            public short AngleJudgement { get; set; }
            /// <summary>
            /// Set角度判定結果
            /// </summary>
            /// <param name="judgement">SDB回應字串-A:OK,H:high,L:low</param>
            public void SetAngleJudgement(string judgement)
            {
                if (judgement.Equals("A"))
                {
                    AngleJudgement = 1;
                }
                else
                {
                    AngleJudgement = 0;
                }
            }
        }
        #endregion

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="numParts">材料數量</param>
        /// <param name="numScrew">鎖付軸數</param>
        public csDataCollection(int numParts,int numScrew)
        {
            BasicSettingWorkType = new Int16[2];
            MainID = new Int16[12];
            MainID2 = new Int16[12];
            PartIDs = new csPartID[numParts];
            for (int i = 0; i < numParts; i++)
            {
                PartIDs[i] = new csPartID();
            }
            ScrewRows = new csScrewRow[numScrew];
            for (int i = 0; i < numScrew; i++)
            {
                ScrewRows[i] = new csScrewRow();
            }
        }
    }
    
}
