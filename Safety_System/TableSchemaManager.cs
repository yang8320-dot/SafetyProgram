/// FILE: Safety_System/TableSchemaManager.cs ///
using System.Collections.Generic;
using System.Linq;

namespace Safety_System
{
    public static class TableSchemaManager
    {
        // 🟢 統一管理自定義選單的預設欄位結構 (Single Source of Truth)
        public const string DefaultCustomSchema = "[日期] TEXT, [內容] TEXT, [附件檔案] TEXT, [備註] TEXT";

        // 集中管理所有資料表的欄位定義 Schema
        public static readonly Dictionary<string, string> SchemaMap = new Dictionary<string, string>
        {
            { "TargetManagement", "[年度] TEXT, [修訂日] TEXT, [單位] TEXT, [目標名稱] TEXT, [管理目標計畫表編號] TEXT, [施實重點項目1] TEXT, [日程1] TEXT, [施實重點項目2] TEXT, [日程2] TEXT, [施實重點項目3] TEXT, [日程3] TEXT, [施實重點項目4] TEXT, [日程4] TEXT, [施實重點項目5] TEXT, [日程5] TEXT, [預估成本] TEXT, [預估成效] TEXT, [計畫績效指標] TEXT, [績效指標計算方式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            
            // 🟢 ISO14001 環境溝通 4 張表
            { "EnvInfoReceive", "[日期] TEXT, [表單單號] TEXT, [來文發文] TEXT, [發文單位] TEXT, [主旨] TEXT, [相關單位] TEXT, [結案] TEXT, [簽核] TEXT, [利害相關者] TEXT, [溝通方式] TEXT, [附件檔案] TEXT, [備註] TEXT, [連結] TEXT" },
            { "InternalComm", "[日期] TEXT, [表單單號] TEXT, [來文發文] TEXT, [發文單位] TEXT, [主旨] TEXT, [內文] TEXT, [聯絡書] TEXT, [相關單位] TEXT, [結案] TEXT, [簽核] TEXT, [利害相關者] TEXT, [溝通方式] TEXT, [附件檔案] TEXT, [備註] TEXT, [連結] TEXT" },
            { "MailReceive", "[日期] TEXT, [表單單號] TEXT, [來文發文] TEXT, [發文單位] TEXT, [主旨] TEXT, [內文] TEXT, [聯絡書] TEXT, [相關單位] TEXT, [結案] TEXT, [簽核] TEXT, [利害相關者] TEXT, [附件檔案] TEXT, [溝通方式] TEXT, [備註] TEXT, [連結] TEXT" },
            { "VisitorRecord", "[日期] TEXT, [發文單位] TEXT, [拜訪目的] TEXT, [拜訪人員] TEXT, [內容概述] TEXT, [會同人員] TEXT, [聯絡書] TEXT, [相關單位] TEXT, [結案] TEXT, [簽核] TEXT, [利害相關者] TEXT, [溝通方式] TEXT, [附件檔案] TEXT, [備註] TEXT" },

            { "NearMiss", "[日期] TEXT, [時間] TEXT, [文件編號] TEXT, [單位] TEXT, [地點] TEXT, [事故類別] TEXT, [事故類型] TEXT, [發生經過] TEXT, [改善對策] TEXT, [屬同事件] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SafetyInspection", "[日期] TEXT, [單位] TEXT, [表單單號] TEXT, [表單主題] TEXT, [申請者] TEXT, [缺失責任人] TEXT, [危害類型主項] TEXT, [危害類型細分類] TEXT, [違規樣態類型] TEXT, [列入安全觀查事項] TEXT, [列入虛驚事項] TEXT, [不安全行為] TEXT, [廠內曾發生工傷事件項目] TEXT, [違規分類] TEXT, [違反規定名稱] TEXT, [違反規定條款] TEXT, [建議改善事項] TEXT, [追蹤改善狀況] TEXT, [改善進度] TEXT, [附件檔案] TEXT, [備註] TEXT, [連結] TEXT" },
            { "SafetyObservation", "[日期] TEXT, [區域] TEXT, [類別] TEXT, [描述] TEXT, [觀查人] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "TrafficInjury", "[日期] TEXT, [姓名] TEXT, [地點] TEXT, [狀態] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WorkInjury", "[日期] TEXT, [單位] TEXT, [姓名] TEXT, [受傷部位] TEXT, [原因] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "MinorInjury", "[日期] TEXT, [時間] TEXT, [文件編號] TEXT, [單位] TEXT, [地點] TEXT, [事故類別] TEXT, [事故類型] TEXT, [發生經過] TEXT, [改善對策] TEXT, [屬同事件] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_IL", "[年月] TEXT, [生產加不良MT] TEXT, [生產量MT] TEXT, [丁基膠MT] TEXT, [結構膠MT] TEXT, [鋁條MT] TEXT, [乾燥劑MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_LM", "[年月] TEXT, [生產加不良MT] TEXT, [生產量MT] TEXT, [PVB膜MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_CR", "[年月] TEXT, [鍍膜加不良MT] TEXT, [鍍膜生產量MT] TEXT, [除膜生產加不良MT] TEXT, [除膜成品產量MT] TEXT, [靶材MT] TEXT, [隔離粉MT] TEXT, [氧化鈰MT] TEXT, [噴砂底板成品鋁MT] TEXT, [噴砂底板成品其他MT] TEXT, [氧化鋁砂金鋼砂MT] TEXT, [D1099MT] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_T", "[年月] TEXT, [強化生產加不良MT] TEXT, [強化生產量] TEXT, [砂布輪MT] TEXT, [砂帶MT] TEXT, [印刷生產加不良MT] TEXT, [印刷生產量MT] TEXT, [油墨MT] TEXT, [網板清洗劑MT] TEXT, [D1599MT] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_GCTE", "[年月] TEXT, [切片生產加不良MT] TEXT, [切片生產量MT] TEXT, [磨邊生產量不良MT] TEXT, [砂布輪砂輪片MT] TEXT, [鑽孔生產加不良MT] TEXT, [鑽孔生產量MT] TEXT, [] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_ML", "[年月] TEXT, [甲醇MT] TEXT, [乙醇MT] TEXT, [潤滑油MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_Water", "[年月] TEXT, [D0902MT] TEXT, [D0201MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ESG_Performance", "[年月] TEXT, [單位] TEXT, [項目] TEXT, [說明] TEXT, [預計執行週期] TEXT, [預估可節省或改善之數據] TEXT, [費用TWD] TEXT, [回應窗口] TEXT, [績效追蹤1] TEXT, [績效追蹤2] TEXT, [統計至12月底之實際數據含計算式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "AirPollution", "[年度] TEXT, [季度] TEXT, [排放量] TEXT, [繳費金額] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "HazardStats", "[年月] TEXT, [單位] TEXT, [品項] TEXT, [品項單位] TEXT, [數量] TEXT, [使用量] TEXT, [庫存量] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WorkInjuryReport", "[年月] TEXT, [受僱勞工男性人數] TEXT, [非屬受僱勞工之其他工作者男性人數] TEXT, [受僱勞工女性人數] TEXT, [非屬受僱勞工之其他工作者女性人數] TEXT, [受僱勞工總計工作日數] TEXT, [非屬受僱勞工之其他工作者總計工作日數] TEXT, [受僱勞工總經歷工時] TEXT, [非屬受僱勞工之其他工作者總經歷工時] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "EnvTesting", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [分類] TEXT, [中文名稱] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ExposureLimits", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [種類] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [CASNO] TEXT, [容許濃度ppm] TEXT, [容許濃度mgm3] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ToxicSubstances", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [CASNO] TEXT, [管制濃度百分比] TEXT, [分級運作量kg] TEXT, [毒性分類] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ConcernedChem", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [CASNO] TEXT, [管制濃度百分比] TEXT, [管制行為] TEXT, [具有危害性之關注化學物質註記] TEXT, [分級運作量kg] TEXT, [定期申報頻率] TEXT, [包裝容器規定] TEXT, [記錄] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "PriorityMgmtChem", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [CASNO] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ControlledChem", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [CASNO] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SpecificChem", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [類別] TEXT, [序] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "OrganicSolvents", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [類別] TEXT, [序] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [化學式] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WorkerHealthProtect", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [項次] TEXT, [序] TEXT, [中文名稱] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "PublicHazardous", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [分類] TEXT, [名稱] TEXT, [種類] TEXT, [分級] TEXT, [管制量] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "AirPollutionEmerg", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [項次] TEXT, [中文名稱] TEXT, [英文名稱] TEXT, [CASNO] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "FactoryHazardous", "[日期] TEXT, [法規名稱] TEXT, [依據法條] TEXT, [內容] TEXT, [分類] TEXT, [序] TEXT, [名稱] TEXT, [種類] TEXT, [管制量] TEXT, [確認日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SDS_Inventory", "[日期] TEXT, [廠內編號] TEXT, [化學物質名稱] TEXT, [其它化學物質名稱] TEXT, [危害標示] TEXT, [CAS_No] TEXT, [危害成份] TEXT, [危害分類] TEXT, [供應商] TEXT, [供應商地址] TEXT, [供應商電話] TEXT, [SDS版本日期] TEXT, [使用單位] TEXT, [使用地點] TEXT, [使用平均量] TEXT, [使用最大量] TEXT, [貯存地點] TEXT, [平均貯存量] TEXT, [最大貯存量] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "FireResponsible", "[單位] TEXT, [場所區域] TEXT, [防火負責人] TEXT, [火源責任人] TEXT, [責任代理人] TEXT, [更新日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "FireEquip", "[日期] TEXT, [責任單位] TEXT, [柱子] TEXT, [編號] TEXT, [位置] TEXT, [具體位置] TEXT, [中繼器] TEXT, [狀態列表] TEXT, [室內消防栓] TEXT, [戶外消防栓] TEXT, [消防水帶有效期限] TEXT, [火警綜合盤] TEXT, [緊急照明燈] TEXT, [緊急照明燈] TEXT, [逃生指示燈_右] TEXT, [逃生指示燈_左] TEXT, [逃生指示燈_出口] TEXT, [滅火器] TEXT, [滅火器有效期限] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "FireSelfInspection", "[日期] TEXT, [檢點表名稱] TEXT, [單位] TEXT, [檢查人] TEXT, [檢查結果] TEXT, [缺失描述] TEXT, [改善對策] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "訓練時數", "[日期] TEXT, [員工姓名] TEXT, [受訓項目] TEXT, [課程名稱] TEXT, [訓練時數] TEXT, [HR外訓申請] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "EnvMonitor", "[日期] TEXT, [SEG編號] TEXT, [測點名稱] TEXT, [噪音_db] TEXT, [粉塵_區域] TEXT, [粉塵_個人] TEXT, [一氧化鉛] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WastewaterPeriodic", "[日期] TEXT, [申報季別] TEXT, [排放水量] TEXT, [COD] TEXT, [SS] TEXT, [BOD] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "DrinkingWater", "[日期] TEXT, [採樣點位置] TEXT, [大腸桿菌群] TEXT, [總菌落數] TEXT, [鉛] TEXT, [濁度] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "IndustrialZoneTest", "[日期] TEXT, [採樣點位置] TEXT, [水溫] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [重金屬] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SoilGasTest", "[日期] TEXT, [採樣井編號] TEXT, [測漏氣體濃度] TEXT, [甲烷] TEXT, [二氧化碳] TEXT, [氧氣] TEXT, [檢測機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WastewaterSelfTest", "[日期] TEXT, [採樣時間] TEXT, [採樣位置] TEXT, [pH值] TEXT, [COD] TEXT, [SS] TEXT, [透視度] TEXT, [檢驗人員] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "CoolingWaterVendor", "[日期] TEXT, [廠商名稱] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [添加藥劑] TEXT, [檢驗結果] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "CoolingWaterSelf", "[日期] TEXT, [水溫] TEXT, [pH值] TEXT, [導電度] TEXT, [濁度] TEXT, [總鐵] TEXT, [銅離子] TEXT, [檢驗人員] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "TCLP", "[日期] TEXT, [樣品名稱] TEXT, [總鉛] TEXT, [鏓鉻] TEXT, [鏓銅] TEXT, [鏓鋇] TEXT, [總鎘] TEXT, [總硒] TEXT, [六價鉻] TEXT, [總汞] TEXT, [總砷] TEXT, [檢驗機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterMeterCalibration", "[日期] TEXT, [水錶編號] TEXT, [水錶位置] TEXT, [校正前讀數] TEXT, [校正後讀數] TEXT, [校正單位] TEXT, [下次校正日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "OtherTests", "[日期] TEXT, [檢測項目] TEXT, [檢測位置] TEXT, [檢測數值] TEXT, [單位] TEXT, [合格標準] TEXT, [檢測機構] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "HealthPromotion", "[日期] TEXT, [活動名稱] TEXT, [參與人數] TEXT, [執行單位] TEXT, [成果摘要] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "PurchaseData", "[日期] TEXT, [開單單位] TEXT, [請購單號] TEXT, [項次] TEXT, [料號] TEXT, [料名] TEXT, [規格] TEXT, [用途] TEXT, [ESG分類] TEXT, [數量] TEXT, [未稅單價] TEXT, [結案日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            
            // 🟢 個人隱藏選單 (Menu1 ~ Menu4)
            { "AccountManage", "[日期] TEXT, [系統名稱] TEXT, [網址] TEXT, [登入帳號] TEXT, [登入密碼] TEXT, [使用者] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "KPI", "[年月] TEXT, [單位] TEXT, [指標名稱] TEXT, [目標值] TEXT, [實際值] TEXT, [達成狀況] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "CultureImprove", "[年月] TEXT, [單位] TEXT, [項目] TEXT, [執行狀況] TEXT, [主責人員] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "PBC", "[年月] TEXT, [單位] TEXT, [姓名] TEXT, [考核項目] TEXT, [自評分數] TEXT, [主管評核] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "DataManage2", "[日期] TEXT, [分類] TEXT, [標題] TEXT, [內容說明] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "DataManage3", "[日期] TEXT, [分類] TEXT, [標題] TEXT, [內容說明] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "DataManage4", "[日期] TEXT, [分類] TEXT, [標題] TEXT, [內容說明] TEXT, [附件檔案] TEXT, [備註] TEXT" }
        };

        // =========================================================================
        // 🟢 下拉連動選單對應字典 (Mapping Dictionaries)
        // ==========================================
        // 1. 危害類型主項 -> 危害類型細分類
        public static readonly Dictionary<string, string[]> HazardSubCategoryMap = new Dictionary<string, string[]>
        {
            { "物理性", new[] { "PH1_物體飛落、掉落", "PH2_倒塌、崩塌", "PH3_物體破裂", "PH4_墜落、滾落", "PH5_跌倒、滑倒", "PH6_衝撞、被撞、碰撞", "PH7_夾、捲、壓傷", "PH8_切、割、刺、擦傷", "PH9_踩踏、踏穿", "PH10_溺斃", "PH11_與高低溫接觸、凍傷、灼燙傷", "PH12_噪音過高", "PH13_照明不足", "PH14_通風不良、缺氧、窒息", "PH16_游離輻射暴露", "PH17_非醫用游離輻射暴露", "PH18_振動、停電", "PH19_漏電、感電_含靜電及火花", "PH20_壓降", "PH21_漏水", "PH22_爆炸(塵爆)", "PH23_異常氣壓", "PH24_異物入眼" } },
            { "化學性", new[] { "CH1_火災", "CH2_爆炸", "CH3_與有害物接觸", "CH4_化學品洩漏_含廢液", "CH5_毒氣(氣體)洩漏", "CH6_異味", "CH7_冒煙", "CH8_缺氧，窒息", "CH9_化學品灼或濺傷" } },
            { "生物性", new[] { "BI1_病媒滋生", "BI2_食物中毒", "BI3_病菌傳染", "BI4_發霉腐敗", "BI5_過敏、不舒服" } },
            { "人因工程", new[] { "ER1_設計不良導致人為失誤", "ER2_操作高度、空間不適造", "ER3_人工搬運超過荷重造成", "ER4_不適宜之工作姿勢造成傷害", "ER5_重複性操作造成傷害", "ER6_人為不當動作" } },
            { "社會性", new[] { "EK1_不規律的工作時間", "EK2_工作不安全感", "EK3_職場暴力與恐嚇", "EK4_勞動人力老化", "EK5_工作的增強化(工作負荷量或是壓力的增加)" } },
            { "其他", new[] { "OT1_交通事故", "OT2_設備、設施損壞", "OT3_影響環境", "OT4_未歸類安全項目", "OT5_其他非安全項目", "OT6_消防相關" } }
        };

        // 2. 危害類型細分類 -> 違規樣態類型
        public static readonly Dictionary<string, string[]> ViolationTypeMap = new Dictionary<string, string[]>
        {
            { "PH1_物體飛落、掉落", new[] { "PH101-丟廢玻璃以拋丟方式【★★★】", "PH102-物件掉落【★★★】", "PH103-天車防滑蛇片失效【★★★★】", "PH104-吊掛作業未戴安全帽【★★★】", "PH105-天車燈座護網懸掛【★★★】", "PH106-其它未歸類物體飛落、掉落【★★★】" } },
            { "PH2_倒塌、崩塌", new[] { "PH201-操作堆高機未戴安全帽或未繫安全帶【★★★】", "PH202-天車吊具作業，人員離開操作位置【★★★】", "PH203-車架鏽蝕【★★★】", "PH204-支撐腳傾斜【★★】", "PH205-倚靠欄杆【★★】", "PH206-D架疊放玻璃二層【★★★】", "PH207-車架超出荷重【★★★】", "PH208-人員離開，設備機具未定位完成【★★★】", "PH209-車架未捆綁固定【★★】", "PH210-高壓鋼瓶未固定【★★★】", "PH211-原(物)料或廢棄物堆疊，過高或未穩固【★★★】", "PH212-車架玻璃破裂【★★★】", "PH213-倚靠玻璃【★★★】", "PH214-吊具一邊底部膠片脫離【★★★】", "PH299-其它未歸類物體倒塌、崩塌危害【★★★】" } },
            { "PH4_墜落、滾落", new[] { "PH401_高處作業未防護【★★★★】", "PH402_A字梯相關缺失【★★★】", "PH403_堆高機相關缺失【★★★】", "PH404_堆高機未繫安全帶、安全帽【★★★】", "PH405_工作階梯無適當扶手【★★★★】", "PH406_開口處未防護【★★★】", "PH407_人員站A字梯頂板作業【★★★★】" } },
            { "PH5_跌倒、滑倒", new[] { "PH501_地上玻璃粉塊未清除【★★★】", "PH502_電線橫跨走道，未有防護措施【★★★】", "PH503_鐵棒插在水溝蓋下方翹起【★★★】", "PH504_鐵板一端翹起【★★★】", "PH505_地上坑洞【★★★】", "PH506_地上凸出物【★★★】", "PH507_護欄未裝回【★★★】", "PH508_其它未歸類跌倒、滑倒【★★★】" } },
            { "PH6_衝撞、被撞、碰撞", new[] { "PH601-置放單片玻璃未採警示措施【★★★】", "PH602-車架置放玻璃尖端凸出未防護【★★★】", "PH603-通道置放車架未警示【★★★】", "PH604-車速過快【★★★★】", "PH605-人員操作堆高機缺失【★★★】", "PH606-堆高機停放手煞車未拉【★★★】", "PH607-堆高機前照燈損壞【★★★】", "PH608-堆高機相關缺失【★★★】", "PH609-行走時使用手機或檢視文件【★★★★】", "PH610-鏟車相關缺失【★★★】", "PH611-手動油壓拖板車相關缺失【★★★】", "PH612-車架置放玻璃凸出未防護【★★★】", "PH613-天車相關缺失【★★★】", "PH299-其它未歸類衝撞、被撞、碰撞【★★★】" } },
            { "PH7_夾、捲、壓傷", new[] { "PH701-防護罩未定位【★★★】", "PH702-防護罩損壞【★★★】", "PH703-砂輪機無護罩【★★★】", "PH704-人員穿越自動運轉設備【★★★★】", "PH705-人員未戴個人防護具【★★★】", "PH706護欄開口處掛鍊未掛上【★★★】", "PH707-電風扇護網中間網格太寬【★★★】" } },
            { "PH8_切、割、刺、擦傷", new[] { "PH801-設備、機具未有防護措施【★★★】", "PH802-玻璃未採警示措施【★★★】", "PH803-玻璃尖端、銳角未防護【★★★】", "PH804-地面碎玻璃未清理【★★★】", "PH805-燈管破裂【★★】", "PH806-車架玻璃玻裂【★★★】", "PH807-木條塊尖端、銳角未防護【★★★】", "PH808-背檔破損未防護【★★★】", "PH809-玻璃置放地上倚靠機台【★★★】", "PH810-人員未戴個人防護具【★★★】", "PH811-樣品版未穩妥置放【★★★】", "PH812-其它未歸類切、割、刺、擦傷危害【★★★】" } },
            { "PH9_踩踏、踏穿", new[] { "PH901_溝渠未加蓋【★★★】", "PH902_其它未歸類踩踏、踏穿危害【★★★】" } },
            { "PH19_漏電、感電_含靜電及火花", new[] { "PH1901_電動拖板車充電器無插頭護蓋【★★★】", "PH1902_插頭接地線端損壞【★★★】", "PH1903_馬達漏水及未固定【★★】", "PH1904_電風扇電線包覆鐵管鏽蝕【★★★】", "PH1905_電線置於潮濕地面上【★★★★】", "PH1906_未有接地功能【★★★】", "PH1907_插頭或插座損壞【★★★】", "PH1908_電線裸露【★★★★】", "PH1909_其它未歸類感電危害【★★★】" } },
            { "CH1_火災", new[] { "CH101_易燃物品儲存場所未設置【★★★】", "CH102_易燃物品存放於較高溫環境【★★★★★】", "CH103_油品暫存桶箱內勿有易燃物【★★★★】" } },
            { "CH3_與有害物接觸", new[] { "CH301_化學品作業未配安個人防護具【★★★】", "CH302_粉塵作業未戴防護具【★★★】", "CH303_玻璃研磨未使用集塵器【★★★】" } },
            { "CH4_化學品洩漏_含廢液", new[] { "CH401_未設有防止洩漏措施【★★★】" } },
            { "OT6_消防相關", new[] { "OT601-消防設備被遮蔽【★★】", "OT602-逃生動線受阻礙【★★】", "OT603-消防設備檢點未確實【★★】", "OT604-滅火器過期、銹蝕、損壞【★★★】", "OT605-消防設備缺少或未設置【★★★】", "OT606-火警標示燈箱底座銜接處損壞【★★★】" } },
            { "OT2_設備、設施損壞", new[] { "OT201_照明設備損壞【★★】", "OT202_落地架損壞【★★★】", "OT203_車架膠條破損【★★★】", "OT204_照明設備未固定【★★】", "OT205_光柵損壞【★★★】" } },
            { "OT4_未歸類安全項目", new[] { "OT401-自動檢查未確實【★★】", "OT402-工安告示板未更新【★】", "OT403-配電盤前動線阻礙【★★】", "OT404-無相關防護措施【★★★】", "OT405-化學品置放未妥善【★★★】", "OT406-作業人員未受訓或未取得相關資格【★★★】", "OT407-原(物)料、設備器具置放未妥善【★★★】", "OT408-緊急制動裝置損壞【★★★】", "OT409-相關標示不足【★★】", "OT410-個人防護具配戴未確實【★★★】", "OT411-車用滅火器無點檢卡【★★★】", "OT499-其它安全危害【★★★】" } },
            { "OT5_其他非安全項目", new[] { "OT501-管線銹蝕、漏水【★★】", "OT502-醫藥箱缺失【★★】", "OT503_能源損耗【★★】", "OT504_洗手台排出水流進雨水溝【★★】", "OT505_空壓機排出水流進雨水溝【★★】", "OT506_水噴槍管閥銜接處滴水【★★】" } },
            { "OT3_影響環境", new[] { "OT301-水溝內菸蒂【★★★】" } }
        };

        // 集中管理需顯示為下拉選單的資料表與欄位
        public static string[] GetDropdownList(string tableName, string columnName)
        {
            if (tableName == "PurchaseData" && columnName == "ESG分類") 
                return new[] { "", "碳費", "環境相關盤查", "再生能源費用", "環保設備", "製程優化", "穩定能源供應", "氣候災害韌性建立", "資源效率與循環", "綠色產品研發設計" };
            
            if (tableName == "FireSelfInspection" && columnName == "檢點表名稱") 
                return new[] { "日常火源自行檢查表", "防火避難設施自行檢查紀錄表", "消防安全設備自行檢查紀錄表" };
            
            if (tableName == "HazardStats" && columnName == "品項") 
                return new[] { "物料/超級柴油", "物料/甲醇(物料)", "物料/乙醇", "物料/潤滑油(物料)", "物料/無鉛汽油", "物料/煤油", "鍍板/甲醇(鍍板)", "鍍板/TMPTA", "工務/潤滑油(重機)", "工務/乙炔", "強化/網版清洗劑", "強化/油墨", "強化/油墨(中加移過來)", "水站/氯錠" };
            
            if (tableName == "HazardStats" && columnName == "品項單位") 
                return new[] { "公升", "公斤", "瓶" };
            
            if (tableName == "FireSelfInspection" && columnName == "單位") 
                return new[] { "改切", "磨邊", "強化", "膠合", "複層", "鍍板", "儲運", "品管", "物料", "事務", "工安", "維修", "儀電", "廠務", "製程", "加工"  };
            
            if (tableName == "SafetyInspection") 
            {
                if (columnName == "單位") return new[] { "改切", "磨邊", "強化", "膠合", "複層", "鍍板", "儲運", "品管", "物料", "事務", "工安", "維修", "儀電", "廠務", "製程", "加工" };
                if (columnName == "改善進度") return new[] { "未結案", "已結案"};    
                if (columnName == "違規分類") return new[] { "危險之虞", "廠內規範", "環境管理", "勞檢曾開立之缺失", "與法規抵觸", "工安建議改善項目"};             
                if (columnName == "危害類型主項") return new[] { "物理性", "化學性", "生物性", "人因工程", "社會性", "其他"};     
                if (columnName == "列入安全觀查事項" || columnName == "列入虛驚事項" || columnName == "不安全行為" || columnName == "廠內曾發生工傷事件項目") return new[] { "", "v"}; 

                if (columnName == "危害類型細分類")
                {
                    List<string> allSubCats = new List<string> { "" };
                    foreach (var arr in HazardSubCategoryMap.Values) allSubCats.AddRange(arr);
                    return allSubCats.Distinct().ToArray();
                }

                if (columnName == "違規樣態類型")
                {
                    List<string> allViolations = new List<string> { "" };
                    foreach (var arr in ViolationTypeMap.Values) allViolations.AddRange(arr);
                    return allViolations.Distinct().ToArray();
                }
            }

            return null; 
        }

        public static string[] GetDependentDropdownList(string tableName, string columnName, string parentValue)
        {
            if (tableName == "SafetyInspection")
            {
                if (columnName == "危害類型細分類")
                {
                    if (HazardSubCategoryMap.ContainsKey(parentValue))
                        return HazardSubCategoryMap[parentValue];
                }
                else if (columnName == "違規樣態類型")
                {
                    string searchKey = parentValue == "PH2_倒塌、崩塌、掉落" ? "PH2_倒塌、崩塌" : parentValue;
                    if (ViolationTypeMap.ContainsKey(searchKey))
                        return ViolationTypeMap[searchKey];
                }
            }
            return new string[] { "" }; 
        }
    }
}
