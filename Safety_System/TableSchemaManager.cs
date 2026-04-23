/// FILE: Safety_System/TableSchemaManager.cs ///
using System.Collections.Generic;

namespace Safety_System
{
    public static class TableSchemaManager
    {
        // 集中管理所有資料表的欄位定義 Schema
        public static readonly Dictionary<string, string> SchemaMap = new Dictionary<string, string>
        {
            { "TargetManagement", "[年度] TEXT, [修訂日] TEXT, [單位] TEXT, [目標名稱] TEXT, [管理目標計畫表編號] TEXT, [施實重點項目1] TEXT, [日程1] TEXT, [施實重點項目2] TEXT, [日程2] TEXT, [施實重點項目3] TEXT, [日程3] TEXT, [施實重點項目4] TEXT, [日程4] TEXT, [施實重點項目5] TEXT, [日程5] TEXT, [預估成本] TEXT, [預估成效] TEXT, [計畫績效指標] TEXT, [績效指標計算方式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "NearMiss", "[日期] TEXT, [時間] TEXT, [文件編號] TEXT, [單位] TEXT, [地點] TEXT, [事故類別] TEXT, [事故類型] TEXT, [發生經過] TEXT, [改善對策] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            //巡檢異常單
			{ "SafetyInspection", "[日期] TEXT, [單位] TEXT, [表單單號] TEXT, [表單主題] TEXT, [重要性] TEXT, [狀態] TEXT, [申請者] TEXT, [承辦人] TEXT, [目前處理者] TEXT, [	申請時間] TEXT, [修改時間] TEXT, [到期時間] TEXT, [缺失責任人] TEXT, [危害類型主項] TEXT, [危害類型細分類] TEXT, [違規樣態類型] TEXT, [列入安全觀查事項] TEXT, [列入虛驚事項] TEXT, [不安全行為] TEXT, [廠內曾發生工傷事件項目] TEXT, [違規分類] TEXT, [違反規定名稱] TEXT, [違反規定條款] TEXT, [建議改善事項] TEXT, [追蹤改善狀況] TEXT, [改善進度] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "SafetyObservation", "[日期] TEXT, [區域] TEXT, [類別] TEXT, [描述] TEXT, [觀查人] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "TrafficInjury", "[日期] TEXT, [姓名] TEXT, [地點] TEXT, [狀態] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WorkInjury", "[日期] TEXT, [姓名] TEXT, [受傷部位] TEXT, [原因] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "MinorInjury", "[日期] TEXT, [姓名] TEXT, [單位] TEXT, [發生經過] TEXT, [處置方式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_IL", "[年月] TEXT, [生產加不良MT] TEXT, [生產量MT] TEXT, [丁基膠MT] TEXT, [結構膠MT] TEXT, [鋁條MT] TEXT, [乾燥劑MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_LM", "[年月] TEXT, [生產加不良MT] TEXT, [生產量MT] TEXT, [PVB膜MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_CR", "[年月] TEXT, [鍍膜加不良MT] TEXT, [鍍膜生產量MT] TEXT, [除膜生產加不良MT] TEXT, [除膜成品產量MT] TEXT, [靶材MT] TEXT, [隔離粉MT] TEXT, [氧化鈰MT] TEXT, [噴砂底板成品鋁MT] TEXT, [噴砂底板成品其他MT] TEXT, [氧化鋁砂金鋼砂MT] TEXT, [D1099MT] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_T", "[年月] TEXT, [強化生產加不良MT] TEXT, [強化生產量] TEXT, [砂布輪MT] TEXT, [砂帶MT] TEXT, [印刷生產加不良MT] TEXT, [印刷生產量MT] TEXT, [油墨MT] TEXT, [網板清洗劑MT] TEXT, [D1599MT] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_GCTE", "[年月] TEXT, [切片生產加不良MT] TEXT, [切片生產量MT] TEXT, [磨邊生產量不良MT] TEXT, [砂布輪砂輪片MT] TEXT, [鑽孔生產加不良MT] TEXT, [鑽孔生產量MT] TEXT, [] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_ML", "[年月] TEXT, [甲醇MT] TEXT, [乙醇MT] TEXT, [潤滑油MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_Water", "[年月] TEXT, [D0902MT] TEXT, [D0201MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ESG_Performance", "[年月] TEXT, [單位] TEXT, [項目] TEXT, [說明] TEXT, [預計執行週期] TEXT, [預估可節省或改善之數據] TEXT, [費用TWD] TEXT, [回應窗口] TEXT, [績效追蹤1] TEXT, [績效追蹤2] TEXT, [統計至12月底之實際數據含計算式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "AirPollution", "[年度] TEXT, [季度] TEXT, [排放量] TEXT, [繳費金額] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//公共危險物統計表
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
            { "FireEquip", "[日期] TEXT, [設備名稱] TEXT, [編號] TEXT, [位置] TEXT, [有效日期] TEXT, [檢查結果] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            //各單位消防自主檢查表
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
            //請購資料
			{ "PurchaseData", "[日期] TEXT, [開單單位] TEXT, [請購單號] TEXT, [項次] TEXT, [料號] TEXT, [料名] TEXT, [規格] TEXT, [用途] TEXT, [ESG分類] TEXT, [數量] TEXT, [未稅單價] TEXT, [結案日期] TEXT, [附件檔案] TEXT, [備註] TEXT" }
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
			
			if (tableName == "SafetyInspection" && columnName == "單位") 
				return new[] { "改切", "磨邊", "強化", "膠合", "複層", "鍍板", "儲運", "品管", "物料", "事務", "工安", "維修", "儀電", "廠務", "製程", "加工" };
			if (tableName == "SafetyInspection" && columnName == "改善進度") 
				return new[] { "未結案", "已結案"};            
			return null; // 若沒有特定設定，回傳 null 代表維持一般文字方塊
        }
    }
}
