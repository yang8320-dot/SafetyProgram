/// FILE: Safety_System/TableSchemaManager.cs ///
using System.Collections.Generic;
using System.Linq;

namespace Safety_System
{
    public static class TableSchemaManager
    {
        public const string DefaultCustomSchema = "[日期] TEXT, [內容] TEXT, [附件檔案] TEXT, [備註] TEXT";

        public static readonly Dictionary<string, string> SchemaMap = new Dictionary<string, string>
        {
            // 水資源管理
            { "WaterMeterReadings", "[日期] TEXT, [星期] TEXT, [用電量] TEXT, [用電量日統計] TEXT, [廢水進流量] TEXT, [廢水進流量日統計] TEXT, [廢水處理量] TEXT, [廢水處理量日統計] TEXT, [水站廢水排放量] TEXT, [水站廢水排放量日統計] TEXT, [納管排放量] TEXT, [納管排放量日統計] TEXT, [回收水6吋] TEXT, [回收水6吋日統計] TEXT, [回收水雙介質A] TEXT, [回收水雙介質A日統計] TEXT, [回收水雙介質B] TEXT, [回收水雙介質B日統計] TEXT, [軟水A通量] TEXT, [軟水B通量] TEXT, [軟水C通量] TEXT, [濃縮水至冷卻水池] TEXT, [濃縮水至冷卻水池日統計] TEXT, [濃縮水至逆洗池] TEXT, [濃縮水至逆洗池日統計] TEXT, [貯存池至循環水池] TEXT, [貯存池至循環水池日統計] TEXT, [製程式至循環水池] TEXT, [製程式至循環水池日統計] TEXT, [污泥產出KG] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterChemicals", "[日期] TEXT, [星期] TEXT, [PAC_KG] TEXT, [NAOH_KG] TEXT, [高分子_KG] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterUsageDaily", "[日期] TEXT, [星期] TEXT, [廠區自來水使用量] TEXT, [行政區自來水使用量] TEXT, [自來水至貯存池] TEXT, [自來水至貯存池日統計] TEXT, [自來水量至清水池] TEXT, [自來水量至清水池日統計] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "DischargeData", "[年月] TEXT, [基本處理費(公噸)] TEXT, [SS(mg/l)] TEXT, [COD(mg/l)] TEXT, [BOD(mg/l)] TEXT, [氨氮(mg/l)] TEXT, [SS(KG))] TEXT, [COD(KG))] TEXT, [水量(元)] TEXT, [SS(元)] TEXT, [COD(元)] TEXT, [合計費用] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterVolume", "[年月] TEXT, [廠區自來水繳費單] TEXT, [行政區自來水繳費單] TEXT, [彰濱二廠自來水繳費單] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WaterPermitMaterial", "[許可字號] TEXT, [許可效期] TEXT, [類別] TEXT, [名稱] TEXT, [申請每日最大量] TEXT, [單位] TEXT, [附件檔案] TEXT, [備註] TEXT" },

            // 法規管理
            { "環保法規", "[日期] TEXT, [法規名稱] TEXT, [條] TEXT, [項] TEXT, [款] TEXT, [目] TEXT, [內容] TEXT, [重點摘要] TEXT, [適用性] TEXT, [有提升績效機會] TEXT, [有潛在不符合風險] TEXT, [鑑別日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "職安衛法規", "[日期] TEXT, [法規名稱] TEXT, [條] TEXT, [項] TEXT, [款] TEXT, [目] TEXT, [內容] TEXT, [重點摘要] TEXT, [適用性] TEXT, [有提升績效機會] TEXT, [有潛在不符合風險] TEXT, [鑑別日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "消防法規", "[日期] TEXT, [法規名稱] TEXT, [條] TEXT, [項] TEXT, [款] TEXT, [目] TEXT, [內容] TEXT, [重點摘要] TEXT, [適用性] TEXT, [有提升績效機會] TEXT, [有潛在不符合風險] TEXT, [鑑別日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "其它法規", "[日期] TEXT, [法規名稱] TEXT, [條] TEXT, [項] TEXT, [款] TEXT, [目] TEXT, [內容] TEXT, [重點摘要] TEXT, [適用性] TEXT, [有提升績效機會] TEXT, [有潛在不符合風險] TEXT, [鑑別日期] TEXT, [附件檔案] TEXT, [備註] TEXT" },

            // 其餘模組
            { "TargetManagement", "[年度] TEXT, [修訂日] TEXT, [單位] TEXT, [目標名稱] TEXT, [管理目標計畫表編號] TEXT, [施實重點項目1] TEXT, [日程1] TEXT, [施實重點項目2] TEXT, [日程2] TEXT, [施實重點項目3] TEXT, [日程3] TEXT, [施實重點項目4] TEXT, [日程4] TEXT, [施實重點項目5] TEXT, [日程5] TEXT, [預估成本] TEXT, [預估成效] TEXT, [計畫績效指標] TEXT, [績效指標計算方式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "EnvInfoReceive", "[日期] TEXT, [表單單號] TEXT, [來文發文] TEXT, [發文單位] TEXT, [主旨] TEXT, [相關單位] TEXT, [結案] TEXT, [簽核] TEXT, [利害相關者] TEXT, [溝通方式] TEXT, [附件檔案] TEXT, [備註] TEXT, [連結] TEXT" },
            { "InternalComm", "[日期] TEXT, [表單單號] TEXT, [來文發文] TEXT, [發文單位] TEXT, [主旨] TEXT, [內文] TEXT, [聯絡書] TEXT, [相關單位] TEXT, [結案] TEXT, [簽核] TEXT, [利害相關者] TEXT, [溝通方式] TEXT, [附件檔案] TEXT, [備註] TEXT, [連結] TEXT" },
            { "MailReceive", "[日期] TEXT, [表單單號] TEXT, [來文發文] TEXT, [發文單位] TEXT, [主旨] TEXT, [內文] TEXT, [聯絡書] TEXT, [相關單位] TEXT, [結案] TEXT, [簽核] TEXT, [利害相關者] TEXT, [附件檔案] TEXT, [溝通方式] TEXT, [備註] TEXT, [連結] TEXT" },
            { "VisitorRecord", "[日期] TEXT, [發文單位] TEXT, [拜訪目的] TEXT, [拜訪人員] TEXT, [內容概述] TEXT, [會同人員] TEXT, [聯絡書] TEXT, [相關單位] TEXT, [結案] TEXT, [簽核] TEXT, [利害相關者] TEXT, [溝通方式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            
			// ---工安---
			{ "NearMiss", "[日期] TEXT, [時間] TEXT, [文件編號] TEXT, [單位] TEXT, [地點] TEXT, [項目] TEXT, [事故類別] TEXT, [類型] TEXT, [發生經過] TEXT, [改善對策] TEXT, [屬同事件] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			{ "MinorInjury", "[日期] TEXT, [時間] TEXT, [文件編號] TEXT, [單位] TEXT, [地點] TEXT, [項目] TEXT, [事故類別] TEXT, [類型] TEXT, [發生經過] TEXT, [改善對策] TEXT, [屬同事件] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			{ "SafetyInspection", "[日期] TEXT, [單位] TEXT, [表單單號] TEXT, [表單主題] TEXT, [申請者] TEXT, [缺失責任人] TEXT, [危害類型主項] TEXT, [危害類型細分類] TEXT, [違規樣態類型] TEXT, [列入安全觀查事項] TEXT, [不安全行為] TEXT, [廠內曾發生工傷事件項目] TEXT, [違規分類] TEXT, [違反規定名稱] TEXT, [違反規定條款] TEXT, [建議改善事項] TEXT, [追蹤改善狀況] TEXT, [改善完成日] TEXT, [改善進度] TEXT, [附件檔案] TEXT, [備註] TEXT, [連結] TEXT" },
            { "SafetyObservation", "[項目編號] TEXT, [日期] TEXT, [單位] TEXT, [區域] TEXT, [觀查事項說明] TEXT, [建議改善說明] TEXT, [追蹤說明] TEXT, [結案日] TEXT, [本案進度] TEXT, [觀查人] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WorkInjury", "[日期] TEXT, [單位] TEXT, [姓名] TEXT, [職務] TEXT, [出生年月日] TEXT, [此工作連續服務] TEXT, [事故發生經過情形] TEXT, [事故發生原因分析] TEXT, [作業標準及訓練] TEXT, [就診醫院] TEXT, [是否住院] TEXT, [受傷程度] TEXT, [防止對策及改善措施] TEXT, [本事件負責人] TEXT, [原因類別] TEXT, [傷害類別] TEXT, [嚴重度] TEXT, [是否修訂標準書] TEXT, [工安意見] TEXT, [廠主管意見] TEXT, [失能天數] TEXT, [後追蹤說明] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			{ "TrafficInjury", "[日期] TEXT, [單位] TEXT, [姓名] TEXT, [職務] TEXT, [出生年月日] TEXT, [住址] TEXT, [電話] TEXT, [事故地點] TEXT, [交通工具] TEXT, [事故發生時間] TEXT, [原因] TEXT, [受傷程度] TEXT, [就診醫院] TEXT, [有無違反交通規則] TEXT, [有無配戴安全帽] TEXT, [駕照號碼] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "LaborInspection", "[日期] TEXT, [公文文號] TEXT, [缺失事項] TEXT, [法規] TEXT, [條] TEXT, [項] TEXT, [款] TEXT, [法令條款說明] TEXT, [限期改善或罰緩] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            
            // 廢棄物
            { "Waste_IL", "[年月] TEXT, [生產加不良MT] TEXT, [生產量MT] TEXT, [丁基膠MT] TEXT, [結構膠MT] TEXT, [鋁條MT] TEXT, [乾燥劑MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_LM", "[年月] TEXT, [生產加不良MT] TEXT, [生產量MT] TEXT, [PVB膜MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_CR", "[年月] TEXT, [鍍膜加不良MT] TEXT, [鍍膜生產量MT] TEXT, [除膜生產加不良MT] TEXT, [除膜成品產量MT] TEXT, [靶材MT] TEXT, [隔離粉MT] TEXT, [氧化鈰MT] TEXT, [噴砂底板成品鋁MT] TEXT, [噴砂底板成品其他MT] TEXT, [氧化鋁砂金鋼砂MT] TEXT, [D1099MT] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_T", "[年月] TEXT, [強化生產加不良MT] TEXT, [強化生產量] TEXT, [砂布輪MT] TEXT, [砂帶MT] TEXT, [印刷生產加不良MT] TEXT, [印刷生產量MT] TEXT, [油墨MT] TEXT, [網板清洗劑MT] TEXT, [D1599MT] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_GCTE", "[年月] TEXT, [切片生產加不良MT] TEXT, [切片生產量MT] TEXT, [磨邊生產量不良MT] TEXT, [砂布輪砂輪片MT] TEXT, [鑽孔生產加不良MT] TEXT, [鑽孔生產量MT] TEXT, [D2499MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_ML", "[年月] TEXT, [甲醇MT] TEXT, [乙醇MT] TEXT, [潤滑油MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "Waste_Water", "[年月] TEXT, [D0902MT] TEXT, [D0201MT] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WastePermitMaterial", "[許可字號] TEXT, [許可效期] TEXT, [製程代碼名稱] TEXT, [原料代碼] TEXT, [原料名稱] TEXT, [其它原料說明] TEXT, [其他製程說明] TEXT, [月最大使用量t] TEXT, [月平均使用量t] TEXT, [重量換算值] TEXT, [換算單位] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WastePermitProduct", "[許可字號] TEXT, [許可效期] TEXT, [製程代碼名稱] TEXT, [產品代碼] TEXT, [產品名稱] TEXT, [其它原料說明] TEXT, [其他製程說明] TEXT, [月最產出量t] TEXT, [月平均產出t] TEXT, [重量換算值] TEXT, [換算單位] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WastePermitWaste", "[許可字號] TEXT, [許可效期] TEXT, [製程代碼] TEXT, [其他製程說明] TEXT, [廢棄物代碼] TEXT, [廢棄物名稱] TEXT, [廢棄物月最大量] TEXT, [廢棄物月平均量] TEXT, [物理性質] TEXT, [有害特性] TEXT, [主要有害成分] TEXT, [貯存地點] TEXT, [貯存設施容施容量] TEXT, [貯存設施密閉性] TEXT, [清除方式] TEXT, [處理方式] TEXT, [中間處理方法] TEXT, [再利用管理方式] TEXT, [最終處置方式] TEXT, [產出廢液製程編號] TEXT, [清除頻率] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "WasteDisposalRecord", "[清運日期] TEXT, [廢棄物代碼] TEXT, [廢棄物名稱] TEXT, [清運重量] TEXT, [附件檔案] TEXT, [備註] TEXT" },

            // ESG 模組架構
            { "ESG_Performance", "[年月] TEXT, [單位] TEXT, [項目] TEXT, [說明] TEXT, [預計執行週期] TEXT, [預估可節省或改善之數據] TEXT, [費用TWD] TEXT, [回應窗口] TEXT, [績效追蹤1] TEXT, [績效追蹤2] TEXT, [統計至12月底之實際數據含計算式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ESG_OccupationalSafety", "[年度] TEXT, [部門] TEXT, [國際指標] TEXT, [ESG領域] TEXT, [指標分類] TEXT, [預防投入/預期改善] TEXT, [指標名稱] TEXT, [實際數據呈現] TEXT, [計算公式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ESG_HealthHygiene", "[年度] TEXT, [部門] TEXT, [國際指標] TEXT, [ESG領域] TEXT, [指標分類] TEXT, [預防投入/預期改善] TEXT, [指標名稱] TEXT, [實際數據呈現] TEXT, [計算公式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ESG_EnvironmentClimate", "[年度] TEXT, [部門] TEXT, [國際指標] TEXT, [ESG領域] TEXT, [指標分類] TEXT, [預防投入/預期改善] TEXT, [指標名稱] TEXT, [實際數據呈現] TEXT, [計算公式] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "ESG_FireResilience", "[年度] TEXT, [部門] TEXT, [國際指標] TEXT, [ESG領域] TEXT, [指標分類] TEXT, [預防投入/預期改善] TEXT, [指標名稱] TEXT, [實際數據呈現] TEXT, [計算公式] TEXT, [附件檔案] TEXT, [備註] TEXT" },

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
            { "FireEquip", "[日期] TEXT, [責任單位] TEXT, [柱子] TEXT, [編號] TEXT, [位置] TEXT, [具體位置] TEXT, [中繼器] TEXT, [狀態列表] TEXT, [室內消防栓] TEXT, [戶外消防栓] TEXT, [消防水帶有效期限] TEXT, [火警綜合盤] TEXT, [緊急照明燈] TEXT, [逃生指示燈_右] TEXT, [逃生指示燈_左] TEXT, [逃生指示燈_出口] TEXT, [滅火器] TEXT, [滅火器有效期限] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "FireSelfInspection", "[日期] TEXT, [檢點表名稱] TEXT, [單位] TEXT, [檢查人] TEXT, [檢查結果] TEXT, [缺失描述] TEXT, [改善對策] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "訓練時數", "[日期] TEXT, [員工姓名] TEXT, [受訓項目] TEXT, [課程名稱] TEXT, [訓練時數] TEXT, [HR外訓申請] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            
			//檢測數據
			//環境監測
			{ "EnvMonitor", "[日期] TEXT, [時間] TEXT, [SEG編號] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//廢水定申檢
            { "WastewaterPeriodic", "[日期] TEXT, [時間] TEXT, [申報季別] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//飲用水
            { "DrinkingWater", "[日期] TEXT, [時間] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//工業區
            { "IndustrialZoneTest", "[日期] TEXT, [時間] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//土壤氣
            { "SoilGasTest", "[日期] TEXT, [時間] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//廢水自主
            { "WastewaterSelfTest", "[日期] TEXT, [時間] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//循環水(廠商)
            { "CoolingWaterVendor", "[日期] TEXT, [時間] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//循環水(自評)
            { "CoolingWaterSelf", "[日期] TEXT, [時間] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//TCLP
            { "TCLP", "[日期] TEXT, [時間] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//水錶校正
            { "WaterMeterCalibration", "[日期] TEXT, [時間] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [管制值] TEXT, [管徑] TEXT, [現場流量計讀值] TEXT, [檢驗單位讀值] TEXT, [器差值] TEXT, [設定參數K值] TEXT, [校正單位] TEXT, [下次校正日期] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
			//其它檢測
            { "OtherTests", "[日期] TEXT, [時間] TEXT, [檢測點] TEXT, [檢測項目] TEXT, [檢測數據] TEXT, [管制值] TEXT, [測定方法] TEXT, [檢驗機構] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            // 檢測報告分析評估表
            { "TestReportEvaluations", "[資料庫] TEXT, [資料表] TEXT, [測定日期] TEXT, [檢測名稱] TEXT, [評估日期] TEXT, [符合度] TEXT, [測定用途] TEXT, [分析與結果說明] TEXT, [最後修改人] TEXT, [修改時間] TEXT" },
			
            { "HealthPromotion", "[日期] TEXT, [活動名稱] TEXT, [參與人數] TEXT, [執行單位] TEXT, [成果摘要] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            { "PurchaseData", "[日期] TEXT, [開單單位] TEXT, [請購單號] TEXT, [項次] TEXT, [料號] TEXT, [料名] TEXT, [規格] TEXT, [用途] TEXT, [ESG分類] TEXT, [數量] TEXT, [未稅單價] TEXT, [結案日期] TEXT, [測定用途] TEXT, [附件檔案] TEXT, [備註] TEXT" },
            
            // 個人隱藏選單
            { "WorkItems", "[日期] TEXT, [執行] TEXT, [問題] TEXT, [改善] TEXT, [追蹤] TEXT, [附件檔案] TEXT, [備註] TEXT" }
        };

        public static string[] GetDropdownList(string tableName, string columnName)
        {
            return App_DropdownManager.GetOptions(tableName, columnName);
        }

        public static string[] GetDependentDropdownList(string tableName, string columnName, string parentValue)
        {
            // 動態搜尋父層欄位名稱
            string parentColName = "";
            foreach (var kvp in App_DropdownManager.DropdownCache)
            {
                var parts = kvp.Key.Split('|');
                if (parts.Length == 4 && parts[0] == tableName && parts[1] == columnName && parts[3] == parentValue)
                {
                    parentColName = parts[2];
                    break;
                }
            }

            return App_DropdownManager.GetOptions(tableName, columnName, parentColName, parentValue) ?? new string[] { "" };
        }
    }
}
