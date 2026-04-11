private void BuildMenu()
        {
            var menuHome = new ToolStripMenuItem("頁首");
            menuHome.Click += (s, e) => LoadWelcomeScreen();

            var menuReports = new ToolStripMenuItem("報表");
            menuReports.DropDownItems.Add(CreateItem("月報表", () => new App_MonthlyReport().GetView()));
            menuReports.DropDownItems.Add(CreateItem("年報表", () => new App_YearlyReport().GetView()));

            var menuSafety = new ToolStripMenuItem("工安");
            menuSafety.DropDownItems.Add(CreateItem("工安看板", () => new App_SafetyDashboard().GetView()));
            menuSafety.DropDownItems.Add(CreateItem("虛驚事件管理", () => new App_GenericTable("Safety", "NearMiss", "虛驚事件管理").GetView()));
            menuSafety.DropDownItems.Add(CreateItem("巡檢記錄管理", () => new App_GenericTable("Safety", "SafetyInspection", "巡檢記錄管理").GetView()));
            menuSafety.DropDownItems.Add(CreateItem("安全觀察紀錄", () => new App_GenericTable("Safety", "SafetyObservation", "安全觀察紀錄").GetView()));
            menuSafety.DropDownItems.Add(CreateItem("交通意外紀錄", () => new App_GenericTable("Safety", "TrafficInjury", "交通意外紀錄").GetView()));
            menuSafety.DropDownItems.Add(CreateItem("工傷事件管理", () => new App_GenericTable("Safety", "WorkInjury", "工傷事件管理").GetView()));

            var menuNursing = new ToolStripMenuItem("護理");
            menuNursing.DropDownItems.Add(CreateItem("護理看板", () => new App_NursingDashboard().GetView()));
            menuNursing.DropDownItems.Add(CreateItem("健康促進活動", () => new App_GenericTable("Nursing", "HealthPromotion", "健康促進活動").GetView()));
            menuNursing.DropDownItems.Add(CreateItem("職災申報紀錄", () => new App_GenericTable("Nursing", "WorkInjuryReport", "職災申報紀錄").GetView()));

            var menuAir = new ToolStripMenuItem("空污");
            menuAir.DropDownItems.Add(CreateItem("空污看板", () => new App_AirDashboard().GetView()));
            menuAir.DropDownItems.Add(CreateItem("空污申報紀錄", () => new App_GenericTable("Air", "AirPollution", "空污申報紀錄").GetView()));

            var menuWater = new ToolStripMenuItem("水污");
            menuWater.DropDownItems.Add(CreateItem("水資源管理看板", () => new App_WaterDashboard().GetView()));
            // 💧 水污因為有特別客製化的欄位 (與特殊計算)，建議保留原有的獨立模組
            menuWater.DropDownItems.Add(CreateItem("【日】廢水處理水量記錄", () => new App_WaterTreatment().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】廢水處理用藥記錄", () => new App_WaterChemicals().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【日】自來水使用量", () => new App_WaterUsageDaily().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【月】納管排放數據", () => new App_DischargeData().GetView()));
            menuWater.DropDownItems.Add(CreateItem("【月】自來水用量統計", () => new App_WaterVolume().GetView()));

            var menuWaste = new ToolStripMenuItem("廢棄物");
            menuWaste.DropDownItems.Add(CreateItem("廢棄物看板", () => new App_WasteDashboard().GetView()));
            menuWaste.DropDownItems.Add(CreateItem("廢棄物統計表", () => new App_GenericTable("Waste", "WasteMonthly", "廢棄物統計表").GetView()));

            var menuFire = new ToolStripMenuItem("消防");
            menuFire.DropDownItems.Add(CreateItem("消防看板", () => new App_FireDashboard().GetView()));
            menuFire.DropDownItems.Add(CreateItem("火源責任人管理", () => new App_GenericTable("Fire", "FireResponsible", "火源責任人管理").GetView()));
            menuFire.DropDownItems.Add(CreateItem("公共危險物統計", () => new App_GenericTable("Fire", "HazardStats", "公共危險物統計").GetView()));
            menuFire.DropDownItems.Add(CreateItem("消防設備巡檢", () => new App_GenericTable("Fire", "FireEquip", "消防設備巡檢").GetView()));

            var menuTest = new ToolStripMenuItem("檢測數據");
            menuTest.DropDownItems.Add(CreateItem("檢測數據看版", () => new App_TestDashboard().GetView()));
            menuTest.DropDownItems.Add(CreateItem("環境監測", () => new App_GenericTable("TestData", "EnvMonitor", "環境監測").GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水定申檢", () => new App_GenericTable("TestData", "WastewaterPeriodic", "廢水定申檢").GetView()));
            menuTest.DropDownItems.Add(CreateItem("飲用水檢測", () => new App_GenericTable("TestData", "DrinkingWater", "飲用水檢測").GetView()));
            menuTest.DropDownItems.Add(CreateItem("工業區檢驗", () => new App_GenericTable("TestData", "IndustrialZoneTest", "工業區檢驗").GetView()));
            menuTest.DropDownItems.Add(CreateItem("土壤氣體檢測", () => new App_GenericTable("TestData", "SoilGasTest", "土壤氣體檢測").GetView()));
            menuTest.DropDownItems.Add(CreateItem("廢水自主檢驗", () => new App_GenericTable("TestData", "WastewaterSelfTest", "廢水自主檢驗").GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(廠商)", () => new App_GenericTable("TestData", "CoolingWaterVendor", "循環水檢測(廠商)").GetView()));
            menuTest.DropDownItems.Add(CreateItem("循環水檢測(自評)", () => new App_GenericTable("TestData", "CoolingWaterSelf", "循環水檢測(自評)").GetView()));
            menuTest.DropDownItems.Add(CreateItem("TCLP", () => new App_GenericTable("TestData", "TCLP", "TCLP毒性特性溶出").GetView()));
            menuTest.DropDownItems.Add(CreateItem("水錶校正", () => new App_GenericTable("TestData", "WaterMeterCalibration", "水錶校正").GetView()));
            menuTest.DropDownItems.Add(CreateItem("其它檢測數據", () => new App_GenericTable("TestData", "OtherTests", "其它檢測數據").GetView()));

            var menuEdu = new ToolStripMenuItem("教育訓練");
            menuEdu.DropDownItems.Add(CreateItem("教育訓練看板", () => new App_EduDashboard().GetView()));
            menuEdu.DropDownItems.Add(CreateItem("訓練時數", () => new App_GenericTable("教育訓練", "訓練時數", "教育訓練時數").GetView()));

            var menuLaw = new ToolStripMenuItem("法規");
            menuLaw.DropDownItems.Add(CreateItem("法規看板", () => new App_LawDashboard().GetView()));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "環保法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "職安衛法規"));
            menuLaw.DropDownItems.Add(CreateLawItem("法規", "其它法規"));

            var menuSettings = new ToolStripMenuItem("設定");
            menuSettings.DropDownItems.Add(CreateItem("操作說明", () => new App_Instruction().GetView()));
            
            var dbConfigItem = new ToolStripMenuItem("資料庫設定");
            dbConfigItem.Click += (s, e) => {
                try {
                    if (AuthManager.VerifyAdmin()) { LoadModule(new App_DbConfig().GetView()); } 
                    else { MessageBox.Show("密碼錯誤或權限不足，拒絕存取。", "授權失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                } catch (Exception ex) {
                    MessageBox.Show($"無法載入資料庫設定：\n{ex.Message}", "系統錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            menuSettings.DropDownItems.Add(dbConfigItem);

            _mainMenu.Items.AddRange(new ToolStripItem[] { 
                menuHome, menuReports, menuSafety, menuNursing, menuAir, 
                menuWater, menuWaste, menuFire, menuTest, menuEdu, menuLaw, menuSettings 
            });
        }
