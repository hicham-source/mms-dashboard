import os
import pandas as pd

REPORTS_DIR = "./reports"
SOURCE_PLAN = os.path.join(REPORTS_DIR, "Auto_Replenishment_Action_Plan.xlsx")
OUTPUT_TRACKER = os.path.join(REPORTS_DIR, "MMS_Store_Execution_Tracker.xlsx")

def build_executive_tracker():
    if not os.path.exists(SOURCE_PLAN):
        print(f"[!] Source replenishment plan not found at: {SOURCE_PLAN}")
        return

    # قراءة بيانات التوريد والمناقلات
    df_plan = pd.read_excel(SOURCE_PLAN)

    # إنشاء ورقة العمل الأولى: متابعة مناقلات وتوريد الأصناف الحرجة
    df_repl_tracker = df_plan.copy()
    df_repl_tracker['Status'] = "Pending"  # خيارات الحالة: Done / In Progress / Pending
    df_repl_tracker['Execution Date'] = ""
    df_repl_tracker['Store Receiver Sign'] = ""
    df_repl_tracker['Operations Notes'] = ""

    # إنشاء ورقة العمل الثانية: متابعة أداء المتاجر والـ Merchandising Directives
    store_tasks = [
        {"Store": "MMS Riyadh The View Mall", "Area Mgr": "Sultan", "Core Objective": "Merchandising front swap to Toys & Beauty", "Target KPI": "Lift STR% to 15%", "Deadline": "2026-10-06", "Status": "In Progress"},
        {"Store": "MMS Riyadh Solitaire", "Area Mgr": "Sultan", "Core Objective": "Maintain 100% on-shelf availability for core winners", "Target KPI": "Achieve 110% Target", "Deadline": "2026-10-08", "Status": "In Progress"},
        {"Store": "MMS Riyadh Localizer", "Area Mgr": "Sultan", "Core Objective": "Cashier upselling bundle training", "Target KPI": "UPT >= 2.3", "Deadline": "2026-10-05", "Status": "Pending"},
        {"Store": "MMS Mall of Dhahran", "Area Mgr": "Sultan", "Core Objective": "Replenish novelty impulse accessories", "Target KPI": "ATV >= 55 SAR", "Deadline": "2026-10-07", "Status": "In Progress"},
        {"Store": "MMS Jeddah Park", "Area Mgr": "Rajib", "Core Objective": "Floor-fill replenishment: reach 45k display units", "Target KPI": "WOC 6-8 Weeks", "Deadline": "2026-10-08", "Status": "Pending"},
        {"Store": "MMS Jeddah Mall of Arabia", "Area Mgr": "Rajib", "Core Objective": "Clear slow moving 3C accessories via bundle deal", "Target KPI": "STR% +10%", "Deadline": "2026-10-10", "Status": "Pending"},
        {"Store": "MMS Jeddah U-Walk", "Area Mgr": "Rajib", "Core Objective": "Immediate IST transfer dispatch to Yasmin Mall", "Target KPI": "Zero OOS on Top 20", "Deadline": "2026-10-04", "Status": "In Progress"},
        {"Store": "MMS Makkah Salam Mall", "Area Mgr": "Rajib", "Core Objective": "Rebalance gondolas towards high margin gifts", "Target KPI": "Target Ach >= 85%", "Deadline": "2026-10-09", "Status": "Pending"},
        {"Store": "MMS Madinah", "Area Mgr": "Rajib", "Core Objective": "Maintain high visual density for pilgrimage footfall", "Target KPI": "ATV >= 48 SAR", "Deadline": "2026-10-06", "Status": "In Progress"},
        {"Store": "MMS Tabuk Park", "Area Mgr": "Rajib", "Core Objective": "Receive WH priority shipment and restock toys", "Target KPI": "Zero OOS on Top 50", "Deadline": "2026-10-07", "Status": "Pending"},
        {"Store": "MMS Abha", "Area Mgr": "Rajib", "Core Objective": "Promotional feature rotation at main entrance", "Target KPI": "Txns Growth +8%", "Deadline": "2026-10-08", "Status": "Pending"},
        {"Store": "MMS Najran Park", "Area Mgr": "Rajib", "Core Objective": "Execute surplus transfer to Western branches", "Target KPI": "Dispatch 100% of plan", "Deadline": "2026-10-05", "Status": "In Progress"}
    ]
    df_store_tracker = pd.DataFrame(store_tasks)
    df_store_tracker['Actual Achieved KPI'] = ""
    df_store_tracker['Manager Feedback'] = ""

    # تصدير ملف الإكسيل بأوراق عمل متعددة وتنسيق احترافي
    with pd.ExcelWriter(OUTPUT_TRACKER, engine='openpyxl') as writer:
        df_store_tracker.to_excel(writer, sheet_name="Store_Commercial_Tasks", index=False)
        df_repl_tracker.to_excel(writer, sheet_name="Stock_Replenishment_Log", index=False)

    print(f"[✓] Operations Meeting Tracker generated successfully: {OUTPUT_TRACKER}")

if __name__ == "__main__":
    build_executive_tracker()