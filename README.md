# signflow-pro code.gs
----------------------------
/**
 * SignFlow Pro - Backend (Multi-Department, High-Res e-Signature & Analytics Engine)
 */

const CONFIG = {
  MAIN_FOLDER_NAME: "SignFlow_Storage",
  SHEET_NAME: "SignFlow_Signatures_DB",
  LOG_SHEET_NAME: "SignFlow_Audit_Log",
  DEPARTMENTS: [
    "ส่วนกลาง",
    "กองบริการ",
    "กองธุรกิจบริการ",
    "กองบำรุงรักษา",
    "กองรายได้"
  ]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SignFlow Pro - Enterprise e-Signature')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * ดึงหรือสร้างโฟลเดอร์หลักและโฟลเดอร์ย่อยประจำ 5 หน่วยงาน
 */
function getOrCreateDepartmentFolder(deptName) {
  let mainFolder;
  const mainFolders = DriveApp.getFoldersByName(CONFIG.MAIN_FOLDER_NAME);
  if (mainFolders.hasNext()) {
    mainFolder = mainFolders.next();
  } else {
    mainFolder = DriveApp.createFolder(CONFIG.MAIN_FOLDER_NAME);
  }

  if (!deptName) return mainFolder;

  const subFolders = mainFolder.getFoldersByName(deptName);
  if (subFolders.hasNext()) {
    return subFolders.next();
  }
  return mainFolder.createFolder(deptName);
}

/**
 * ดึง Spreadsheet กลางของระบบ
 */
function getSpreadsheet() {
  const files = DriveApp.getFilesByName(CONFIG.SHEET_NAME);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  } else {
    const ss = SpreadsheetApp.create(CONFIG.SHEET_NAME);
    const sheet = ss.getActiveSheet();
    sheet.setName("Signatures");
    sheet.appendRow(["ID", "Name", "Department", "FileID", "MimeType", "DataURI", "CreatedAt"]);
    sheet.setFrozenRows(1);
    return ss;
  }
}

/**
 * ดึงหรือสร้าง Sheet บันทึกลายเซ็น
 */
function getOrCreateDbSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("Signatures");
  if (!sheet) {
    sheet = ss.insertSheet("Signatures");
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Name", "Department", "FileID", "MimeType", "DataURI", "CreatedAt"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * ดึงหรือสร้าง Sheet บันทึกประวัติการใช้งาน (Audit Log)
 */
function getOrCreateLogSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sheet.appendRow(["LogID", "Timestamp", "FileName", "Department", "SignatureCount", "ActionType"]);
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["LogID", "Timestamp", "FileName", "Department", "SignatureCount", "ActionType"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 1. ดึงรายการลายเซ็น (กรองตามหน่วยงาน)
 */
function getSignaturesStock(filterDept) {
  try {
    const sheet = getOrCreateDbSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const signatures = [];
    for (let i = 1; i < data.length; i++) {
      const dept = String(data[i][2]).trim();
      if (!filterDept || dept === filterDept.trim()) {
        signatures.push({
          id: String(data[i][0]),
          name: String(data[i][1]),
          department: dept,
          fileId: String(data[i][3]),
          mimeType: String(data[i][4]),
          dataUri: String(data[i][5]),
          createdAt: String(data[i][6])
        });
      }
    }
    return signatures;
  } catch (err) {
    throw new Error("เกิดข้อผิดพลาดในการดึงข้อมูลลายเซ็น: " + err.message);
  }
}

/**
 * 2. บันทึกลายเซ็นใหม่ลงในโฟลเดอร์ของหน่วยงานนั้นๆ
 */
function uploadSignatureToStock(fileName, department, base64Data) {
  try {
    const targetFolder = getOrCreateDepartmentFolder(department);
    const splitData = base64Data.split(',');
    const contentType = splitData[0].match(/:(.*?);/)[1];
    const decodedBytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(decodedBytes, contentType, fileName);
    
    const file = targetFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const sheet = getOrCreateDbSheet();
    const sigId = "SIG-" + new Date().getTime();
    sheet.appendRow([
      sigId,
      fileName,
      department,
      file.getId(),
      contentType,
      base64Data,
      new Date().toISOString()
    ]);

    return {
      success: true,
      signature: {
        id: sigId,
        name: fileName,
        department: department,
        fileId: file.getId(),
        dataUri: base64Data
      }
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 3. ลบลายเซ็นออกจากคลัง
 */
function deleteSignatureFromStock(sigId) {
  try {
    const sheet = getOrCreateDbSheet();
    const data = sheet.getDataRange().getValues();
    let rowToDelete = -1;
    let fileIdToDelete = "";

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(sigId)) {
        rowToDelete = i + 1;
        fileIdToDelete = String(data[i][3]);
        break;
      }
    }

    if (rowToDelete !== -1) {
      sheet.deleteRow(rowToDelete);
      if (fileIdToDelete) {
        try {
          DriveApp.getFileById(fileIdToDelete).setTrashed(true);
        } catch (e) {}
      }
      return { success: true };
    }
    return { success: false, message: "ไม่พบลายเซ็นที่ต้องการลบ" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 4. บันทึกสถิติประวัติการลงนามเอกสาร (Audit Log)
 */
function logSigningAction(fileName, department, sigCount, actionType) {
  try {
    const sheet = getOrCreateLogSheet();
    const logId = "LOG-" + new Date().getTime();
    const timestamp = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([logId, timestamp, String(fileName), String(department), Number(sigCount), String(actionType)]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 5. ดึงข้อมูลสถิติรวมและแยกตาม 5 หน่วยงาน (แปลงข้อมูลเป็น String 100% ป้องกัน RPC Fail)
 */
function getSigningStatistics() {
  try {
    const sheet = getOrCreateLogSheet();
    const data = sheet.getDataRange().getValues();
    
    const deptStats = {
      "ส่วนกลาง": 0,
      "กองบริการ": 0,
      "กองธุรกิจบริการ": 0,
      "กองบำรุงรักษา": 0,
      "กองรายได้": 0
    };
    
    let totalDocs = 0;
    const recentActivities = [];

    if (data.length > 1) {
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        if (!row[0] && !row[1] && !row[2]) continue;

        totalDocs++;

        let timeStr = "";
        if (row[1] instanceof Date) {
          timeStr = Utilities.formatDate(row[1], "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");
        } else {
          timeStr = String(row[1] || "");
        }

        const fileName = String(row[2] || "เอกสาร");
        const dept = String(row[3] || "").trim();
        const sigCount = Number(row[4]) || 0;
        const actionType = String(row[5] || "ลงนาม");

        if (deptStats.hasOwnProperty(dept)) {
          deptStats[dept]++;
        }

        if (recentActivities.length < 5) {
          recentActivities.push({
            timestamp: timeStr,
            fileName: fileName,
            department: dept,
            sigCount: sigCount,
            actionType: actionType
          });
        }
      }
    }

    let sigCount = 0;
    try {
      const sigSheet = getOrCreateDbSheet();
      sigCount = Math.max(0, sigSheet.getLastRow() - 1);
    } catch (e) {
      sigCount = 0;
    }

    return {
      success: true,
      totalSignedDocs: totalDocs,
      totalSignaturesInStock: sigCount,
      departmentStats: deptStats,
      recentActivities: recentActivities
    };
  } catch (err) {
    return { 
      success: false, 
      message: "Server Error: " + err.message 
    };
  }
}
-----------------------------------------------------------------
# index.html
-----------------------------------------------------------------
<!DOCTYPE html>
<html lang="th" class="h-full">
<head>
  <meta charset="utf-8">
  <meta content="width=device-width, initial-scale=1.0" name="viewport">
  <title>SignFlow Pro - Enterprise e-Signature</title>

  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      darkMode: "class",
      theme: {
        extend: {
          colors: {
            "primary": "#004ac6",
            "primary-container": "#2563eb",
            "on-primary": "#ffffff",
            "surface": "#f8f9ff",
            "surface-container": "#e5eeff",
            "surface-container-low": "#eff4ff",
            "surface-container-lowest": "#ffffff",
            "on-surface": "#0b1c30",
            "on-surface-variant": "#434655",
            "outline": "#737686",
            "outline-variant": "#c3c6d7"
          }
        }
      }
    };
  </script>

  <!-- Google Fonts: Inter + Dancing Script + Sarabun -->
  <link href="https://fonts.googleapis.com/css2?family=Dancing+Script:wght@400;600;700&family=Inter:wght@300;400;500;600;700&family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200" rel="stylesheet">

  <!-- PDF.js CDN (Cloudflare) -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
  
  <!-- jsPDF CDN (Cloudflare) -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>

  <style>
    html, body {
      height: 100%;
      margin: 0;
      padding: 0;
      overflow: hidden;
      font-family: 'Inter', 'Sarabun', sans-serif;
    }

    .custom-scrollbar::-webkit-scrollbar {
      width: 5px;
      height: 5px;
    }
    .custom-scrollbar::-webkit-scrollbar-track {
      background: transparent;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb {
      background: #c3c6d7;
      border-radius: 9999px;
    }
    .custom-scrollbar::-webkit-scrollbar-thumb:hover {
      background: #737686;
    }

    /* กล่องลายเซ็นบนเอกสาร */
    .placed-sig-box {
      position: absolute;
      cursor: move;
      user-select: none;
      transition: border-color 0.15s, box-shadow 0.15s;
    }
    .placed-sig-box.active-box {
      border: 2px dashed #004ac6;
      background-color: rgba(0, 74, 198, 0.05);
      z-index: 40;
    }
    .placed-sig-box.confirmed-box {
      border: 2px solid #16a34a;
      background-color: rgba(22, 163, 74, 0.03);
      z-index: 30;
    }

    .sig-resize-handle {
      position: absolute;
      right: -6px;
      bottom: -6px;
      width: 14px;
      height: 14px;
      background: #004ac6;
      border: 2px solid #ffffff;
      border-radius: 9999px;
      cursor: se-resize;
      box-shadow: 0 1px 3px rgba(0,0,0,0.3);
    }
    .confirmed-box .sig-resize-handle {
      background: #16a34a;
    }
  </style>
</head>

<body class="bg-surface text-on-surface h-full w-full overflow-hidden select-none">

  <!-- ========================================== -->
  <!-- 1. LOGIN SCREEN                           -->
  <!-- ========================================== -->
  <div id="login-screen" class="fixed inset-0 z-50 flex items-center justify-center bg-surface p-4 lg:p-8">
    <div class="w-full max-w-5xl bg-surface-container-lowest rounded-2xl shadow-2xl overflow-hidden grid grid-cols-1 lg:grid-cols-12 relative border border-outline-variant/20">
      
      <!-- ฝั่งซ้าย: แบรนด์และความปลอดภัย -->
      <div class="lg:col-span-5 bg-[#00174b] p-8 lg:p-12 flex flex-col justify-between relative overflow-hidden text-white">
        <div class="absolute inset-0 bg-gradient-to-b from-[#00174b]/90 via-[#002b7a]/85 to-[#00174b]/95"></div>
        <div class="relative z-10">
          <div class="flex items-center gap-3 mb-8">
            <div class="w-10 h-10 rounded-xl bg-white/10 flex items-center justify-center backdrop-blur-md">
              <span class="material-symbols-outlined text-2xl text-blue-400">verified_user</span>
            </div>
            <span class="text-2xl font-bold tracking-wider">SignFlow Pro</span>
          </div>
          <h2 class="text-3xl font-bold mb-4">Enterprise Grade Security</h2>
          <p class="text-blue-100/80 leading-relaxed text-sm">
            ระบบลงนามเอกสารอิเล็กทรอนิกส์ระดับองค์กร แยกคลังลายเซ็นตาม 5 หน่วยงาน และระบบสถิติตรวจสอบย้อนหลัง
          </p>
        </div>
        <div class="relative z-10 mt-12 space-y-4">
          <div class="flex items-center gap-3 text-blue-200">
            <span class="material-symbols-outlined text-cyan-300">lock_reset</span>
            <span class="text-xs">256-bit SSL Encryption</span>
          </div>
          <div class="flex items-center gap-3 text-blue-200">
            <span class="material-symbols-outlined text-cyan-300">corporate_fare</span>
            <span class="text-xs">Multi-Department System</span>
          </div>
          <div class="flex items-center gap-3 text-blue-200">
            <span class="material-symbols-outlined text-cyan-300">analytics</span>
            <span class="text-xs">Real-time Usage Insights</span>
          </div>
        </div>
      </div>

      <!-- ฝั่งขวา: ฟอร์มล็อกอิน -->
      <div class="lg:col-span-7 p-8 lg:p-12 flex flex-col justify-center bg-surface-container-lowest relative">
        <div class="max-w-md w-full mx-auto">
          <div class="mb-8">
            <h1 class="text-2xl font-bold mb-2">ยินดีต้อนรับ</h1>
            <p class="text-sm text-on-surface-variant">กรุณาเข้าสู่ระบบเพื่อจัดการและลงนามเอกสารของคุณ</p>
          </div>
          <form class="space-y-6" onsubmit="handleLogin(event)">
            <div class="space-y-2">
              <label class="block text-sm font-medium">ชื่อผู้ใช้งาน</label>
              <div class="relative flex items-center">
                <span class="absolute left-4 text-outline material-symbols-outlined">person</span>
                <input required type="text" id="username" class="w-full pl-12 pr-4 py-3 bg-surface-container-low rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-primary" placeholder="ระบุชื่อผู้ใช้งาน" value="admin">
              </div>
            </div>
            <div class="space-y-2">
              <label class="block text-sm font-medium">รหัสผ่าน</label>
              <div class="relative flex items-center">
                <span class="absolute left-4 text-outline material-symbols-outlined">lock</span>
                <input required type="password" id="password" class="w-full pl-12 pr-12 py-3 bg-surface-container-low rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-primary" placeholder="••••••••" value="123456">
                <button type="button" class="absolute right-4 text-outline hover:text-on-surface" onclick="togglePassword()">
                  <span class="material-symbols-outlined" id="eye-icon">visibility</span>
                </button>
              </div>
            </div>
            <div class="flex items-center justify-between text-xs text-on-surface-variant">
              <span>หากลืมรหัสผ่าน กรุณาติดต่อ 4503</span>
            </div>
            <button type="submit" class="w-full py-3.5 bg-primary-container text-white font-medium rounded-xl hover:bg-primary transition shadow-md flex items-center justify-center gap-2">
              <span>เข้าสู่ระบบ</span>
              <span class="material-symbols-outlined text-sm">arrow_forward</span>
            </button>
          </form>
          <div class="mt-8 pt-6 text-center border-t border-outline-variant/20">
            <p class="text-xs text-outline">SignFlow Pro Enterprise Document Workflow v1.0</p>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- ========================================== -->
  <!-- 2. MAIN APPLICATION WORKSPACE              -->
  <!-- ========================================== -->
  <div id="app-screen" class="hidden h-screen w-screen flex flex-col overflow-hidden bg-surface">

    <!-- Top Header: แถบเปลี่ยนหน้าและย่อ-ขยายอยู่ตรงกลางด้านบน ไม่บังหัวเอกสาร -->
    <header class="h-16 flex-shrink-0 bg-white shadow-xs z-30 flex items-center justify-between px-6 border-b border-outline-variant/20">
      
      <!-- ซ้าย: โลโก้และชิปเอกสาร -->
      <div class="flex items-center gap-5">
        <div class="flex items-center gap-2 cursor-pointer">
          <span class="material-symbols-outlined text-primary text-2xl font-bold">signature</span>
          <span class="text-xl font-bold text-on-surface tracking-tight">SignFlow Pro</span>
        </div>
        
        <div id="header-doc-chip" class="flex items-center gap-2 bg-surface-container-low px-3.5 py-1.5 rounded-xl border border-outline-variant/20 shadow-xs max-w-[240px]">
          <span class="material-symbols-outlined text-primary text-lg flex-shrink-0">upload_file</span>
          <span class="text-xs font-semibold truncate" id="header-filename">ยังไม่ได้เลือกเอกสาร</span>
          <button onclick="document.getElementById('pdf-file-input').click()" class="p-0.5 rounded text-outline hover:text-on-surface" title="เปลี่ยนเอกสาร">
            <span class="material-symbols-outlined text-sm">swap_horiz</span>
          </button>
          <input type="file" id="pdf-file-input" accept="application/pdf" class="hidden" onchange="handlePdfUpload(event)">
        </div>
      </div>

      <!-- กลาง: แถบเปลี่ยนหน้าและย่อ-ขยายเอกสาร -->
      <div id="header-pdf-toolbar" class="flex items-center gap-2 bg-surface-container-low px-3.5 py-1.5 rounded-2xl border border-outline-variant/20 shadow-xs">
        <button onclick="changePage(-1)" class="p-1 rounded-lg hover:bg-white text-on-surface transition" title="หน้าก่อนหน้า">
          <span class="material-symbols-outlined text-base">chevron_left</span>
        </button>
        <span class="text-xs font-bold text-on-surface min-w-[75px] text-center" id="page-num-display">หน้า 0 / 0</span>
        <button onclick="changePage(1)" class="p-1 rounded-lg hover:bg-white text-on-surface transition" title="หน้าถัดไป">
          <span class="material-symbols-outlined text-base">chevron_right</span>
        </button>
        
        <div class="w-[1px] h-4 bg-outline-variant/30 mx-1"></div>
        
        <button onclick="zoomDoc(-0.1)" class="p-1 rounded-lg hover:bg-white text-on-surface transition" title="ย่อเอกสาร">
          <span class="material-symbols-outlined text-base">remove</span>
        </button>
        <span class="text-xs font-bold text-primary min-w-[45px] text-center" id="zoom-display">100%</span>
        <button onclick="zoomDoc(0.1)" class="p-1 rounded-lg hover:bg-white text-on-surface transition" title="ขยายเอกสาร">
          <span class="material-symbols-outlined text-base">add</span>
        </button>
      </div>

      <!-- ขวา: เมนูปิด-เปิดแถบข้าง & Profile -->
      <div class="flex items-center gap-3">
        <button onclick="toggleLeftSidebar()" class="p-2 rounded-xl hover:bg-surface-container text-on-surface-variant transition border border-outline-variant/20" title="ซ่อน/แสดงเมนูซ้าย">
          <span class="material-symbols-outlined text-lg">dock_to_left</span>
        </button>
        <button id="toggle-right-btn" onclick="toggleRightSidebar()" class="p-2 rounded-xl hover:bg-surface-container text-on-surface-variant transition border border-outline-variant/20" title="ซ่อน/แสดงเครื่องมือขวา">
          <span class="material-symbols-outlined text-lg">dock_to_right</span>
        </button>
        <div class="w-9 h-9 rounded-xl bg-gradient-to-tr from-primary to-primary-container flex items-center justify-center text-white font-bold shadow-sm">
          <span class="material-symbols-outlined text-lg">person</span>
        </div>
      </div>
    </header>

    <!-- Workspace Container (Fixed Height 100vh - 4rem) -->
    <div class="flex flex-1 h-[calc(100vh-4rem)] w-full overflow-hidden relative">

      <!-- LEFT SIDEBAR -->
      <aside id="left-sidebar" class="w-64 h-full bg-surface-container-low flex flex-col pt-4 pb-4 transition-all duration-300 border-r border-outline-variant/20 z-20 flex-shrink-0 overflow-hidden">
        
        <div class="flex-shrink-0 px-4 mb-2 flex items-center justify-between">
          <span class="text-xs font-bold text-outline uppercase tracking-wider">เมนูหลัก</span>
          <button onclick="toggleLeftSidebar()" class="p-1 rounded-full hover:bg-surface-container text-outline">
            <span class="material-symbols-outlined text-sm">chevron_left</span>
          </button>
        </div>

        <nav class="flex-shrink-0 px-3 space-y-1">
          <button onclick="switchView('signing')" id="menu-btn-signing" 
                  class="w-full flex items-center px-3 py-2 rounded-lg bg-primary-container text-white font-semibold shadow-sm text-sm transition">
            <span class="material-symbols-outlined mr-3 text-lg">edit_document</span>ลงนามเอกสาร
          </button>
          
          <button onclick="switchView('vault')" id="menu-btn-vault" 
                  class="w-full flex items-center px-3 py-2 rounded-lg text-on-surface-variant hover:bg-surface-container transition text-sm">
            <span class="material-symbols-outlined mr-3 text-lg">draw</span>จัดการคลังลายเซ็น
          </button>

          <button onclick="switchView('analytics')" id="menu-btn-analytics" 
                  class="w-full flex items-center px-3 py-2 rounded-lg text-on-surface-variant hover:bg-surface-container transition text-sm">
            <span class="material-symbols-outlined mr-3 text-lg">analytics</span>สถิติการใช้งาน
          </button>
        </nav>

        <div id="thumbnails-section" class="flex-1 flex flex-col min-h-0 px-4 mt-3 border-t border-outline-variant/10 pt-3">
          <div class="flex-shrink-0 text-[11px] font-bold text-outline uppercase tracking-wider mb-2" id="thumbnail-title">
            พรีวิวหน้าเอกสาร (0)
          </div>
          <div id="thumbnails-list" class="flex-1 overflow-y-auto custom-scrollbar space-y-3 pr-1 pb-2">
            <div class="text-xs text-outline text-center py-8">กรุณาเลือกไฟล์ PDF เพื่อดูพรีวิว</div>
          </div>
        </div>
      </aside>

      <!-- ============================================== -->
      <!-- VIEW 1: หน้าลงนามเอกสาร (SIGNING WORKSPACE)      -->
      <!-- ============================================== -->
      <div id="view-signing" class="flex-1 min-w-0 w-full h-full flex overflow-hidden">
        
        <!-- CENTER: PDF CANVAS VIEWPORT (Scroll ถึงหัวกระดาษได้สมบูรณ์) -->
        <main class="flex-1 min-w-0 h-full bg-[#151921] flex flex-col items-center justify-center relative overflow-hidden shadow-inner">
          <div class="flex-1 w-full h-full overflow-auto custom-scrollbar flex items-start justify-center p-8" id="viewport-container">
            <div id="pdf-wrapper" class="relative bg-white shadow-2xl rounded-lg select-none border border-outline-variant/10 transition-all my-auto">
              <canvas id="pdf-canvas" class="block"></canvas>
              <div id="signature-boxes-container" class="absolute inset-0 pointer-events-none"></div>
            </div>
          </div>
        </main>

        <!-- RIGHT SIDEBAR: แถบเครื่องมือวางลายเซ็น -->
        <aside id="right-sidebar" class="w-80 h-full bg-surface-container-low/95 backdrop-blur-md flex flex-col justify-between border-l border-outline-variant/20 z-20 flex-shrink-0 overflow-hidden transition-all duration-300">
          
          <div class="p-4 space-y-4 overflow-y-auto custom-scrollbar flex-1">
            
            <!-- แหล่งที่มาของเอกสาร -->
            <div class="space-y-1.5">
              <span class="text-[11px] uppercase tracking-wider text-outline font-bold">แหล่งที่มาของเอกสาร</span>
              <div class="p-3 rounded-2xl bg-white shadow-xs border border-outline-variant/20 flex items-center justify-between">
                <div class="flex items-center gap-3 min-w-0">
                  <div class="w-10 h-10 rounded-xl bg-primary/10 flex items-center justify-center text-primary flex-shrink-0">
                    <span class="material-symbols-outlined">description</span>
                  </div>
                  <div class="min-w-0">
                    <p class="text-xs font-bold text-on-surface truncate" id="side-filename">ยังไม่ได้เลือกเอกสาร</p>
                    <p class="text-[10px] text-outline" id="side-fileinfo">- MB • 0 หน้า</p>
                  </div>
                </div>
                <button onclick="document.getElementById('pdf-file-input').click()" class="p-1 rounded-lg text-outline hover:text-on-surface hover:bg-surface-container transition-colors" title="เปลี่ยนเอกสาร">
                  <span class="material-symbols-outlined text-sm">swap_horiz</span>
                </button>
              </div>
            </div>

            <!-- ดรอปดาวน์เลือก 5 หน่วยงาน -->
            <div class="p-3.5 bg-white rounded-2xl border border-outline-variant/30 shadow-xs space-y-2">
              <label class="block text-[11px] uppercase tracking-wider text-outline font-bold flex items-center gap-1.5">
                <span class="material-symbols-outlined text-primary text-[18px]">corporate_fare</span>
                <span>เลือกหน่วยงานเพื่อพรีวิว</span>
              </label>
              <select id="signing-department-select" onchange="onDepartmentFilterChanged(this.value)"
                      class="w-full p-2 bg-surface-container-low rounded-xl text-xs font-semibold text-on-surface border border-outline-variant/20 focus:outline-none focus:ring-2 focus:ring-primary">
                <option value="ส่วนกลาง">ส่วนกลาง</option>
                <option value="กองบริการ">กองบริการ</option>
                <option value="กองธุรกิจบริการ">กองธุรกิจบริการ</option>
                <option value="กองบำรุงรักษา">กองบำรุงรักษา</option>
                <option value="กองรายได้">กองรายได้</option>
              </select>
            </div>

            <!-- พรีวิวลายเซ็นประจำหน่วยงาน -->
            <div class="space-y-2">
              <div class="flex items-center justify-between">
                <span class="text-[11px] uppercase tracking-wider text-outline font-bold">ลายมือชื่อประจำหน่วย</span>
                <span class="text-[10px] text-primary cursor-pointer hover:underline" onclick="switchView('vault')">+ เพิ่มในคลัง</span>
              </div>
              
              <div id="department-signature-grid" class="grid grid-cols-2 gap-2 min-h-[70px]">
                <div class="col-span-2 text-center py-4 text-xs text-outline">กำลังโหลดลายเซ็น...</div>
              </div>
            </div>

            <!-- เมนูกล่องพิมพ์ข้อความ (Collapsible Accordion Dropdown) -->
            <div class="bg-white rounded-2xl border border-outline-variant/30 shadow-xs overflow-hidden transition-all">
              <button type="button" onclick="toggleTextCustomizerAccordion()" 
                      class="w-full p-3.5 flex items-center justify-between bg-white hover:bg-surface-container-low transition">
                <div class="flex items-center gap-2">
                  <div class="w-7 h-7 rounded-lg bg-primary/10 text-primary flex items-center justify-center">
                    <span class="material-symbols-outlined text-[18px]">text_fields</span>
                  </div>
                  <span class="text-xs font-bold text-on-surface">พิมพ์ข้อความ / วันที่</span>
                </div>
                <span id="accordion-icon" class="material-symbols-outlined text-outline transition-transform duration-200">expand_less</span>
              </button>

              <div id="text-customizer-content" class="p-3.5 pt-0 space-y-3 block">
                <div>
                  <textarea id="custom-text-input" rows="2" 
                    oninput="autoResizeTextarea(this)"
                    class="w-full p-2.5 bg-surface-container-low rounded-xl text-xs text-on-surface border border-outline-variant/20 focus:outline-none focus:ring-2 focus:ring-primary transition resize-y min-h-[50px] max-h-[140px]"
                    placeholder="พิมพ์ชื่อ, ตำแหน่ง หรือข้อความ..."></textarea>
                </div>

                <div class="grid grid-cols-2 gap-2">
                  <div>
                    <label class="block text-[10px] text-outline font-medium mb-1">ฟอนต์</label>
                    <select id="font-family-select" class="w-full p-1.5 bg-surface-container-low rounded-lg text-[11px] text-on-surface border border-outline-variant/20 focus:outline-none">
                      <option value="'Dancing Script', cursive">ลายมือ (Script)</option>
                      <option value="'Sarabun', sans-serif">ราชการ (Sarabun)</option>
                      <option value="'Inter', sans-serif">โมเดิร์น (Inter)</option>
                      <option value="'Courier New', monospace">พิมพ์ดีด (Monospace)</option>
                    </select>
                  </div>
                  <div>
                    <label class="block text-[10px] text-outline font-medium mb-1">ขนาด</label>
                    <select id="font-size-select" class="w-full p-1.5 bg-surface-container-low rounded-lg text-[11px] text-on-surface border border-outline-variant/20 focus:outline-none">
                      <option value="18">เล็ก (18px)</option>
                      <option value="24" selected>ปานกลาง (24px)</option>
                      <option value="32">ใหญ่ (32px)</option>
                    </select>
                  </div>
                </div>

                <div>
                  <label class="block text-[10px] text-outline font-medium mb-1">ความหนาตัวอักษร</label>
                  <div class="grid grid-cols-3 gap-1 bg-surface-container-low p-1 rounded-lg border border-outline-variant/20">
                    <button type="button" onclick="setFontWeight('300')" id="weight-btn-300" class="py-1 text-[11px] rounded text-outline hover:text-on-surface transition font-light">บาง</button>
                    <button type="button" onclick="setFontWeight('400')" id="weight-btn-400" class="py-1 text-[11px] rounded text-outline hover:text-on-surface transition font-normal">ปกติ</button>
                    <button type="button" onclick="setFontWeight('700')" id="weight-btn-700" class="py-1 text-[11px] rounded bg-white text-primary font-bold shadow-xs transition">หนา</button>
                  </div>
                </div>

                <div class="flex items-center justify-between pt-1">
                  <div class="flex items-center bg-surface-container-low rounded-lg p-0.5 border border-outline-variant/20">
                    <button type="button" onclick="setTextAlign('left')" id="align-btn-left" class="p-1 rounded text-outline hover:text-on-surface transition" title="ชิดซ้าย">
                      <span class="material-symbols-outlined text-[16px]">format_align_left</span>
                    </button>
                    <button type="button" onclick="setTextAlign('center')" id="align-btn-center" class="p-1 rounded bg-white text-primary shadow-xs transition" title="ตรงกลาง">
                      <span class="material-symbols-outlined text-[16px]">format_align_center</span>
                    </button>
                    <button type="button" onclick="setTextAlign('right')" id="align-btn-right" class="p-1 rounded text-outline hover:text-on-surface transition" title="ชิดขวา">
                      <span class="material-symbols-outlined text-[16px]">format_align_right</span>
                    </button>
                  </div>

                  <div class="flex items-center gap-1.5">
                    <label class="text-[10px] text-outline font-medium">สี:</label>
                    <input type="color" id="text-color-picker" value="#004ac6" class="w-6 h-6 rounded cursor-pointer border border-outline-variant/30">
                  </div>
                </div>

                <button type="button" onclick="generateAndPlaceTextSignature()" 
                        class="w-full py-2 bg-primary/10 hover:bg-primary/20 text-primary rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 transition active:scale-[0.98]">
                  <span class="material-symbols-outlined text-[16px]">add_location_alt</span>
                  <span>วางข้อความนี้ลงเอกสาร</span>
                </button>
              </div>
            </div>

            <!-- สถานะจุดวาง -->
            <div class="flex items-center justify-between pt-1">
              <span class="text-[11px] text-primary font-bold" id="placed-count-display">วางแล้ว 0 จุด</span>
              <button onclick="clearAllSignaturesOnPage()" class="text-[11px] text-red-600 hover:underline flex items-center gap-0.5">
                <span class="material-symbols-outlined text-[14px]">delete_sweep</span> ลบทั้งหมดในหน้านี้
              </button>
            </div>

          </div>

          <!-- ปุ่มส่งออกเอกสาร ตรึงล่างสุด -->
          <div class="flex-shrink-0 p-4 bg-surface-container border-t border-outline-variant/10 space-y-2">
            <button onclick="downloadSignedPdfDocument()" class="w-full h-11 bg-primary text-white rounded-xl text-sm font-medium flex items-center justify-center gap-2 shadow-md hover:bg-primary/90 transition active:scale-[0.98]">
              <span class="material-symbols-outlined text-lg">picture_as_pdf</span>
              <span>ดาวน์โหลดเอกสาร PDF (HD)</span>
            </button>
            
            <button onclick="printSignedDocumentDirectly()" class="w-full h-11 bg-white text-on-surface border border-outline-variant/30 rounded-xl text-sm font-medium flex items-center justify-center gap-2 hover:bg-surface-container-high transition active:scale-[0.98]">
              <span class="material-symbols-outlined text-lg text-primary">print</span>
              <span>สั่งพิมพ์เอกสาร</span>
            </button>
          </div>
        </aside>
      </div>

      <!-- ============================================== -->
      <!-- VIEW 2: หน้าจัดการคลังลายเซ็น (SIGNATURE VAULT) -->
      <!-- ============================================== -->
      <div id="view-vault" class="hidden flex-1 min-w-0 w-full h-full bg-surface overflow-y-auto custom-scrollbar p-6 lg:p-8">
        <div class="max-w-6xl mx-auto space-y-6">
          
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-outline-variant/20 pb-4">
            <div>
              <h2 class="text-2xl font-bold text-on-surface">จัดการคลังลายเซ็นแยกตามหน่วยงาน</h2>
              <p class="text-xs text-on-surface-variant mt-1">คลิกที่ภาพลายเซ็นเพื่อขยายดูขนาดใหญ่ ตรวจสอบลายเส้นและความซ้ำซ้อน</p>
            </div>
            <button onclick="openUploadModal()" class="px-4 py-2.5 bg-primary text-white rounded-xl text-sm font-medium flex items-center gap-2 shadow hover:bg-primary/90 transition">
              <span class="material-symbols-outlined text-lg">add_photo_alternate</span>
              <span>+ อัปโหลดลายเซ็นใหม่</span>
            </button>
          </div>

          <div class="flex items-center gap-2 overflow-x-auto pb-2 border-b border-outline-variant/10" id="vault-tabs-bar"></div>
          <div id="vault-signatures-container" class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-4"></div>
        </div>
      </div>

      <!-- ============================================== -->
      <!-- VIEW 3: หน้าสถิติการใช้งาน (ANALYTICS)          -->
      <!-- ============================================== -->
      <div id="view-analytics" class="hidden flex-1 min-w-0 w-full h-full bg-surface overflow-y-auto custom-scrollbar p-6 lg:p-8">
        <div class="max-w-5xl mx-auto space-y-6">
          <div class="border-b border-outline-variant/20 pb-4 flex items-center justify-between">
            <div>
              <h2 class="text-2xl font-bold text-on-surface">สถิติการลงนามเอกสาร</h2>
              <p class="text-xs text-on-surface-variant mt-1">สรุปข้อมูลการลงนามเอกสารจริงและแยกตาม 5 หน่วยงาน</p>
            </div>
            <button onclick="updateAnalyticsView()" class="px-3 py-1.5 bg-white border border-outline-variant/30 rounded-xl text-xs font-semibold text-primary hover:bg-surface-container flex items-center gap-1">
              <span class="material-symbols-outlined text-sm">refresh</span> รีเฟรชข้อมูล
            </button>
          </div>

          <!-- การ์ดภาพรวม 3 ค่า -->
          <div class="grid grid-cols-1 sm:grid-cols-3 gap-4">
            <div class="p-5 bg-white rounded-2xl border border-outline-variant/20 shadow-xs space-y-1">
              <span class="text-xs text-outline font-medium">เอกสารที่ลงนามสำเร็จทั้งหมด</span>
              <p class="text-3xl font-extrabold text-primary" id="stat-signed-count">0 ฉบับ</p>
            </div>
            <div class="p-5 bg-white rounded-2xl border border-outline-variant/20 shadow-xs space-y-1">
              <span class="text-xs text-outline font-medium">ลายเซ็นทั้งหมดในคลัง (5 หน่วยงาน)</span>
              <p class="text-3xl font-extrabold text-on-surface" id="stat-total-sigs">0 รายการ</p>
            </div>
            <div class="p-5 bg-white rounded-2xl border border-outline-variant/20 shadow-xs space-y-1">
              <span class="text-xs text-outline font-medium">สถานะฐานข้อมูล</span>
              <p class="text-xl font-bold text-green-600 flex items-center gap-1 mt-1">
                <span class="material-symbols-outlined text-green-500">check_circle</span> เชื่อมต่อสมบูรณ์
              </p>
            </div>
          </div>

          <!-- ยอดลงนามแยก 5 หน่วยงาน -->
          <div class="p-5 bg-white rounded-2xl border border-outline-variant/20 shadow-xs space-y-4">
            <h3 class="text-sm font-bold text-on-surface flex items-center gap-2">
              <span class="material-symbols-outlined text-primary text-lg">corporate_fare</span>
              <span>จำนวนการลงนามเอกสารแยกตามหน่วยงาน</span>
            </h3>
            <div class="grid grid-cols-1 sm:grid-cols-5 gap-3" id="analytics-dept-breakdown"></div>
          </div>

          <!-- รายการประวัติล่าสุด -->
          <div class="p-5 bg-white rounded-2xl border border-outline-variant/20 shadow-xs space-y-3">
            <h3 class="text-sm font-bold text-on-surface flex items-center gap-2">
              <span class="material-symbols-outlined text-primary text-lg">history</span>
              <span>ประวัติการลงนามล่าสุด (Recent Activities)</span>
            </h3>
            <div id="analytics-recent-list" class="divide-y divide-outline-variant/10 text-xs text-on-surface">
              <div class="py-4 text-center text-outline">ยังไม่มีประวัติการลงนาม</div>
            </div>
          </div>
        </div>
      </div>

    </div>
  </div>

  <!-- Modal สำหรับขยายดูภาพลายเซ็นขนาดใหญ่ (Image Lightbox Modal) -->
  <div id="lightbox-modal" class="hidden fixed inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-center justify-center p-4" onclick="closeLightboxModal()">
    <div class="bg-white rounded-2xl p-6 max-w-lg w-full shadow-2xl border border-outline-variant/20 space-y-4 relative" onclick="event.stopPropagation()">
      <div class="flex items-center justify-between border-b border-outline-variant/10 pb-3">
        <div>
          <h3 class="text-base font-bold text-on-surface" id="lightbox-title">ดูลายเซ็น</h3>
          <p class="text-[11px] text-outline" id="lightbox-dept">-</p>
        </div>
        <button onclick="closeLightboxModal()" class="text-outline hover:text-on-surface p-1 rounded-lg">
          <span class="material-symbols-outlined text-xl">close</span>
        </button>
      </div>

      <div class="h-64 flex items-center justify-center p-4 bg-[#f8fafc] rounded-xl border border-outline-variant/20 overflow-hidden">
        <img id="lightbox-img" src="" class="max-h-full max-w-full object-contain filter drop-shadow-md">
      </div>

      <div class="flex justify-end pt-1">
        <button onclick="closeLightboxModal()" class="px-4 py-2 bg-surface-container text-on-surface rounded-xl text-xs font-semibold hover:bg-surface-container-high transition">
          ปิดหน้าต่าง
        </button>
      </div>
    </div>
  </div>

  <!-- Modal สำหรับอัปโหลดลายเซ็นใหม่ -->
  <div id="upload-sig-modal" class="hidden fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
    <div class="bg-white rounded-2xl p-6 max-w-md w-full shadow-2xl border border-outline-variant/20 space-y-4">
      <div class="flex items-center justify-between border-b border-outline-variant/10 pb-3">
        <h3 class="text-base font-bold text-on-surface">อัปโหลดลายเซ็นใหม่ลงคลัง</h3>
        <button onclick="closeUploadModal()" class="text-outline hover:text-on-surface">
          <span class="material-symbols-outlined text-xl">close</span>
        </button>
      </div>

      <div class="space-y-3">
        <div>
          <label class="block text-xs font-semibold text-on-surface mb-1">เลือกหน่วยงานสังกัด *</label>
          <select id="upload-dept-select" class="w-full p-2.5 bg-surface-container-low rounded-xl text-xs border border-outline-variant/20 focus:outline-none focus:ring-2 focus:ring-primary font-medium">
            <option value="ส่วนกลาง">ส่วนกลาง</option>
            <option value="กองบริการ">กองบริการ</option>
            <option value="กองธุรกิจบริการ">กองธุรกิจบริการ</option>
            <option value="กองบำรุงรักษา">กองบำรุงรักษา</option>
            <option value="กองรายได้">กองรายได้</option>
          </select>
        </div>

        <div>
          <label class="block text-xs font-semibold text-on-surface mb-1">ชื่อเจ้าของลายเซ็น / ตำแหน่ง *</label>
          <input type="text" id="upload-owner-name" class="w-full p-2.5 bg-surface-container-low rounded-xl text-xs border border-outline-variant/20 focus:outline-none focus:ring-2 focus:ring-primary" placeholder="เช่น นายสมชาย ใจดี (ผอ.กอง)">
        </div>

        <div>
          <label class="block text-xs font-semibold text-on-surface mb-1">ไฟล์ภาพลายเซ็น (PNG พื้นใส, JPG, WebP) *</label>
          <input type="file" id="upload-file-input" accept="image/png,image/jpeg,image/webp" 
                 class="w-full text-xs text-outline file:mr-3 file:py-2 file:px-3 file:rounded-xl file:border-0 file:text-xs file:font-semibold file:bg-primary-container file:text-white hover:file:bg-primary cursor-pointer">
        </div>
      </div>

      <div class="flex items-center justify-end gap-2 pt-2 border-t border-outline-variant/10">
        <button onclick="closeUploadModal()" class="px-4 py-2 text-xs font-semibold text-outline hover:text-on-surface">ยกเลิก</button>
        <button onclick="submitUploadSignature()" class="px-5 py-2.5 bg-primary text-white rounded-xl text-xs font-semibold shadow hover:bg-primary/90">บันทึกลงคลัง</button>
      </div>
    </div>
  </div>

  <!-- Loading Screen Overlay -->
  <div id="loading-overlay" class="hidden fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center flex-col text-white gap-3">
    <div class="w-10 h-10 border-4 border-white/20 border-t-white rounded-full animate-spin"></div>
    <span class="text-sm font-medium" id="loading-text">กำลังประมวลผล...</span>
  </div>

  <!-- ========================================== -->
  <!-- 3. JAVASCRIPT LOGIC                        -->
  <!-- ========================================== -->
  <script>
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    const DEPARTMENTS = ["ส่วนกลาง", "กองบริการ", "กองธุรกิจบริการ", "กองบำรุงรักษา", "กองรายได้"];

    let pdfDoc = null;
    let pdfBytes = null;
    let currentPageNum = 1;
    let totalPages = 0;
    let currentZoom = 1.0;
    let selectedDepartment = "ส่วนกลาง";
    let selectedVaultTab = "ส่วนกลาง";
    let allSignaturesCache = {};
    let pdfFileName = "Document.pdf";
    let placedSignatures = [];

    let currentTextAlign = 'center';
    let currentFontWeight = '700';

    function togglePassword() {
      const p = document.getElementById('password');
      const icon = document.getElementById('eye-icon');
      if (p.type === 'password') {
        p.type = 'text';
        icon.textContent = 'visibility_off';
      } else {
        p.type = 'password';
        icon.textContent = 'visibility';
      }
    }

    function handleLogin(e) {
      e.preventDefault();
      const username = document.getElementById('username').value.trim();
      const password = document.getElementById('password').value;

      const CORRECT_USER = "admin";
      const CORRECT_PASS = "123456";

      if (username === CORRECT_USER && password === CORRECT_PASS) {
        document.getElementById('login-screen').classList.add('hidden');
        document.getElementById('app-screen').classList.remove('hidden');
        initVaultTabs();
        loadDepartmentSignatures(selectedDepartment);
      } else {
        alert("ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง!");
      }
    }

    function showLoading(msg) {
      document.getElementById('loading-text').textContent = msg;
      document.getElementById('loading-overlay').classList.remove('hidden');
    }

    function hideLoading() {
      document.getElementById('loading-overlay').classList.add('hidden');
    }

    function switchView(viewName) {
      const views = ['signing', 'vault', 'analytics'];
      views.forEach(v => {
        document.getElementById(`view-${v}`).classList.add('hidden');
        const btn = document.getElementById(`menu-btn-${v}`);
        btn.className = "w-full flex items-center px-3 py-2 rounded-lg text-on-surface-variant hover:bg-surface-container transition text-sm";
      });

      document.getElementById(`view-${viewName}`).classList.remove('hidden');
      const activeBtn = document.getElementById(`menu-btn-${viewName}`);
      activeBtn.className = "w-full flex items-center px-3 py-2 rounded-lg bg-primary-container text-white font-semibold shadow-sm text-sm transition";

      const docChip = document.getElementById('header-doc-chip');
      const pdfToolbar = document.getElementById('header-pdf-toolbar');
      const toggleRightBtn = document.getElementById('toggle-right-btn');
      const thumbSection = document.getElementById('thumbnails-section');

      if (viewName === 'signing') {
        docChip.classList.remove('hidden');
        pdfToolbar.classList.remove('hidden');
        toggleRightBtn.classList.remove('hidden');
        thumbSection.classList.remove('hidden');
        if (pdfDoc) renderPage(currentPageNum);
      } else {
        docChip.classList.add('hidden');
        pdfToolbar.classList.add('hidden');
        toggleRightBtn.classList.add('hidden');
        thumbSection.classList.add('hidden');
        if (viewName === 'vault') {
          renderVaultSignatures();
        } else if (viewName === 'analytics') {
          updateAnalyticsView();
        }
      }
    }

    let leftOpen = true;
    let rightOpen = true;

    function toggleLeftSidebar() {
      const bar = document.getElementById('left-sidebar');
      leftOpen = !leftOpen;
      if (leftOpen) {
        bar.style.width = '16rem';
        bar.style.display = 'flex';
      } else {
        bar.style.width = '0px';
        bar.style.display = 'none';
      }
    }

    function toggleRightSidebar() {
      const bar = document.getElementById('right-sidebar');
      rightOpen = !rightOpen;
      if (rightOpen) {
        bar.style.width = '20rem';
        bar.style.display = 'flex';
      } else {
        bar.style.width = '0px';
        bar.style.display = 'none';
      }
    }

    // --- 1. ระบบจัดการคลังลายเซ็นแยก 5 หน่วยงาน ---
    function onDepartmentFilterChanged(dept) {
      selectedDepartment = dept;
      loadDepartmentSignatures(dept);
    }

    function loadDepartmentSignatures(dept) {
      const grid = document.getElementById('department-signature-grid');
      grid.innerHTML = '<div class="col-span-2 text-center py-4 text-xs text-outline">กำลังโหลดลายเซ็น...</div>';

      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler(list => {
            allSignaturesCache[dept] = list || [];
            renderDepartmentGrid(allSignaturesCache[dept]);
          })
          .withFailureHandler(err => {
            grid.innerHTML = `<div class="col-span-2 text-center text-xs text-red-500 py-2">${err.message}</div>`;
          })
          .getSignaturesStock(dept);
      } else {
        if (!allSignaturesCache[dept]) {
          allSignaturesCache[dept] = [
            { id: '1', name: `ลายเซ็น ผอ. (${dept})`, department: dept, dataUri: createMockSignatureCanvas(`ผอ. ${dept}`) }
          ];
        }
        renderDepartmentGrid(allSignaturesCache[dept]);
      }
    }

    function createMockSignatureCanvas(text) {
      const c = document.createElement('canvas');
      c.width = 200; c.height = 80;
      const ctx = c.getContext('2d');
      ctx.font = 'bold 26px "Dancing Script", cursive';
      ctx.fillStyle = '#004ac6';
      ctx.fillText(text, 15, 50);
      return c.toDataURL('image/png');
    }

    function renderDepartmentGrid(list) {
      const grid = document.getElementById('department-signature-grid');
      grid.innerHTML = '';

      if (!list || list.length === 0) {
        grid.innerHTML = '<div class="col-span-2 text-center text-xs text-outline py-4">ไม่มีลายเซ็นในหน่วยงานนี้</div>';
        return;
      }

      list.forEach(sig => {
        const item = document.createElement('div');
        item.className = "p-2 rounded-2xl bg-white border border-outline-variant/30 hover:border-primary relative cursor-pointer shadow-xs transition group";
        item.innerHTML = `
          <div class="h-12 flex items-center justify-center select-none overflow-hidden" onclick="placeSignatureFromStock('${sig.dataUri}')">
            <img src="${sig.dataUri}" class="max-h-full max-w-full object-contain">
          </div>
          <span class="block text-[9px] text-center text-on-surface-variant font-medium truncate mt-1">${sig.name}</span>
        `;
        grid.appendChild(item);
      });
    }

    function placeSignatureFromStock(dataUri) {
      if (!pdfDoc) {
        alert("กรุณาอัปโหลดเอกสาร PDF ก่อนวางลายเซ็น");
        return;
      }
      const offset = (placedSignatures.filter(s => s.page === currentPageNum).length * 25) % 150;
      const newSig = {
        id: 'sig_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4),
        page: currentPageNum,
        x: 80 + offset,
        y: 120 + offset,
        width: 170,
        height: 65,
        dataUri: dataUri,
        confirmed: false
      };

      placedSignatures.push(newSig);
      updatePlacedCountDisplay();
      renderSignatureBoxesOnPage();
    }

    function initVaultTabs() {
      const bar = document.getElementById('vault-tabs-bar');
      bar.innerHTML = '';
      DEPARTMENTS.forEach(dept => {
        const btn = document.createElement('button');
        btn.id = `vault-tab-${dept}`;
        btn.className = `px-4 py-2 rounded-xl text-xs font-bold whitespace-nowrap transition ${
          dept === selectedVaultTab ? 'bg-primary text-white shadow-sm' : 'text-on-surface-variant hover:bg-surface-container'
        }`;
        btn.textContent = dept;
        btn.onclick = () => {
          selectedVaultTab = dept;
          initVaultTabs();
          renderVaultSignatures();
        };
        bar.appendChild(btn);
      });
    }

    function renderVaultSignatures() {
      const container = document.getElementById('vault-signatures-container');
      container.innerHTML = '<div class="col-span-full text-center py-8 text-xs text-outline">กำลังโหลดลายเซ็น...</div>';

      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler(list => {
            allSignaturesCache[selectedVaultTab] = list || [];
            displayVaultCards(allSignaturesCache[selectedVaultTab]);
          })
          .getSignaturesStock(selectedVaultTab);
      } else {
        if (!allSignaturesCache[selectedVaultTab]) {
          allSignaturesCache[selectedVaultTab] = [
            { id: '1', name: `ลายเซ็น ผอ. (${selectedVaultTab})`, department: selectedVaultTab, dataUri: createMockSignatureCanvas(`ผอ. ${selectedVaultTab}`) }
          ];
        }
        displayVaultCards(allSignaturesCache[selectedVaultTab]);
      }
    }

    function displayVaultCards(list) {
      const container = document.getElementById('vault-signatures-container');
      container.innerHTML = '';

      if (!list || list.length === 0) {
        container.innerHTML = `
          <div class="col-span-full text-center py-12 text-outline bg-white rounded-2xl border border-dashed border-outline-variant/30">
            <span class="material-symbols-outlined text-4xl mb-1 text-outline">draw</span>
            <p class="text-xs">ยังไม่มีลายเซ็นในหมวด ${selectedVaultTab}</p>
            <button onclick="openUploadModal()" class="mt-3 text-xs text-primary font-bold hover:underline">+ เพิ่มลายเซ็นแรก</button>
          </div>
        `;
        return;
      }

      list.forEach(sig => {
        const card = document.createElement('div');
        card.className = "p-4 bg-white rounded-2xl border border-outline-variant/30 shadow-xs flex flex-col justify-between hover:border-primary transition group relative";
        
        // คลิกเพื่อขยายดูภาพใหญ่
        card.innerHTML = `
          <div class="h-24 flex items-center justify-center p-2 bg-surface-container-low rounded-xl mb-3 cursor-zoom-in relative group/img" 
               onclick="openLightboxModal('${sig.name}', '${sig.department}', '${sig.dataUri}')" title="คลิกเพื่อขยายดูลายเซ็น">
            <img src="${sig.dataUri}" class="max-h-full max-w-full object-contain">
            <span class="absolute bottom-1 right-1 p-1 bg-white/80 rounded-md text-outline opacity-0 group-hover/img:opacity-100 transition">
              <span class="material-symbols-outlined text-xs">zoom_in</span>
            </span>
          </div>
          <div>
            <p class="text-xs font-bold text-on-surface truncate">${sig.name}</p>
            <span class="text-[10px] text-outline font-medium">สังกัด: ${sig.department}</span>
          </div>
          <button onclick="deleteSignatureFromVault('${sig.id}')" title="ลบลายเซ็นนี้"
                  class="absolute top-2 right-2 p-1.5 rounded-full bg-red-50 text-red-500 opacity-0 group-hover:opacity-100 hover:bg-red-500 hover:text-white transition shadow-sm">
            <span class="material-symbols-outlined text-[16px]">delete</span>
          </button>
        `;
        container.appendChild(card);
      });
    }

    function openLightboxModal(name, dept, dataUri) {
      document.getElementById('lightbox-title').textContent = name;
      document.getElementById('lightbox-dept').textContent = `สังกัด: ${dept}`;
      document.getElementById('lightbox-img').src = dataUri;
      document.getElementById('lightbox-modal').classList.remove('hidden');
    }

    function closeLightboxModal() {
      document.getElementById('lightbox-modal').classList.add('hidden');
    }

    function deleteSignatureFromVault(sigId) {
      if (!confirm("คุณต้องการลบลายเซ็นนี้ออกจากคลังถาวรใช่หรือไม่?")) return;
      showLoading("กำลังลบลายเซ็น...");

      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler(res => {
            hideLoading();
            if (res.success) {
              renderVaultSignatures();
              if (selectedDepartment === selectedVaultTab) {
                loadDepartmentSignatures(selectedDepartment);
              }
            } else {
              alert("ลบไม่สำเร็จ: " + res.message);
            }
          })
          .withFailureHandler(err => {
            hideLoading();
            alert(err.message);
          })
          .deleteSignatureFromStock(sigId);
      } else {
        setTimeout(() => {
          hideLoading();
          allSignaturesCache[selectedVaultTab] = allSignaturesCache[selectedVaultTab].filter(s => s.id !== sigId);
          renderVaultSignatures();
        }, 400);
      }
    }

    function openUploadModal() {
      document.getElementById('upload-dept-select').value = selectedVaultTab;
      document.getElementById('upload-sig-modal').classList.remove('hidden');
    }

    function closeUploadModal() {
      document.getElementById('upload-sig-modal').classList.add('hidden');
      document.getElementById('upload-owner-name').value = '';
      document.getElementById('upload-file-input').value = '';
    }

    function submitUploadSignature() {
      const dept = document.getElementById('upload-dept-select').value;
      const ownerName = document.getElementById('upload-owner-name').value.trim();
      const fileInput = document.getElementById('upload-file-input');

      if (!ownerName) {
        alert("กรุณาระบุชื่อเจ้าของลายเซ็น");
        return;
      }
      if (!fileInput.files || fileInput.files.length === 0) {
        alert("กรุณาเลือกไฟล์ภาพลายเซ็น");
        return;
      }

      const file = fileInput.files[0];
      const reader = new FileReader();
      showLoading("กำลังอัปโหลดและบันทึกลงโฟลเดอร์ Drive...");

      reader.onload = function(evt) {
        const rawBase64 = evt.target.result;

        if (typeof google !== 'undefined' && google.script && google.script.run) {
          google.script.run
            .withSuccessHandler(res => {
              hideLoading();
              if (res.success) {
                closeUploadModal();
                selectedVaultTab = dept;
                initVaultTabs();
                renderVaultSignatures();
                if (selectedDepartment === dept) {
                  loadDepartmentSignatures(dept);
                }
              } else {
                alert("เกิดข้อผิดพลาด: " + res.message);
              }
            })
            .withFailureHandler(err => {
              hideLoading();
              alert(err.message);
            })
            .uploadSignatureToStock(ownerName, dept, rawBase64);
        } else {
          setTimeout(() => {
            hideLoading();
            closeUploadModal();
            if (!allSignaturesCache[dept]) allSignaturesCache[dept] = [];
            allSignaturesCache[dept].push({ id: Date.now().toString(), name: ownerName, department: dept, dataUri: rawBase64 });
            renderVaultSignatures();
          }, 600);
        }
      };
      reader.readAsDataURL(file);
    }

    // --- ระบบคำนวณและแสดงผลสถิติจริง (Analytics Engine) ---
    function updateAnalyticsView() {
      showLoading("กำลังประมวลผลสถิติ...");
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler(data => {
            hideLoading();
            if (data && data.success) {
              renderAnalyticsData(data);
            } else {
              alert("ไม่สามารถดึงข้อมูลสถิติได้: " + (data ? data.message : "ไม่ทราบสาเหตุ"));
            }
          })
          .withFailureHandler(err => {
            hideLoading();
            alert("เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์: " + err.message);
          })
          .getSigningStatistics();
      } else {
        setTimeout(() => {
          hideLoading();
          renderAnalyticsData({
            totalSignedDocs: 5,
            totalSignaturesInStock: 8,
            departmentStats: { "ส่วนกลาง": 2, "กองบริการ": 1, "กองธุรกิจบริการ": 1, "กองบำรุงรักษา": 1, "กองรายได้": 0 },
            recentActivities: [
              { timestamp: "2026-10-08 14:15:20", fileName: "คำสั่งแต่งตั้ง.pdf", department: "ส่วนกลาง", sigCount: 2, actionType: "ดาวน์โหลด PDF" }
            ]
          });
        }, 400);
      }
    }

    function renderAnalyticsData(data) {
      document.getElementById('stat-signed-count').textContent = `${data.totalSignedDocs} ฉบับ`;
      document.getElementById('stat-total-sigs').textContent = `${data.totalSignaturesInStock} รายการ`;

      const deptContainer = document.getElementById('analytics-dept-breakdown');
      deptContainer.innerHTML = '';
      DEPARTMENTS.forEach(dept => {
        const count = data.departmentStats[dept] || 0;
        const box = document.createElement('div');
        box.className = "p-3.5 bg-surface-container-low rounded-xl border border-outline-variant/20 flex flex-col justify-between";
        box.innerHTML = `
          <span class="text-xs font-semibold text-on-surface truncate">${dept}</span>
          <div class="mt-2 flex items-baseline justify-between">
            <span class="text-xl font-bold text-primary">${count}</span>
            <span class="text-[10px] text-outline font-medium">ฉบับ</span>
          </div>
        `;
        deptContainer.appendChild(box);
      });

      const listContainer = document.getElementById('analytics-recent-list');
      listContainer.innerHTML = '';
      if (!data.recentActivities || data.recentActivities.length === 0) {
        listContainer.innerHTML = '<div class="py-4 text-center text-outline">ยังไม่มีประวัติการลงนาม</div>';
        return;
      }

      data.recentActivities.forEach(item => {
        const row = document.createElement('div');
        row.className = "py-2.5 flex items-center justify-between";
        row.innerHTML = `
          <div class="flex items-center gap-2.5 min-w-0">
            <span class="material-symbols-outlined text-primary text-base">task</span>
            <div class="min-w-0">
              <p class="font-bold text-on-surface truncate max-w-[200px] sm:max-w-xs">${item.fileName}</p>
              <p class="text-[10px] text-outline">${item.department} • วาง ${item.sigCount} จุด</p>
            </div>
          </div>
          <div class="text-right flex-shrink-0">
            <span class="inline-block px-2 py-0.5 rounded-full text-[9px] font-bold ${item.actionType.includes('ดาวน์โหลด') ? 'bg-blue-100 text-blue-700' : 'bg-purple-100 text-purple-700'}">${item.actionType}</span>
            <p class="text-[10px] text-outline mt-0.5">${item.timestamp}</p>
          </div>
        `;
        listContainer.appendChild(row);
      });
    }

    // --- 2. การจัดการและแสดงผล PDF ---
    function handlePdfUpload(e) {
      try {
        const file = e.target.files[0];
        if (!file) return;

        pdfFileName = file.name;

        const headerFilename = document.getElementById('header-filename');
        if (headerFilename) headerFilename.textContent = file.name;

        const sideFilename = document.getElementById('side-filename');
        if (sideFilename) sideFilename.textContent = file.name;

        const sideFileinfo = document.getElementById('side-fileinfo');
        if (sideFileinfo) sideFileinfo.textContent = `${(file.size / (1024 * 1024)).toFixed(1)} MB`;

        const reader = new FileReader();
        reader.onload = function(evt) {
          try {
            pdfBytes = new Uint8Array(evt.target.result);
            placedSignatures = [];
            updatePlacedCountDisplay();
            loadPdfDocument(pdfBytes.slice(0));
          } catch(err) {
            console.error("FileReader onload error:", err);
            alert("เกิดข้อผิดพลาดในการโหลดไฟล์: " + err.message);
          }
        };
        reader.onerror = function() {
          alert("ไม่สามารถอ่านไฟล์ได้ กรุณาลองใหม่อีกครั้ง");
        };
        reader.readAsArrayBuffer(file);
      } catch(err) {
        console.error("handlePdfUpload error:", err);
        alert("เกิดข้อผิดพลาดในการอัปโหลด: " + err.message);
      }
    }

    async function loadPdfDocument(data) {
      showLoading("กำลังโหลดเอกสาร PDF...");
      try {
        const loadingTask = pdfjsLib.getDocument({ data: data });
        pdfDoc = await loadingTask.promise;
        totalPages = pdfDoc.numPages;
        currentPageNum = 1;
        
        const sideFileinfo = document.getElementById('side-fileinfo');
        if (sideFileinfo) {
          sideFileinfo.textContent = `${(data.length / (1024 * 1024)).toFixed(1)} MB • ${totalPages} หน้า`;
        }

        const thumbTitle = document.getElementById('thumbnail-title');
        if (thumbTitle) {
          thumbTitle.textContent = `พรีวิวหน้าเอกสาร (${totalPages})`;
        }

        await renderPage(currentPageNum);
        renderThumbnails();
        hideLoading();
      } catch (err) {
        hideLoading();
        console.error("PDF Loading Error:", err);
        alert("ไม่สามารถเปิดเอกสาร PDF ได้: " + err.message);
      }
    }

    async function renderPage(pageNum) {
      if (!pdfDoc) return;
      const page = await pdfDoc.getPage(pageNum);
      const viewport = page.getViewport({ scale: currentZoom * 1.5 });

      const canvas = document.getElementById('pdf-canvas');
      const ctx = canvas.getContext('2d');
      canvas.width = viewport.width;
      canvas.height = viewport.height;

      const displayW = viewport.width / 1.5;
      const displayH = viewport.height / 1.5;

      canvas.style.width = `${displayW}px`;
      canvas.style.height = `${displayH}px`;

      const wrapper = document.getElementById('pdf-wrapper');
      const sigContainer = document.getElementById('signature-boxes-container');
      if (wrapper) {
        wrapper.style.width = `${displayW}px`;
        wrapper.style.height = `${displayH}px`;
      }
      if (sigContainer) {
        sigContainer.style.width = `${displayW}px`;
        sigContainer.style.height = `${displayH}px`;
      }

      const renderContext = {
        canvasContext: ctx,
        viewport: viewport
      };
      await page.render(renderContext).promise;

      document.getElementById('page-num-display').textContent = `หน้า ${pageNum} / ${totalPages}`;
      document.getElementById('zoom-display').textContent = `${Math.round(currentZoom * 100)}%`;

      renderSignatureBoxesOnPage();
      updateThumbnailsActiveState();
    }

    function changePage(delta) {
      if (!pdfDoc) return;
      const target = currentPageNum + delta;
      if (target >= 1 && target <= totalPages) {
        currentPageNum = target;
        renderPage(currentPageNum);
      }
    }

    function zoomDoc(delta) {
      const newZoom = currentZoom + delta;
      if (newZoom >= 0.5 && newZoom <= 2.2) {
        currentZoom = newZoom;
        renderPage(currentPageNum);
      }
    }

    async function renderThumbnails() {
      const container = document.getElementById('thumbnails-list');
      container.innerHTML = '';

      for (let i = 1; i <= totalPages; i++) {
        const page = await pdfDoc.getPage(i);
        const vp = page.getViewport({ scale: 0.25 });

        const item = document.createElement('div');
        item.id = `thumb-page-${i}`;
        item.className = `rounded-xl p-2 cursor-pointer transition border-2 ${
          i === currentPageNum ? 'border-primary bg-primary-container/10' : 'border-transparent bg-white/70 hover:border-primary/50'
        }`;

        const c = document.createElement('canvas');
        c.width = vp.width;
        c.height = vp.height;
        c.className = 'w-full rounded bg-white shadow-xs';
        const ctx = c.getContext('2d');
        await page.render({ canvasContext: ctx, viewport: vp }).promise;

        const label = document.createElement('span');
        label.className = `text-[10px] mt-1 block text-center font-medium ${i === currentPageNum ? 'text-primary font-bold' : 'text-on-surface-variant'}`;
        label.textContent = `หน้า ${i}${i === currentPageNum ? ' (กำลังดู)' : ''}`;

        item.appendChild(c);
        item.appendChild(label);
        item.onclick = () => {
          currentPageNum = i;
          renderPage(currentPageNum);
        };
        container.appendChild(item);
      }
    }

    function updateThumbnailsActiveState() {
      for (let i = 1; i <= totalPages; i++) {
        const el = document.getElementById(`thumb-page-${i}`);
        if (!el) continue;
        const label = el.querySelector('span');
        if (i === currentPageNum) {
          el.className = 'rounded-xl p-2 cursor-pointer transition border-2 border-primary bg-primary-container/10';
          if (label) {
            label.className = 'text-[10px] mt-1 block text-center font-bold text-primary';
            label.textContent = `หน้า ${i} (กำลังดู)`;
          }
        } else {
          el.className = 'rounded-xl p-2 cursor-pointer transition border-2 border-transparent bg-white/70 hover:border-primary/50';
          if (label) {
            label.className = 'text-[10px] mt-1 block text-center font-medium text-on-surface-variant';
            label.textContent = `หน้า ${i}`;
          }
        }
      }
    }

    // --- 3. ระบบพิมพ์ข้อความและปรับแต่ง (Text Generator) ---
    function toggleTextCustomizerAccordion() {
      const content = document.getElementById('text-customizer-content');
      const icon = document.getElementById('accordion-icon');
      if (content.classList.contains('hidden')) {
        content.classList.remove('hidden');
        icon.textContent = 'expand_less';
      } else {
        content.classList.add('hidden');
        icon.textContent = 'expand_more';
      }
    }

    function autoResizeTextarea(el) {
      el.style.height = 'auto';
      el.style.height = Math.min(el.scrollHeight, 140) + 'px';
    }

    function setFontWeight(weight) {
      currentFontWeight = weight;
      ['300', '400', '700'].forEach(w => {
        const btn = document.getElementById(`weight-btn-${w}`);
        if (w === weight) {
          btn.className = "py-1 text-[11px] rounded bg-white text-primary font-bold shadow-xs transition";
        } else {
          btn.className = "py-1 text-[11px] rounded text-outline hover:text-on-surface transition";
        }
      });
    }

    function setTextAlign(align) {
      currentTextAlign = align;
      ['left', 'center', 'right'].forEach(a => {
        const btn = document.getElementById(`align-btn-${a}`);
        if (a === align) {
          btn.className = "p-1 rounded bg-white text-primary shadow-xs transition";
        } else {
          btn.className = "p-1 rounded text-outline hover:text-on-surface transition";
        }
      });
    }

    function generateAndPlaceTextSignature() {
      if (!pdfDoc) {
        alert("กรุณาอัปโหลดเอกสาร PDF ก่อนวางข้อความ");
        return;
      }
      const text = document.getElementById('custom-text-input').value.trim();
      if (!text) {
        alert("กรุณาพิมพ์ข้อความที่ต้องการวางก่อนครับ");
        return;
      }

      const fontFam = document.getElementById('font-family-select').value;
      const fontSize = parseInt(document.getElementById('font-size-select').value, 10);
      const color = document.getElementById('text-color-picker').value;

      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      const fontSpec = `${currentFontWeight} ${fontSize * 2}px ${fontFam}`;
      ctx.font = fontSpec;

      const lines = text.split('\n');
      let maxLineWidth = 0;
      lines.forEach(line => {
        const w = ctx.measureText(line).width;
        if (w > maxLineWidth) maxLineWidth = w;
      });

      const lineHeight = fontSize * 2.4;
      canvas.width = Math.max(200, maxLineWidth + 40);
      canvas.height = Math.max(80, (lines.length * lineHeight) + 30);

      ctx.font = fontSpec;
      ctx.fillStyle = color;
      ctx.textBaseline = 'middle';

      lines.forEach((line, index) => {
        let xPos = 20;
        if (currentTextAlign === 'center') {
          ctx.textAlign = 'center';
          xPos = canvas.width / 2;
        } else if (currentTextAlign === 'right') {
          ctx.textAlign = 'right';
          xPos = canvas.width - 20;
        } else {
          ctx.textAlign = 'left';
          xPos = 20;
        }
        const yPos = 25 + (index * lineHeight) + (lineHeight / 2);
        ctx.fillText(line, xPos, yPos);
      });

      const textDataUri = canvas.toDataURL('image/png');
      const offset = (placedSignatures.filter(s => s.page === currentPageNum).length * 25) % 150;
      const newSig = {
        id: 'sig_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4),
        page: currentPageNum,
        x: 80 + offset,
        y: 120 + offset,
        width: Math.min(220, Math.max(160, canvas.width / 2)),
        height: Math.max(65, canvas.height / 2),
        dataUri: textDataUri,
        confirmed: false
      };

      placedSignatures.push(newSig);
      updatePlacedCountDisplay();
      renderSignatureBoxesOnPage();
    }

    // --- 4. การจัดการกล่องบน Canvas ---
    function updatePlacedCountDisplay() {
      const count = placedSignatures.length;
      document.getElementById('placed-count-display').textContent = `วางแล้ว ${count} จุด`;
    }

    function clearAllSignaturesOnPage() {
      const onThisPage = placedSignatures.filter(s => s.page === currentPageNum);
      if (onThisPage.length === 0) return;
      if (!confirm(`คุณต้องการลบลายเซ็นทั้งหมด (${onThisPage.length} จุด) ในหน้านี้ใช่หรือไม่?`)) return;

      placedSignatures = placedSignatures.filter(s => s.page !== currentPageNum);
      updatePlacedCountDisplay();
      renderSignatureBoxesOnPage();
    }

    function removeSignatureItem(sigId) {
      placedSignatures = placedSignatures.filter(s => s.id !== sigId);
      updatePlacedCountDisplay();
      renderSignatureBoxesOnPage();
    }

    function toggleConfirmSignature(sigId) {
      const sig = placedSignatures.find(s => s.id === sigId);
      if (sig) {
        sig.confirmed = !sig.confirmed;
        renderSignatureBoxesOnPage();
      }
    }

    function renderSignatureBoxesOnPage() {
      const container = document.getElementById('signature-boxes-container');
      container.innerHTML = '';

      const currentPageSigs = placedSignatures.filter(s => s.page === currentPageNum);

      currentPageSigs.forEach(sig => {
        const box = document.createElement('div');
        box.id = sig.id;
        box.className = `placed-sig-box rounded-xl group pointer-events-auto shadow-sm ${
          sig.confirmed ? 'confirmed-box' : 'active-box'
        }`;
        box.style.left = `${sig.x}px`;
        box.style.top = `${sig.y}px`;
        box.style.width = `${sig.width}px`;
        box.style.height = `${sig.height}px`;

        box.innerHTML = `
          <img src="${sig.dataUri}" class="w-full h-full object-contain pointer-events-none p-1.5" />
          <button onclick="removeSignatureItem('${sig.id}')" title="ลบลายเซ็นจุดนี้" 
                  class="absolute -top-3 -left-3 w-6 h-6 rounded-full bg-red-500 hover:bg-red-600 text-white flex items-center justify-center shadow-md transition transform hover:scale-110 z-50">
            <span class="material-symbols-outlined text-[14px]">close</span>
          </button>
          <button onclick="toggleConfirmSignature('${sig.id}')" title="${sig.confirmed ? 'คลิกเพื่อแก้ไขตำแหน่ง' : 'คลิกเพื่อยืนยันจุดวางนี้'}" 
                  class="absolute -top-3 -right-3 h-6 px-2 rounded-full ${sig.confirmed ? 'bg-green-600 hover:bg-green-700' : 'bg-primary hover:bg-primary/90'} text-white text-[11px] font-bold flex items-center gap-1 shadow-md transition transform hover:scale-105 z-50">
            <span class="material-symbols-outlined text-[14px]">${sig.confirmed ? 'check' : 'done'}</span>
            <span>${sig.confirmed ? 'ยืนยันแล้ว' : 'ยืนยันจุดวาง'}</span>
          </button>
          <div class="sig-resize-handle"></div>
        `;

        setupBoxInteractions(box, sig);
        container.appendChild(box);
      });
    }

    function setupBoxInteractions(boxEl, sigData) {
      let isDragging = false;
      let isResizing = false;
      let startX = 0, startY = 0;
      let initX = 0, initY = 0, initW = 0, initH = 0;

      const resizeHandle = boxEl.querySelector('.sig-resize-handle');

      boxEl.addEventListener('mousedown', (e) => {
        if (e.target.closest('button') || e.target === resizeHandle) return;
        isDragging = true;
        startX = e.clientX;
        startY = e.clientY;
        initX = sigData.x;
        initY = sigData.y;

        const onMouseMove = (moveEvt) => {
          if (!isDragging) return;
          const dx = moveEvt.clientX - startX;
          const dy = moveEvt.clientY - startY;
          sigData.x = Math.max(0, initX + dx);
          sigData.y = Math.max(0, initY + dy);
          boxEl.style.left = `${sigData.x}px`;
          boxEl.style.top = `${sigData.y}px`;
        };

        const onMouseUp = () => {
          isDragging = false;
          document.removeEventListener('mousemove', onMouseMove);
          document.removeEventListener('mouseup', onMouseUp);
        };

        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onMouseUp);
        e.preventDefault();
      });

      resizeHandle.addEventListener('mousedown', (e) => {
        isResizing = true;
        startX = e.clientX;
        startY = e.clientY;
        initW = sigData.width;
        initH = sigData.height;

        const onResizeMove = (moveEvt) => {
          if (!isResizing) return;
          const dx = moveEvt.clientX - startX;
          const dy = moveEvt.clientY - startY;
          sigData.width = Math.max(60, initW + dx);
          sigData.height = Math.max(25, initH + dy);
          boxEl.style.width = `${sigData.width}px`;
          boxEl.style.height = `${sigData.height}px`;
        };

        const onResizeUp = () => {
          isResizing = false;
          document.removeEventListener('mousemove', onResizeMove);
          document.removeEventListener('mouseup', onResizeUp);
        };

        document.addEventListener('mousemove', onResizeMove);
        document.addEventListener('mouseup', onResizeUp);
        e.stopPropagation();
        e.preventDefault();
      });
    }

    // --- 5. Export & Print Engine (Scale 3.5x คมชัดสูง + บันทึกสถิติลง Sheet) ---
    async function getCompositePageCanvas(targetPageNum) {
      const page = await pdfDoc.getPage(targetPageNum);
      const PRINT_SCALE = 3.5;
      const viewport = page.getViewport({ scale: PRINT_SCALE });

      const compositeCanvas = document.createElement('canvas');
      compositeCanvas.width = viewport.width;
      compositeCanvas.height = viewport.height;
      const ctx = compositeCanvas.getContext('2d');

      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = 'high';

      await page.render({ canvasContext: ctx, viewport: viewport }).promise;

      const pageSigs = placedSignatures.filter(s => s.page === targetPageNum);

      if (pageSigs.length > 0) {
        const screenCanvas = document.getElementById('pdf-canvas');
        const displayW = parseFloat(screenCanvas.style.width);
        const displayH = parseFloat(screenCanvas.style.height);

        const ratioX = viewport.width / displayW;
        const ratioY = viewport.height / displayH;

        for (const sig of pageSigs) {
          const sigImg = new Image();
          sigImg.src = sig.dataUri;
          await new Promise(r => sigImg.onload = r);

          ctx.drawImage(
            sigImg,
            sig.x * ratioX,
            sig.y * ratioY,
            sig.width * ratioX,
            sig.height * ratioY
          );
        }
      }

      return {
        canvas: compositeCanvas,
        widthPt: viewport.width / PRINT_SCALE,
        heightPt: viewport.height / PRINT_SCALE
      };
    }

    function sendAuditLog(actionType) {
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run.logSigningAction(
          pdfFileName,
          selectedDepartment,
          placedSignatures.length,
          actionType
        );
      }
    }

    async function downloadSignedPdfDocument() {
      if (!pdfDoc) {
        alert("กรุณาอัปโหลดเอกสาร PDF ก่อนบันทึก");
        return;
      }
      if (placedSignatures.length === 0) {
        alert("กรุณาวางลายเซ็นหรือข้อความบนเอกสารอย่างน้อย 1 จุดก่อนดาวน์โหลด");
        return;
      }

      showLoading("กำลังสร้างไฟล์เอกสาร PDF ความละเอียดสูง...");
      try {
        const { jsPDF } = window.jspdf;
        let pdfOutput = null;

        for (let pageIdx = 1; pageIdx <= totalPages; pageIdx++) {
          showLoading(`กำลังประมวลผลหน้าที่ ${pageIdx} จาก ${totalPages} (ความละเอียดสูง)...`);
          const { canvas, widthPt, heightPt } = await getCompositePageCanvas(pageIdx);
          
          const imgData = canvas.toDataURL('image/jpeg', 0.98);
          const orientation = widthPt > heightPt ? 'l' : 'p';

          if (pageIdx === 1) {
            pdfOutput = new jsPDF({
              orientation: orientation,
              unit: 'pt',
              format: [widthPt, heightPt],
              compress: true
            });
            pdfOutput.addImage(imgData, 'JPEG', 0, 0, widthPt, heightPt, undefined, 'FAST');
          } else {
            pdfOutput.addPage([widthPt, heightPt], orientation);
            pdfOutput.addImage(imgData, 'JPEG', 0, 0, widthPt, heightPt, undefined, 'FAST');
          }
        }

        const baseName = pdfFileName.replace(/\.[^/.]+$/, "");
        pdfOutput.save(`Signed_${baseName}.pdf`);

        sendAuditLog("ดาวน์โหลด PDF");
        hideLoading();
      } catch (err) {
        hideLoading();
        console.error(err);
        alert("เกิดข้อผิดพลาดในการสร้างไฟล์ PDF: " + err.message);
      }
    }

    async function printSignedDocumentDirectly() {
      if (!pdfDoc) {
        alert("กรุณาอัปโหลดเอกสาร PDF ก่อนสั่งพิมพ์");
        return;
      }
      if (placedSignatures.length === 0) {
        alert("กรุณาวางลายเซ็นหรือข้อความบนเอกสารอย่างน้อย 1 จุดก่อนสั่งพิมพ์");
        return;
      }

      showLoading("กำลังเตรียมเอกสารความละเอียดสูงสำหรับสั่งพิมพ์...");
      try {
        let pagesHtml = "";

        for (let pageIdx = 1; pageIdx <= totalPages; pageIdx++) {
          showLoading(`กำลังเตรียมหน้าที่ ${pageIdx} จาก ${totalPages}...`);
          const { canvas } = await getCompositePageCanvas(pageIdx);
          const dataUrl = canvas.toDataURL('image/jpeg', 0.98);

          pagesHtml += `
            <div class="print-page">
              <img src="${dataUrl}" />
            </div>
          `;
        }

        const printWindow = window.open('', '_blank');
        printWindow.document.write(`
          <!DOCTYPE html>
          <html>
          <head>
            <title>พิมพ์เอกสาร - ${pdfFileName}</title>
            <style>
              @page { size: auto; margin: 0mm; }
              body { margin: 0; padding: 0; background-color: #ffffff; }
              .print-page {
                width: 100%; height: 100vh;
                display: flex; justify-content: center; align-items: center;
                page-break-after: always; page-break-inside: avoid;
                box-sizing: border-box;
              }
              .print-page:last-child { page-break-after: auto; }
              img { max-width: 100%; max-height: 100%; width: auto; height: auto; display: block; object-fit: contain; }
            </style>
          </head>
          <body>
            ${pagesHtml}
            <script>
              window.onload = function() {
                setTimeout(function() {
                  window.print();
                  window.close();
                }, 500);
              };
            <\/script>
          </body>
          </html>
        `);
        printWindow.document.close();

        sendAuditLog("สั่งพิมพ์เอกสาร");
        hideLoading();
      } catch (err) {
        hideLoading();
        console.error(err);
        alert("เกิดข้อผิดพลาดในการพิมพ์เอกสาร: " + err.message);
      }
    }
  </script>
</body>
</html>
----------------------------------------------------------------------

1931001	สสน.
1931101	กรก.สสน.
1931201	กรร.สสน.
1931301	กรด.สสน.
1931401	กธบ.สสน.
