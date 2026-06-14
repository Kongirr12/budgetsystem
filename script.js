// ==============================
// การตั้งค่า API (เชื่อมกับ Google Apps Script)
// ==============================
// ⚠️ เปลี่ยน URL ด้านล่างให้เป็น URL Web App ของคุณเองที่ได้จากการ Deploy ใน Google Apps Script
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbx5R3HaXBNjOLzQb-Wdo69gNJnQYh0eux5ONsJKSbCHlIuafdbRijGxkjSgQx4EKXYezw/exec";

// 🚀 ฟังก์ชันครอบจักรวาล ทำหน้าที่แทน google.script.run
async function callGoogleScript(action, params = {}) {
  try {
    const response = await fetch(WEB_APP_URL, {
      method: 'POST',
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ action: action, ...params })
    });
    const result = await response.json();
    return result;
  } catch (error) {
    console.error("API Error:", error);
    throw error;
  }
}

// ==============================
// ตัวแปรหลัก
// ==============================
var summaryData = []; 
var projects = []; 
var latestTransactions = []; 
var isLoggedIn = localStorage.getItem('isLoggedIn') === 'true'; 
var fullName = localStorage.getItem('fullName') || ''; 
var userRole = localStorage.getItem('userRole') || ''; 
var authToken = localStorage.getItem('authToken') || ''; 
var assignableUsers = []; 
var tomSelectOwner = null; 
var budgetChartInstance = null; 
var uploadModal = null;
var currentUploadTxId = null;
var uploadMode = 'read';     
var stagedUploads = [];   
var stagedDeletions = []; 

function loadingStart() { $('#loading').removeClass('hidden'); }
function loadingEnd() { $('#loading').addClass('hidden'); }

// ==============================
// Helpers
// ==============================
function getUserFullName(userId) {
  if (!userId) return '(ไม่ได้ระบุ)';
  const user = assignableUsers.find(function(u) { return u.id === userId; });
  return user ? user.fullName : '(ไม่พบชื่อ)';
}

function _parseOwnerIds(ownerData) {
  if (!ownerData) return []; 
  if (typeof ownerData === 'string' && ownerData.startsWith('[')) {
    try {
      const ids = JSON.parse(ownerData);
      return Array.isArray(ids) ? ids : [];
    } catch (e) { return []; }
  }
  if (typeof ownerData === 'string' && ownerData.length > 0) return [ownerData];
  return [];
}

function onFailure(error) {
  loadingEnd();
  console.error('GAS Server Error:', error);
  let errorName = "Error";
  let errorMessage = "";
  if (typeof error === 'object' && error !== null) {
    errorName = error.name || "Error";
    errorMessage = typeof error.message === 'object' ? JSON.stringify(error.message) : error.message;
  } else if (typeof error === 'string') {
    errorMessage = error;
  }
  errorMessage = String(errorMessage || "การเชื่อมต่อล้มเหลว หรือ Server ไม่ตอบสนอง");
  let errorText = '[' + errorName + '] ' + errorMessage + '. กรุณาลองใหม่อีกครั้ง หรือติดต่อผู้ดูแลระบบ';

  if (errorMessage.includes("Invalid or expired token")) {
    errorText = "เซสชันหมดอายุ กรุณาล็อกอินใหม่อีกครั้ง";
    forceLogout(); 
  }
  if (errorMessage.includes("HTTP 429") || errorMessage.includes("Too Many Requests")) {
    errorName = "TooManyRequests";
    errorText = "ขณะนี้ Server กำลังประมวลผลคำขอจำนวนมาก (HTTP 429) กรุณารอสักครู่แล้วลองใหม่อีกครั้ง";
  }
  Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาด (' + errorName + ')', text: errorText });
}

function forceLogout() {
  isLoggedIn = false; fullName = ''; userRole = ''; authToken = '';
  localStorage.clear();
  window.location.reload(); // รีโหลดหน้าแทนการใช้ getScriptUrl
}

// ==============================
// เริ่มต้นระบบ
// ==============================
document.addEventListener('DOMContentLoaded', function() {
  toastr.options = { "positionClass": "toast-bottom-right", "timeOut": "3000", "progressBar": true };

  updateNavbar(); 
  loadProjectSummary(); 

  if (isLoggedIn) {
    loadProjects();
    showSection('public'); 
  } else {
    showSection('public');
  }

  // ตั้งค่า Admin Tabs & DataTables
  const adminTabs = document.getElementById('adminTabs');
  if (adminTabs) {
    const projectTab = adminTabs.querySelector('a[href="#tab-projects"]');
    const userTab = adminTabs.querySelector('a[href="#tab-users"]');
    const settingTab = adminTabs.querySelector('a[href="#tab-settings"]'); 
    if(projectTab) {
      projectTab.addEventListener('shown.bs.tab', function() {
         if ($.fn.DataTable.isDataTable('#tableProjects')) $('#tableProjects').DataTable().columns.adjust();
         else loadProjectData();
      });
    }
    if(userTab) {
      userTab.addEventListener('shown.bs.tab', function() {
        if ($.fn.DataTable.isDataTable('#tableUsers')) $('#tableUsers').DataTable().columns.adjust();
        else loadUsersData();
      });
    }
    if(settingTab) {
      settingTab.addEventListener('shown.bs.tab', function() { loadSystemSettings(); });
    }
    if (isLoggedIn && document.querySelector('#tab-transactions').classList.contains('active')) {
       loadDataTransactions();
    }
  }

  // ปุ่ม Excel
  const btnTemplate = document.getElementById('btn-template-project');
  const btnExport = document.getElementById('btn-export-project');
  const btnImport = document.getElementById('btn-import-project');
  const fileUploader = document.getElementById('projectFileUploader');
  if (btnTemplate) btnTemplate.addEventListener('click', downloadProjectTemplate);
  if (btnExport) btnExport.addEventListener('click', exportProjects);
  if (btnImport) btnImport.addEventListener('click', function() { if (fileUploader) fileUploader.click(); });
  if (fileUploader) fileUploader.addEventListener('change', handleProjectImport);

  // ตั้งค่า Modal แนบไฟล์
  if (document.getElementById('modalUploadFiles')) {
    uploadModal = new bootstrap.Modal(document.getElementById('modalUploadFiles'));
    setupDropzones(); 
    document.getElementById('btnUploadEdit').addEventListener('click', function() { setUploadMode('edit'); });
    document.getElementById('btnUploadCancel').addEventListener('click', function() { cancelUploadChanges(); });
    document.getElementById('btnUploadSave').addEventListener('click', function() { saveFileChangesClient(); });
  }

  // ฟอร์มล็อกอิน
  document.getElementById('loginForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    if (!username || !password) return Swal.fire('ผิดพลาด', 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน', 'error');

    loadingStart();
    // 🚀 เปลี่ยนการดึงข้อมูลล็อกอินมาใช้ API
    callGoogleScript('authenticateUser', { username: username, password: password })
      .then(function(response) {
        loadingEnd();
        if (response.success) {
          isLoggedIn = true; fullName = response.fullName; userRole = response.role; authToken = response.token;
          localStorage.setItem('isLoggedIn', 'true');
          localStorage.setItem('fullName', fullName);
          localStorage.setItem('userRole', userRole);
          localStorage.setItem('authToken', authToken);
          updateNavbar(); 
          showSection('public');
          loadProjectSummary(); 
          loadProjects(); 
          loadDataTransactions(); 
          Swal.fire('สำเร็จ', response.message, 'success');
          bootstrap.Modal.getInstance(document.getElementById('loginModal')).hide();
        } else {
          Swal.fire('ผิดพลาด', response.message, 'error');
        }
      })
      .catch(onFailure);
  });

  // ปุ่ม Logout
  document.getElementById('logoutBtn').addEventListener('click', function() {
    Swal.fire({
      title: 'ยืนยันการล็อกเอาท์',
      text: 'คุณต้องการออกจากระบบใช่หรือไม่?',
      icon: 'question', showCancelButton: true, confirmButtonText: 'ล็อกเอาท์', cancelButtonText: 'ยกเลิก',
      confirmButtonColor: 'var(--danger-color)'
    }).then(function(result) {
      if (result.isConfirmed) {
        forceLogout(); // เรียกใช้ฟังก์ชันเคลียร์และรีเฟรชหน้าเว็บ
      }
    });
  });
});

// ==============================
// Section และ Navbar
// ==============================
function showSection(sectionId) {
  if (sectionId === 'admin' && !isLoggedIn) {
    Swal.fire('กรุณาล็อกอิน', 'คุณต้องล็อกอินเพื่อใช้งานส่วนดำเนินการ', 'warning');
    bootstrap.Modal.getOrCreateInstance(document.getElementById('loginModal')).show();
    return;
  }
  document.querySelectorAll('.section').forEach(function(s) { s.classList.remove('active'); });
  const el = document.getElementById(sectionId);
  if (el) el.classList.add('active');
  window.scrollTo(0, 0);
  document.querySelectorAll('#navLinks .nav-link').forEach(function(link) {
    link.classList.remove('active');
    if (link.getAttribute('href') === '#' + sectionId) link.classList.add('active');
  });
}

function updateNavbar() {
  const navLinks = document.getElementById('navLinks');
  const userInfo = document.getElementById('userInfo');
  const logoutBtn = document.getElementById('logoutBtn');
  const navProjectTab = document.getElementById('nav-project-tab');
  const navUserTab = document.getElementById('nav-user-tab');
  const btnAddProject = document.getElementById('btn-add-project');
  const btnAddUser = document.getElementById('btn-add-user');
  const formTransaction = document.getElementById('form-transaction');
  const navSettingTab = document.getElementById('nav-setting-tab'); 
  const btnTemplateProject = document.getElementById('btn-template-project');
  const btnExportProject = document.getElementById('btn-export-project');
  const btnImportProject = document.getElementById('btn-import-project');

  if (isLoggedIn) {
    userInfo.textContent = 'สวัสดี, ' + fullName + ' (' + userRole + ')'; 
    logoutBtn.style.display = 'inline-block';
    navLinks.innerHTML = '<li class="nav-item"><a class="nav-link" href="#admin" onclick="showSection(\'admin\')"><i class="fa-solid fa-user-shield me-1"></i> ดำเนินการ</a></li><li class="nav-item"><a class="nav-link" href="#public" onclick="showSection(\'public\')"><i class="fa-solid fa-chart-line me-1"></i> ข้อมูลทั่วไป</a></li>'; 
    navLinks.classList.add('me-auto'); navLinks.classList.remove('ms-auto');

    const dashboardPanel = document.getElementById('admin-dashboard-panel');
    if (dashboardPanel) dashboardPanel.classList.remove('d-none');

    if (typeof renderAdminDashboard === 'function') {
        callGoogleScript('getProjectSummary', { token: authToken })
         .then(function(res) { if(res.success) renderAdminDashboard(res.dashboardData || res.data); })
         .catch(console.error);
    }

    if (userRole === 'Admin') {
      if (navProjectTab) navProjectTab.classList.remove('d-none');
      if (navUserTab) navUserTab.classList.remove('d-none');
      if (navSettingTab) navSettingTab.classList.remove('d-none'); 
      if (btnAddProject) btnAddProject.classList.remove('d-none');
      if (btnAddUser) btnAddUser.classList.remove('d-none');
      if (formTransaction) formTransaction.classList.remove('d-none');
      if (btnTemplateProject) btnTemplateProject.classList.remove('d-none');
      if (btnExportProject) btnExportProject.classList.remove('d-none');
      if (btnImportProject) btnImportProject.classList.remove('d-none');
      if (assignableUsers.length === 0) loadAssignableUsers();

    } else if (userRole === 'Staff') {
      if (navProjectTab) navProjectTab.classList.add('d-none');
      if (navUserTab) navUserTab.classList.add('d-none');
      if (navSettingTab) navSettingTab.classList.add('d-none'); 
      if (btnAddProject) btnAddProject.classList.add('d-none');
      if (btnAddUser) btnAddUser.classList.add('d-none');
      if (formTransaction) formTransaction.classList.remove('d-none'); 
      if (btnTemplateProject) btnTemplateProject.classList.add('d-none');
      if (btnExportProject) btnExportProject.classList.add('d-none');
      if (btnImportProject) btnImportProject.classList.add('d-none');
      if (assignableUsers.length === 0) loadAssignableUsers();

    } else {
      if (navProjectTab) navProjectTab.classList.add('d-none');
      if (navUserTab) navUserTab.classList.add('d-none');
      if (navSettingTab) navSettingTab.classList.add('d-none'); 
      if (btnAddProject) btnAddProject.classList.add('d-none');
      if (btnAddUser) btnAddUser.classList.add('d-none');
      if (formTransaction) formTransaction.classList.add('d-none'); 
      if (btnTemplateProject) btnTemplateProject.classList.add('d-none');
      if (btnExportProject) btnExportProject.classList.add('d-none');
      if (btnImportProject) btnImportProject.classList.add('d-none');
    }
  } else {
    userInfo.textContent = ''; logoutBtn.style.display = 'none';
    navLinks.innerHTML = '<li class="nav-item"><a class="nav-link" href="#" data-bs-toggle="modal" data-bs-target="#loginModal"><i class="fa-solid fa-right-to-bracket me-1"></i> เข้าสู่ระบบ</a></li>'; 
    navLinks.classList.add('ms-auto'); navLinks.classList.remove('me-auto');
    const dashboardPanel = document.getElementById('admin-dashboard-panel');
    if (dashboardPanel) dashboardPanel.classList.add('d-none');
    if (navProjectTab) navProjectTab.classList.add('d-none');
    if (navUserTab) navUserTab.classList.add('d-none');
    if (navSettingTab) navSettingTab.classList.add('d-none'); 
    if (btnAddProject) btnAddProject.classList.add('d-none');
    if (btnAddUser) btnAddUser.classList.add('d-none');
    if (formTransaction) formTransaction.classList.add('d-none');
    if (btnTemplateProject) btnTemplateProject.classList.add('d-none');
    if (btnExportProject) btnExportProject.classList.add('d-none');
    if (btnImportProject) btnImportProject.classList.add('d-none');
    assignableUsers = []; 
  }
}

function requireLogin() {
  if (!isLoggedIn || !authToken) { 
    Swal.fire('กรุณาล็อกอิน', 'คุณต้องล็อกอินเพื่อใช้งานส่วนดำเนินการ', 'warning');
    bootstrap.Modal.getOrCreateInstance(document.getElementById('loginModal')).show();
    return false;
  }
  return true;
}

// ==============================
// ข้อมูลผู้ใช้
// ==============================
function loadAssignableUsers() {
  if (userRole !== 'Admin' && userRole !== 'Staff') return; 
  callGoogleScript('getAssignableUsersList', { token: authToken })
    .then(function(users) { assignableUsers = users; })
    .catch(onFailure);
}

// ==============================
// โมดูล Transactions & Projects
// ==============================
function loadProjects() {
  if (!isLoggedIn) return;
  loadingStart();
  callGoogleScript('getProjects', { token: authToken })
    .then(function(data) {
      loadingEnd();
      projects = data; 
      const select = document.getElementById('project');
      if (!select) return; 
      select.innerHTML = '<option value="">-- กรุณาเลือก --</option>';
      data.forEach(function(project) {
        const option = document.createElement('option');
        option.value = project.code; option.textContent = project.code + ' - ' + project.name; 
        select.appendChild(option);
      });
      onProjectChange();
    })
    .catch(onFailure);
}

function onProjectChange() {
  const code = document.getElementById('project').value;
  const project = projects.find(function(p) { return p.code === code; }); 
  const projectData = summaryData.find(function(p) { return p.code === code; });
  const infoDiv = document.getElementById('projectInfo');
  if (!infoDiv) return;

  const amountInput = document.getElementById('amount');
  const amountLabel = document.getElementById('labelAmount');
  const submitButton = document.getElementById('btnSubmitTransaction');
  if (project) {
    infoDiv.classList.remove('d-none');
    document.getElementById('infoName').textContent = project.name;
    const budget = parseFloat(project.budget) || 0; 
    document.getElementById('infoBudget').textContent = budget.toLocaleString() + " บาท";
    const ownerIds = _parseOwnerIds(project.owner);
    document.getElementById('infoOwner').textContent = ownerIds.map(getUserFullName).join(', ') || '(ไม่ได้ระบุ)';

    if (projectData) {
      document.getElementById('balanceAmount').textContent = Number(projectData.balance).toLocaleString() + " บาท";
      document.getElementById('sequenceCount').textContent = projectData.txCount || "0";
    } else {
      document.getElementById('balanceAmount').textContent = budget.toLocaleString() + " บาท";
      document.getElementById('sequenceCount').textContent = "0";
    }
    
    if (budget === 0) {
      amountInput.value = 0; amountInput.disabled = true;
      amountLabel.innerHTML = '<i class="fa-solid fa-file-lines me-1 text-info"></i> บันทึกรายงาน (ไม่ใช้งบประมาณ)';
      submitButton.innerHTML = '<i class="fa-solid fa-paper-plane me-1"></i> บันทึกรายงาน';
      submitButton.className = 'btn btn-info';
    } else {
      amountInput.value = ''; amountInput.disabled = false; amountInput.placeholder = 'ระบุจำนวนเงินที่ต้องการเบิก';
      amountLabel.innerHTML = '<i class="fa-solid fa-money-bill-wave me-1 text-success"></i> จำนวนเงินที่ต้องการเบิก';
      submitButton.innerHTML = '<i class="fa-solid fa-paper-plane me-1"></i> บันทึกการเบิกเงิน';
      submitButton.className = 'btn btn-primary';
    }
  } else {
    infoDiv.classList.add('d-none');
    document.getElementById('infoName').textContent = "-";
    document.getElementById('infoBudget').textContent = "-";
    document.getElementById('infoOwner').textContent = "-";
    document.getElementById('balanceAmount').textContent = "-";
    document.getElementById('sequenceCount').textContent = "-";
    amountInput.value = ''; amountInput.disabled = false;
    submitButton.className = 'btn btn-primary';
  }
}

function submitTransaction() {
  if (userRole === 'Viewer') return toastr.error("⚠️ คุณไม่มีสิทธิ์ในการเบิกเงิน");
  if (!requireLogin()) return;
  const projectCode = document.getElementById('project').value;
  const amount = parseFloat(document.getElementById('amount').value);
  if (!projectCode) return toastr.error("⚠️ กรุณาเลือกโครงการก่อนทำรายการ");
  if (amount === null || amount === undefined || amount < 0 || isNaN(amount)) return toastr.error("⚠️ กรุณากรอกจำนวนเงินที่ถูกต้อง");

  const project = projects.find(function(p) { return p.code === projectCode; }); 
  const projectName = project ? project.name : '(ไม่ทราบชื่อโครงการ)';
  const budget = parseFloat(project.budget) || 0;

  let swalTitle = 'ยืนยันการเบิกเงิน';
  let swalHtml = '<p><i class="fa-solid fa-diagram-project text-primary me-1"></i> โครงการ: <strong>' + projectName + '</strong></p>' +
      '<p><i class="fa-solid fa-coins text-warning me-1"></i> จำนวนเงิน: <strong>' + amount.toLocaleString() + '</strong> บาท</p>';
  
  if (budget === 0 && amount === 0) {
    swalTitle = 'ยืนยันการบันทึกรายงาน';
    swalHtml = '<p><i class="fa-solid fa-diagram-project text-primary me-1"></i> โครงการ: <strong>' + projectName + '</strong></p><p class="text-info">บันทึกรายงานนี้ (ไม่ใช้งบประมาณ)</p>';
  }

  Swal.fire({
    title: swalTitle, html: swalHtml, icon: 'question', showCancelButton: true, confirmButtonText: 'ใช่,ยืนยัน', confirmButtonColor: 'var(--success-color)'
  }).then(function(result) { 
    if (result.isConfirmed) {
      loadingStart();
      callGoogleScript('submitTransaction', { token: authToken, projectCode: projectCode, amount: amount })
        .then(function(response) {
          loadingEnd();
          if (response.success) {
            let successTitle = amount === 0 ? 'บันทึกรายงานสำเร็จ!' : 'สำเร็จ!';
            let successHtml = amount === 0 ? 'ระบบได้สร้างรายการสำหรับแนบไฟล์แล้ว' : response.message + '<br>✅ ครั้งที่เบิก: ' + response.sequence + '<br>💰 ยอดคงเหลือ: ' + Number(response.balance).toLocaleString() + ' บาท';
            Swal.fire({ icon: 'success', title: successTitle, html: successHtml });
            onProjectChange(); loadDataTransactions(); loadProjectSummary(); 
          } else { Swal.fire('เกิดข้อผิดพลาด!', response.message, 'error'); }
        }).catch(onFailure);
    }
  });
}

function loadDataTransactions() {
  if (!requireLogin()) return; 
  loadingStart();
  callGoogleScript('getDataTransactions', { token: authToken })
    .then(function(data) {
      loadingEnd(); latestTransactions = data; showTableTransactions(data); 
    }).catch(onFailure);
}

function showTableTransactions(items) {
  const tableId = '#tableTransactions';
  if ($.fn.DataTable.isDataTable(tableId)) $(tableId).DataTable().destroy();
  if (!items || items.length === 0) {
    $(tableId).html("<thead><tr><th>#</th><th>โครงการ</th><th>จำนวนเงิน</th><th>ครั้งที่</th><th>คงเหลือ</th><th>Action</th></tr></thead><tbody><tr><td colspan='6' class='text-center'>ไม่พบข้อมูล</td></tr></tbody>");
    return;
  }
  new DataTable(tableId, {
    destroy: true, responsive: true, pageLength: 10, data: items, order: [[4, 'desc']],
    dom: '<"d-flex justify-content-between align-items-center mb-3"Bf>rt<"d-flex justify-content-between mt-3"ip>',
    buttons: [{
      extend: 'print', text: '<i class="fa-solid fa-print me-1"></i> พิมพ์รายการ', className: 'btn btn-primary btn-sm rounded-pill px-3',
      title: 'รายการเบิกงบประมาณ โรงเรียนมหาชัยพิทยาคาร',
      messageBottom: '<div style="text-align: right; margin-top: 30px; font-size: 14px; color: #555;">ผู้พัฒนาระบบ: นายก้องนที อุ่นเจริญ (ตำแหน่ง ครู)</div>',
      exportOptions: { columns: ':not(:last-child)' }
    }],
    columns: [
      { title: "#", data: null, render: function(d, t, r, m) { return m.row + 1; }, className: 'text-center' },
      { title: "รหัส/โครงการ", data: null, render: function(d, t, row) { return row[1] + ' ' + row[2]; } },
      { title: "จำนวนเงินที่เบิก", data: 3, render: function(d) { return Number(d).toLocaleString(); }, className: 'text-end' },
      { title: "ครั้งที่เบิก", data: 5, className: 'text-center' },
      { title: "เงินคงเหลือ", data: 6, render: function(d) { return Number(d).toLocaleString(); }, className: 'text-end' },
      { title: "Action", data: 0, orderable: false, className: 'text-center', render: function(data, type, row) { 
          const hasAttachments = (row[7] && row[7] !== "[]") || (row[8] && row[8] !== "[]");
          let editBtn = '<button class="btn btn-sm btn-warning" onclick="editTransactions(\'' + data + '\')" title="แก้ไข"><i class="fas fa-pen"></i></button>';
          let deleteBtn = '<button class="btn btn-sm btn-danger ms-1" onclick="deleteTransactions(\'' + data + '\')" title="ลบ"><i class="fas fa-trash"></i></button>';
          let uploadBtn = '<button class="btn btn-sm ' + (hasAttachments ? 'btn-success' : 'btn-info') + ' ms-1" onclick="openUploadModal(\'' + data + '\')" title="แนบไฟล์"><i class="fas fa-paperclip"></i></button>';
          if (userRole === 'Admin') return editBtn + deleteBtn + uploadBtn;
          if (userRole === 'Staff') return uploadBtn; 
          return 'N/A'; 
      }}
    ],
    language: { url: 'https://cdn.datatables.net/plug-ins/1.11.3/i18n/th.json' }
  });
}

function deleteTransactions(txId) {
  if (userRole !== 'Admin') return toastr.error("⚠️ คุณไม่มีสิทธิ์ลบรายการ");
  if (!requireLogin()) return;
  const transaction = latestTransactions.find(function(tx) { return tx[0] === txId; }); 
  const projectCode = transaction ? transaction[1] : '';
  const projectName = transaction ? transaction[2] : '(ไม่ทราบชื่อ)';

  Swal.fire({
    title: 'ลบรายการ',
    html: '<p>ต้องการลบรายการของ ' + projectName + ' นี้หรือไม่?</p>',
    icon: 'warning', showCancelButton: true, confirmButtonText: 'ใช่,ลบ', confirmButtonColor: 'var(--danger-color)'
  }).then(function(result) { 
    if (result.isConfirmed) {
      loadingStart();
      callGoogleScript('deleteTransactionById', { token: authToken, txId: txId })
        .then(function(resp) {
          loadingEnd();
          if (resp.success) { Swal.fire('สำเร็จ', 'ลบเรียบร้อยแล้ว', 'success'); loadProjectSummary(); loadDataTransactions(); } 
          else { Swal.fire('ผิดพลาด', resp.message, 'error'); }
        }).catch(onFailure);
    }
  });
}

function editTransactions(txId) {
  if (userRole !== 'Admin' && userRole !== 'Staff') return toastr.error("⚠️ คุณไม่มีสิทธิ์แก้ไข");
  if (!requireLogin()) return;
  loadingStart();
  callGoogleScript('getTransactionById', { token: authToken, txId: txId })
    .then(function(tx) {
      loadingEnd();
      if (!tx) return Swal.fire('ผิดพลาด', 'ไม่พบรายการ', 'error');
      const hasAttachments = (tx.receipts && tx.receipts !== "[]") || (tx.reports && tx.reports !== "[]");
      const isZeroAmountTx = parseFloat(tx.amount) === 0;

      Swal.fire({
        title: 'แก้ไขรายการ',
        html: '<p><strong>โครงการ:</strong> ' + tx.projectName + '</p>' +
            '<input type="number" id="swal-input-amount" class="swal2-input" value="' + tx.amount + '" ' + (hasAttachments || isZeroAmountTx ? 'disabled' : '') + '>',
        showCancelButton: true, confirmButtonText: 'บันทึก',
        preConfirm: function() { 
          if (hasAttachments || isZeroAmountTx) return parseFloat(tx.amount);
          const newAmt = parseFloat(document.getElementById('swal-input-amount').value);
          if (isNaN(newAmt) || newAmt < 0 || newAmt === 0) { Swal.showValidationMessage('กรุณากรอกจำนวนเงินให้ถูกต้อง (ห้ามเป็น 0)'); return false; }
          return newAmt;
        }
      }).then(function(res) { 
        if (res.isConfirmed && res.value !== undefined) {
          if (res.value === parseFloat(tx.amount)) return toastr.info("ไม่มีการเปลี่ยนแปลง");
          loadingStart();
          callGoogleScript('updateTransaction', { token: authToken, id: tx.id, newAmount: res.value })
            .then(function(resp) {
              loadingEnd();
              if (resp.success) { Swal.fire('สำเร็จ', resp.message, 'success'); loadProjectSummary(); loadDataTransactions(); } 
              else { Swal.fire('ผิดพลาด', resp.message, 'error'); }
            }).catch(onFailure);
        }
      });
    }).catch(onFailure);
}

// ==============================
// Dashboard & Summary
// ==============================
function loadProjectSummary() {
  loadingStart();
  const tokenToSend = isLoggedIn ? authToken : null;
  callGoogleScript('getProjectSummary', { token: tokenToSend })
    .then(function(res) {
      loadingEnd();
      if (res.success) {
        summaryData = res.data; 
        const filterSelect = document.getElementById('filterProject');
        const currentProjectVal = filterSelect.value; 
        filterSelect.innerHTML = '<option value="">-- แสดงทั้งหมด --</option>';
        summaryData.forEach(function(item) { 
          const opt = document.createElement('option'); opt.value = item.code; opt.textContent = item.code + ' - ' + item.name; filterSelect.appendChild(opt);
        });
        filterSelect.value = currentProjectVal;
        filterSummaryTable(); 
        onProjectChange();
        if (userRole === 'Admin') renderAdminDashboard(summaryData);
      } else toastr.error("ไม่สามารถโหลดข้อมูลสรุปได้");
    }).catch(onFailure);
}

function filterSummaryTable() {
  const selectedCode = document.getElementById('filterProject').value;
  const selectedStatus = document.getElementById('filterStatus').value;
  let filteredData = summaryData;
  if (selectedCode) filteredData = filteredData.filter(function(item) { return item.code === selectedCode; });
  if (selectedStatus) filteredData = filteredData.filter(function(item) { return item.status === selectedStatus; });
  renderSummaryTable(filteredData);
  renderSummaryDashboard(filteredData);
}

function getStatusBadge(status) {
  if (status === "เสร็จสิ้น") return '<span class="badge bg-success-subtle text-success-emphasis rounded-pill">' + status + '</span>';
  if (status === "กำลังดำเนินการ") return '<span class="badge bg-warning-subtle text-warning-emphasis rounded-pill">' + status + '</span>';
  return '<span class="badge bg-secondary-subtle text-secondary-emphasis rounded-pill">' + status + '</span>';
}

function renderSummaryTable(data) {
  const tableId = '#summaryTable';
  if ($.fn.DataTable.isDataTable(tableId)) $(tableId).DataTable().destroy();
  const tbody = document.querySelector(tableId + ' tbody'); 
  tbody.innerHTML = ''; 
  if (!data || data.length === 0) { tbody.innerHTML = '<tr><td colspan="5" class="text-center">ไม่พบข้อมูล</td></tr>'; return; }

  data.forEach(function(item, index) { 
    const tr = document.createElement('tr');
    tr.innerHTML = '<td class="text-center">' + (index + 1) + '</td><td>' + item.code + ' - ' + item.name + '</td><td class="text-center">' + getStatusBadge(item.status) + '</td><td class="text-center">' + item.txCount + '</td><td class="text-center">' + (item.txCount > 0 ? (item.txWithFiles + ' / ' + item.txCount) : '-') + '</td>';
    tbody.appendChild(tr);
  });

  new DataTable(tableId, {
    destroy: true, responsive: true, pageLength: 10, order: [[0, 'asc']],
    dom: '<"d-flex justify-content-between align-items-center mb-3"Bf>rt<"d-flex justify-content-between mt-3"ip>',
    buttons: [{
      extend: 'print', text: '<i class="fa-solid fa-print me-1"></i> พิมพ์สรุป', className: 'btn btn-primary btn-sm rounded-pill px-3', 
      title: 'สรุปการดำเนินการ โรงเรียนมหาชัยพิทยาคาร', 
      messageBottom: '<div style="text-align: right; margin-top: 30px; font-size: 14px; color: #555;">ผู้พัฒนาระบบ: นายก้องนที อุ่นเจริญ (ตำแหน่ง ครู)</div>',
      exportOptions: { columns: ':visible' }
    }],
    language: { url: 'https://cdn.datatables.net/plug-ins/1.11.3/i18n/th.json' }
  });
}

function renderSummaryDashboard(data) {
  let completed = 0, inProgress = 0, notStarted = 0;
  data.forEach(function(item) {
    if (item.status === "เสร็จสิ้น") completed++;
    else if (item.status === "กำลังดำเนินการ") inProgress++;
    else notStarted++;
  });
  document.getElementById('totalProjects').textContent = data.length;
  document.getElementById('totalCompleted').textContent = completed;
  document.getElementById('totalInProgress').textContent = inProgress;
  document.getElementById('totalNotStarted').textContent = notStarted;
}

function renderAdminDashboard(data) {
  const dashboardPanel = document.getElementById('admin-dashboard-panel');
  if (!dashboardPanel) return;
  if (!['Admin', 'Staff', 'Viewer'].includes(userRole)) { dashboardPanel.classList.add('d-none'); return; }
  dashboardPanel.classList.remove('d-none');

  let totalBudget = 0, totalUsed = 0;
  if (data && data.length > 0) {
    data.forEach(function(item) { totalBudget += parseFloat(item.budget) || 0; totalUsed += parseFloat(item.used) || 0; });
  }
  const totalBalance = totalBudget - totalUsed;
  const utilization = totalBudget > 0 ? (totalUsed / totalBudget) * 100 : 0;
  const fmt = new Intl.NumberFormat('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  
  document.getElementById('dash-total-budget').textContent = fmt.format(totalBudget);
  document.getElementById('dash-total-used').textContent = fmt.format(totalUsed);
  document.getElementById('dash-total-balance').textContent = fmt.format(totalBalance);
  document.getElementById('dash-utilization').textContent = utilization.toFixed(2);
  const progressBar = document.getElementById('dash-progress-bar');
  if (progressBar) { progressBar.style.width = Math.min(utilization, 100) + '%'; progressBar.setAttribute('aria-valuenow', utilization); }

  const ctx = document.getElementById('budgetChart');
  if (!ctx) return;
  if (budgetChartInstance) budgetChartInstance.destroy(); 
  budgetChartInstance = new Chart(ctx, {
    type: 'doughnut',
    data: { labels: ['ใช้ไปแล้ว', 'คงเหลือ'], datasets: [{ data: [totalUsed, totalBalance], backgroundColor: ['#ffc107', '#198754'], borderWidth: 0 }] },
    options: { responsive: true, maintainAspectRatio: false, cutout: '60%', plugins: { legend: { position: 'bottom' } } }
  });
}

// ==============================
// Admin CRUD (Projects & Users)
// ==============================
function loadProjectData() {
  if (userRole !== 'Admin' || !requireLogin()) return;
  loadingStart();
  callGoogleScript('getProjectData', { token: authToken })
    .then(function(results) { loadingEnd(); showProjectTable(results); })
    .catch(onFailure);
}

function showProjectTable(items) {
  const tableId = '#tableProjects';
  if ($.fn.DataTable.isDataTable(tableId)) $(tableId).DataTable().destroy();
  if (!items || items.length === 0) { $(tableId).html("<thead><tr><th>ที่</th><th>รหัส</th><th>ชื่อโครงการ</th><th>งบประมาณ</th><th>ผู้รับผิดชอบ</th><th>การจัดการ</th></tr></thead><tbody><tr><td colspan='6' class='text-center'>ไม่พบข้อมูล</td></tr></tbody>"); return; }
  new DataTable(tableId, {
    destroy: true, responsive: true, pageLength: 10, data: items, order: [[0, 'asc']],
    dom: '<"d-flex justify-content-between align-items-center mb-3"Bf>rt<"d-flex justify-content-between mt-3"ip>',
    buttons: [{
      extend: 'print', text: '<i class="fa-solid fa-print me-1"></i> พิมพ์ตาราง', className: 'btn btn-primary btn-sm rounded-pill px-3', title: 'ภาพรวมโครงการ โรงเรียนมหาชัยพิทยาคาร',
      exportOptions: { columns: ':not(:last-child)' }
    }],
    columns: [
      { title: "ที่", data: null, render: function(d, t, r, m) { return m.row + 1; }, className: 'text-center' },
      { title: "รหัสโครงการ", data: 0, className: 'text-center' },
      { title: "ชื่อโครงการ", data: 1 },
      { title: "งบประมาณ", data: 2, render: function(data) { return Number(data || 0).toLocaleString(); }, className: 'text-end' },
      { title: "ผู้รับผิดชอบ", data: 3, render: function(data) { const ownerIds = _parseOwnerIds(data); return ownerIds.map(getUserFullName).join(', ') || '(ไม่ได้ระบุ)'; } },
      { title: "การจัดการ", data: 0, orderable: false, className: 'text-center', render: function(data) { 
          if (userRole !== 'Admin') return 'N/A';
          return '<button class="btn btn-sm btn-warning me-1" onclick="openProjectModal(\'edit\',\'' + data + '\')"><i class="fa-solid fa-pen-to-square"></i></button><button class="btn btn-sm btn-danger" onclick="deleteProjectConfirm(\'' + data + '\')"><i class="fa-solid fa-trash"></i></button>';
      }}
    ], language: { url: 'https://cdn.datatables.net/plug-ins/1.11.3/i18n/th.json' }
  });
}

function openProjectModal(mode, code = "") {
  if (userRole !== 'Admin' || !requireLogin()) return;
  const modal = new bootstrap.Modal(document.getElementById('projectModal'));
  $('#myFormAddProject')[0].reset(); $('#projectMode').val(mode); $('#projectId').val(code);

  if (tomSelectOwner) tomSelectOwner.destroy(); 
  tomSelectOwner = new TomSelect('#projectOwner', {
    valueField: 'id', labelField: 'fullName', searchField: ['fullName'], options: assignableUsers, create: false, placeholder: 'เลือกผู้รับผิดชอบ'
  });

  if (mode === "add") {
    $('#labelModalProjectModal').html('<i class="fa-solid fa-plus me-2"></i> เพิ่มโครงการ');
    $('#projectCode').prop('disabled', false); tomSelectOwner.setValue([]);
    modal.show();
  } else if (mode === "edit") {
    $('#labelModalProjectModal').html('<i class="fa-solid fa-pen-to-square me-2"></i> แก้ไขโครงการ');
    loadingStart();
    callGoogleScript('getProjectByCode', { token: authToken, code: code })
      .then(function(project) {
        loadingEnd();
        if (!project) return Swal.fire({ icon: 'error', title: 'ไม่พบโครงการนี้' });
        $('#projectCode').val(project.ProjectCode).prop('disabled', true);
        $('#projectName').val(project.ProjectName);
        $('#projectBudget').val(project.Budget);
        tomSelectOwner.setValue(_parseOwnerIds(project.Owner) || []); 
        modal.show();
      }).catch(onFailure);
  }
}

function saveProject() {
  if (userRole !== 'Admin' || !requireLogin()) return;
  const mode = $('#projectMode').val(); const code = $('#projectCode').val().trim(); const name = $('#projectName').val().trim(); const budget = $('#projectBudget').val().trim();
  const owner = tomSelectOwner.getValue(); 
  if (!code || !name || !budget) return Swal.fire({ icon: 'warning', title: 'กรุณากรอกข้อมูลให้ครบถ้วน' });

  loadingStart();
  const actionName = mode === "add" ? 'addProject' : 'updateProject';
  callGoogleScript(actionName, { token: authToken, code: code, name: name, budget: budget, owner: owner })
    .then(function(msg) {
      loadingEnd();
      const icon = msg.startsWith('✅') ? 'success' : 'error';
      Swal.fire({ icon: icon, title: msg });
      if (icon === 'success') {
        bootstrap.Modal.getInstance(document.getElementById('projectModal')).hide();
        loadProjectData(); loadProjects(); loadProjectSummary(); 
      }
    }).catch(onFailure);
}

function deleteProjectConfirm(code) {
  if (userRole !== 'Admin' || !requireLogin()) return;
  Swal.fire({ title: 'ลบโครงการ?', text: 'การดำเนินการนี้ไม่สามารถย้อนกลับได้', icon: 'warning', showCancelButton: true, confirmButtonColor: 'var(--danger-color)' })
  .then(function(result) {
    if (result.isConfirmed) {
      loadingStart();
      callGoogleScript('deleteProject', { token: authToken, code: code })
        .then(function(msg) {
          loadingEnd();
          if (msg.startsWith('🗑️')) { Swal.fire('สำเร็จ', msg, 'success'); loadProjectData(); loadProjects(); loadProjectSummary(); }
          else { Swal.fire('ผิดพลาด', msg, 'error'); }
        }).catch(onFailure);
    }
  });
}

function loadUsersData() {
  if (userRole !== 'Admin' || !requireLogin()) return;
  loadingStart();
  callGoogleScript('getUserData', { token: authToken })
    .then(function(results) { loadingEnd(); showUserTable(results); })
    .catch(onFailure);
}

function showUserTable(users) {
  const tableId = '#tableUsers';
  if ($.fn.DataTable.isDataTable(tableId)) $(tableId).DataTable().destroy();
  if (!users || users.length === 0) return $(tableId).html("<thead><tr><td class='text-center'>ไม่พบข้อมูล</td></tr></thead>");
  new DataTable(tableId, {
    destroy: true, responsive: true, data: users,
    columns: [
      { title: "#", data: null, render: function(d, t, r, m) { return m.row + 1; } },
      { title: "Username", data: "username" }, { title: "Full Name", data: "fullName" },
      { title: "Status", data: "status", render: function(d) { return d === 'active' ? '<span class="badge bg-success">Active</span>' : '<span class="badge bg-danger">Inactive</span>'; } },
      { title: "Role", data: "role" },
      { title: "Action", data: "id", orderable: false, render: function(data) { 
          if (userRole !== 'Admin') return 'N/A';
          return '<button class="btn btn-sm btn-warning me-1" onclick="editUser(\'' + data + '\')"><i class="fa-solid fa-pen"></i></button><button class="btn btn-sm btn-danger" onclick="deleteUserConfirm(\'' + data + '\')"><i class="fa-solid fa-trash"></i></button>';
      }}
    ]
  });
}

function openUserModal() {
  if (userRole !== 'Admin' || !requireLogin()) return;
  $("#formUser")[0].reset(); $("#userId").val(''); $("#roleUser").val('Viewer'); $("#status").val('active');
  $("#modalUserLabel").html('<i class="fas fa-user-plus me-2"></i> เพิ่มผู้ใช้งาน');
  new bootstrap.Modal(document.getElementById('modalUser')).show();
}

function editUser(id) {
  if (userRole !== 'Admin' || !requireLogin()) return;
  loadingStart();
  callGoogleScript('getUserById', { token: authToken, id: id })
    .then(function(user) {
      loadingEnd();
      if (!user) return Swal.fire('ไม่พบผู้ใช้งาน', '', 'error');
      $("#userId").val(user.id); $("#usernameUser").val(user.username); $("#passwordUser").val(user.password);
      $("#fullName").val(user.fullName); $("#status").val(user.status); $("#roleUser").val(user.role); 
      $("#modalUserLabel").html('<i class="fas fa-user-pen me-2"></i> แก้ไขผู้ใช้งาน');
      new bootstrap.Modal(document.getElementById('modalUser')).show();
    }).catch(onFailure);
}

function deleteUserConfirm(id) {
  if (userRole !== 'Admin' || !requireLogin()) return;
  Swal.fire({ title: 'แน่ใจหรือไม่?', icon: 'warning', showCancelButton: true, confirmButtonColor: 'var(--danger-color)' })
  .then(function(result) {
    if (result.isConfirmed) {
      loadingStart();
      callGoogleScript('deleteUser', { token: authToken, id: id })
        .then(function(msg) { loadingEnd(); Swal.fire('สำเร็จ', msg, 'success'); loadUsersData(); loadAssignableUsers(); })
        .catch(onFailure);
    }
  });
}

$("#formUser").submit(function (e) {
  e.preventDefault();
  if (userRole !== 'Admin' || !requireLogin()) return;
  const data = { id: $("#userId").val(), username: $("#usernameUser").val().trim(), password: $("#passwordUser").val(), fullName: $("#fullName").val(), status: $("#status").val(), role: $("#roleUser").val() };
  loadingStart();
  callGoogleScript('saveUserData', { token: authToken, userData: data }) // เปลี่ยนเป็น params.userData ตามที่ Code.gs รับ
    .then(function(msg) {
      loadingEnd();
      if(msg.includes('สำเร็จ')) { Swal.fire('สำเร็จ', msg, 'success'); bootstrap.Modal.getInstance(document.getElementById('modalUser')).hide(); loadUsersData(); loadAssignableUsers(); }
      else Swal.fire('ผิดพลาด', msg, 'error');
    }).catch(onFailure);
});

// ==============================
// การตั้งค่า & Excel
// ==============================
function loadSystemSettings() {
  if (userRole !== 'Admin' || !requireLogin()) return;
  loadingStart();
  callGoogleScript('getDriveFolderId', { token: authToken })
    .then(function(res) { loadingEnd(); if(res.success) $('#driveFolderId').val(res.folderId); })
    .catch(onFailure);
}

$("#formSettings").submit(function (e) {
  e.preventDefault();
  const newId = $('#driveFolderId').val().trim();
  if (!newId) return Swal.fire('เตือน', 'กรุณากรอก Folder ID', 'warning');
  Swal.fire({ title: 'ยืนยัน', icon: 'warning', showCancelButton: true }).then(function(res) {
    if(res.isConfirmed) {
      loadingStart();
      callGoogleScript('setDriveFolderId', { token: authToken, newId: newId })
        .then(function(r) { loadingEnd(); if(r.success) Swal.fire('สำเร็จ', r.message, 'success'); })
        .catch(onFailure);
    }
  });
});

function exportProjects() {
  if (userRole !== 'Admin' || !requireLogin()) return;
  loadingStart();
  callGoogleScript('getProjectData', { token: authToken })
    .then(function(data) {
      if (!data || data.length===0) { loadingEnd(); return toastr.info('ไม่พบข้อมูล'); }
      const workbook = new ExcelJS.Workbook(); const worksheet = workbook.addWorksheet('Export');
      worksheet.columns = [ { header: 'Code', key: 'code', width: 25 }, { header: 'Name', key: 'name', width: 40 }, { header: 'Budget', key: 'budget', width: 20 }, { header: 'Owner', key: 'owner', width: 40 } ];
      data.forEach(function(r) { worksheet.addRow({ code: r[0], name: r[1], budget: r[2], owner: r[3] }); });
      workbook.xlsx.writeBuffer().then(function(b) { saveAs(new Blob([b]), "Project_Export.xlsx"); loadingEnd(); });
    }).catch(onFailure);
}

function downloadProjectTemplate() {
  const workbook = new ExcelJS.Workbook(); const worksheet = workbook.addWorksheet('Template');
  worksheet.columns = [ { header: 'ProjectCode', key: 'code', width: 25 }, { header: 'ProjectName', key: 'name', width: 40 }, { header: 'Budget', key: 'budget', width: 20 } ];
  workbook.xlsx.writeBuffer().then(function(b) { saveAs(new Blob([b]), "Template.xlsx"); });
}

async function handleProjectImport(e) {
  const file = e.target.files[0]; if (!file) return;
  loadingStart();
  try {
    const workbook = new ExcelJS.Workbook(); await workbook.xlsx.load(await file.arrayBuffer());
    const worksheet = workbook.getWorksheet(1); const projectsToImport = [];
    worksheet.eachRow(function(row, rNum) {
      if (rNum === 1) return;
      const p = { code: String(row.getCell(1).value||'').trim(), name: String(row.getCell(2).value||'').trim(), budget: parseFloat(row.getCell(3).value)||0 };
      if (p.code && p.name) projectsToImport.push(p);
    });
    callGoogleScript('importProjects', { token: authToken, projectsToImport: projectsToImport })
      .then(function(res) {
        loadingEnd();
        if(res.success) { Swal.fire('สำเร็จ', 'นำเข้า ' + res.added + ' รายการ', 'success'); loadProjectData(); loadProjects(); loadProjectSummary(); }
      }).catch(onFailure);
  } catch(err) { onFailure(err); } finally { $(e.target).val(null); }
}

// ==============================
// ไฟล์แนบ (Batch Upload)
// ==============================
function setUploadMode(mode) {
  uploadMode = mode;
  document.querySelector('#modalUploadFiles .upload-modal-body').dataset.mode = mode;
  if (mode === 'edit') { document.getElementById('upload-footer-readonly').classList.add('d-none'); document.getElementById('upload-footer-edit').classList.remove('d-none'); } 
  else { document.getElementById('upload-footer-readonly').classList.remove('d-none'); document.getElementById('upload-footer-edit').classList.add('d-none'); }
}

function openUploadModal(txId) {
  if (!requireLogin()) return;
  currentUploadTxId = txId; setUploadMode('read'); stagedUploads = []; stagedDeletions = [];
  document.getElementById('uploadTxIdDisplay').textContent = txId; document.getElementById('preview-receipts').innerHTML = ''; document.getElementById('preview-reports').innerHTML = '';
  uploadModal.show(); loadingStart();

  callGoogleScript('getTransactionById', { token: authToken, txId: txId })
    .then(function(tx) {
      loadingEnd();
      if (!tx) { uploadModal.hide(); return Swal.fire('ผิดพลาด', 'ไม่พบ', 'error'); }
      document.getElementById('uploadTxProjectName').textContent = tx.projectName;
      if (tx.receipts) { try { JSON.parse(tx.receipts).forEach(function(id) { renderThumbnail(id, id, 'receipts', false, true, ''); }); } catch(e){} }
      if (tx.reports) { try { JSON.parse(tx.reports).forEach(function(id) { renderThumbnail(id, id, 'reports', false, true, ''); }); } catch(e){} }
    }).catch(onFailure);
}

function setupDropzones() {
  document.querySelectorAll('.dropzone').forEach(function(zone) {
    zone.addEventListener('dragover', function(e) { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', function(e) { e.preventDefault(); zone.classList.remove('dragover'); });
    zone.addEventListener('drop', function(e) { e.preventDefault(); zone.classList.remove('dragover'); if(uploadMode!=='edit') return; processFiles(e.dataTransfer.files, zone.dataset.fileType); });
    zone.addEventListener('click', function() { if(uploadMode!=='edit') return; const up = document.getElementById('fileUploader'); up.dataset.fileType = zone.dataset.fileType; up.click(); });
  });
  document.getElementById('fileUploader').addEventListener('change', function(e) { if(uploadMode!=='edit') return; processFiles(e.target.files, e.target.dataset.fileType); e.target.value = null; });
}

function processFiles(files, type) {
  for (const f of files) {
    if (!['image/jpeg','image/png','image/gif','application/pdf'].includes(f.type)) { toastr.error('ไฟล์ไม่รองรับ'); continue; }
    if (f.size > 10*1024*1024) { toastr.error('ไฟล์ใหญ่เกินไป'); continue; }
    stageFileForUpload(f, type);
  }
}

async function stageFileForUpload(file, type) {
  const tempId = 'temp-' + Date.now(); const isImage = file.type.startsWith('image/');
  const thumbEl = renderThumbnail(tempId, file.name, type, true, isImage, '');
  const loadDiv = document.createElement('div'); loadDiv.className = 'thumbnail-loading'; thumbEl.appendChild(loadDiv);
  try {
    const b64 = await base64Encode(file);
    stagedUploads.push({ tempId: tempId, base64Data: b64.split(',')[1], mimeType: file.type, fileName: file.name, fileType: type });
    if(isImage) { const img = thumbEl.querySelector('img'); if(img){ img.src = b64; img.style.display = 'block'; } }
    loadDiv.remove();
  } catch(e) { thumbEl.remove(); toastr.error('อ่านไฟล์ล้มเหลว'); }
}

function deleteFileClient(thumbEl) {
  if(uploadMode!=='edit') return;
  const id = thumbEl.dataset.fileId;
  if(thumbEl.dataset.isStaged === 'true') { stagedUploads = stagedUploads.filter(function(f){ return f.tempId !== id; }); thumbEl.remove(); } 
  else { stagedDeletions.push({ fileId: id, fileType: thumbEl.dataset.fileType }); thumbEl.style.display = 'none'; }
}

function cancelUploadChanges() {
  setUploadMode('read');
  document.querySelectorAll('.thumbnail[data-is-staged="true"]').forEach(function(t){ t.remove(); });
  document.querySelectorAll('#modalUploadFiles .thumbnail[style*="display: none"]').forEach(function(t){ t.style.display = 'flex'; });
  stagedUploads = []; stagedDeletions = [];
}

function saveFileChangesClient() {
  loadingStart();
  const uploads = stagedUploads.map(function(f){ return { base64Data: f.base64Data, mimeType: f.mimeType, fileName: f.fileName, fileType: f.fileType }; });
  callGoogleScript('saveFileChanges', { token: authToken, txId: currentUploadTxId, stagedUploads: uploads, stagedDeletions: stagedDeletions })
    .then(function(res) {
      loadingEnd();
      if(res.success) { uploadModal.hide(); Swal.fire('สำเร็จ', res.message, 'success'); loadDataTransactions(); loadProjectSummary(); }
      else Swal.fire('ผิดพลาด', res.message, 'error');
    }).catch(onFailure);
}

function renderThumbnail(id, name, type, isStaged, isImg, url) {
  const c = document.getElementById('preview-' + type); const t = document.createElement('div'); t.className = 'thumbnail'; t.dataset.fileId = id; t.dataset.fileType = type; t.dataset.isStaged = isStaged;
  let click = 'style="cursor:default"';
  if(!isStaged) { name = id.substring(0,10)+'...'; url = 'https://lh3.googleusercontent.com/d/'+id; click = 'onclick="window.open(\'https://drive.google.com/file/d/'+id+'/view\', \'_blank\')"'; }
  t.innerHTML = isImg ? '<img src="'+url+'" '+(isStaged&&!url?'style="display:none;"':'')+'>' : '<i class="fa-solid fa-file-pdf text-danger fa-3x mt-2"></i>';
  t.innerHTML += '<span class="file-info" '+click+'>'+name+'</span>';
  const del = document.createElement('button'); deleteBtn.className = 'delete-file'; deleteBtn.innerHTML = '×'; deleteBtn.onclick = function(e){ e.stopPropagation(); deleteFileClient(t); }; t.appendChild(del);
  c.appendChild(t); return t;
}

function base64Encode(file) { return new Promise(function(res, rej){ const r = new FileReader(); r.readAsDataURL(file); r.onload = function(){ res(r.result); }; r.onerror = rej; }); }
