<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1" />
<title>點名冊</title>
<link rel="manifest" href="./manifest.webmanifest">
<meta name="theme-color" content="#173B31">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="點名冊">
<link rel="apple-touch-icon" href="./icon-192.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Klee+One:wght@400;600&family=Noto+Sans+TC:wght@400;500;700;900&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root{
    --navy:#123A5A;
    --navy-2:#194A6D;
    --navy-3:#245B7D;
    --blue-bg:#DCEBF4;
    --blue-panel:#EAF4F8;
    --turq:#63C8C3;
    --turq-soft:#DDF3F1;
    --yellow:#F4CF55;
    --yellow-soft:#FFF2B8;
    --white:#F9FCFD;
    --ink:#183A50;
    --muted:#678191;
    --red:#D7655D;
    --red-soft:#FBECE9;
    --line:#C9DDE6;
    --radius:18px;
  }

  *{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
  html,body{margin:0;padding:0}
  body{
    min-height:100vh;
    color:var(--ink);
    font-family:'Noto Sans TC','Klee One',sans-serif;
    background:var(--blue-bg);
    padding-bottom:30px;
  }

  /* No decorative circles / background ornaments */
  body::before{display:none}

  .wrap{
    position:relative;z-index:1;
    max-width:720px;margin:0 auto;padding:14px 14px 8px;
  }

  header.top{
    display:flex;align-items:center;justify-content:space-between;gap:12px;
    min-height:70px;
    padding:12px 15px;
    margin-bottom:16px;
    border:2px solid var(--navy);
    border-radius:19px;
    background:var(--turq);
    box-shadow:0 4px 0 rgba(18,58,90,.16);
  }

  header.top .brand{
    font-size:25px;font-weight:900;letter-spacing:1px;color:var(--navy);
    line-height:1.1;
  }

  /* Remove English subtitle entirely */
  header.top .brand small{display:none}

  .icon-btn{
    background:var(--white);color:var(--navy);
    border:2px solid var(--navy);border-radius:11px;
    padding:8px 11px;font-size:12px;font-weight:800;font-family:inherit;
    cursor:pointer;box-shadow:0 3px 0 rgba(18,58,90,.18);
  }
  .icon-btn:active{transform:translateY(2px);box-shadow:0 1px 0 rgba(18,58,90,.18)}

  .screen{display:none;animation:fade .2s ease}
  .screen.active{display:block}
  @keyframes fade{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}

  .course-card{
    display:flex;align-items:center;gap:10px;
    background:var(--blue-panel);
    border:2px solid var(--navy);
    border-radius:16px;
    padding:12px 14px;margin-bottom:10px;
    cursor:pointer;box-shadow:0 3px 0 rgba(18,58,90,.10);
  }
  .course-card .name{font-size:16px;font-weight:900;color:var(--navy)}
  .course-card .meta{font-size:11px;color:var(--muted);margin-top:3px}
  .course-card .go{color:var(--navy);font-size:22px;font-weight:900}
  .course-card .del{
    background:var(--red-soft);border:1px solid #E5A19B;color:var(--red);
    border-radius:8px;font-size:11px;font-weight:800;padding:5px 7px;
  }

  .empty-note{
    border:2px dashed #82B9CB;border-radius:17px;
    padding:25px 16px;text-align:center;color:var(--muted);
    font-size:13px;line-height:1.8;background:var(--blue-panel);
  }

  .primary-btn{
    width:100%;padding:12px 14px;border-radius:13px;
    border:2px solid var(--navy);
    background:var(--yellow);color:var(--navy);
    font-weight:900;font-size:14px;font-family:inherit;cursor:pointer;
    letter-spacing:.4px;box-shadow:0 3px 0 rgba(18,58,90,.18);
  }
  .primary-btn:active{transform:translateY(2px);box-shadow:0 1px 0 rgba(18,58,90,.18)}
  .primary-btn.ghost{background:var(--blue-panel);color:var(--navy)}
  .primary-btn.danger{background:var(--red);color:var(--white)}
  .primary-btn.line-send-btn{font-size:17px;padding:15px 14px;letter-spacing:1px}
  .primary-btn:disabled{opacity:.42;cursor:not-allowed;box-shadow:none}

  .stack{display:flex;flex-direction:column;gap:8px;margin-top:13px}
  .row2{display:flex;gap:8px}
  .row2>*{flex:1}

  label.field-label{
    display:block;font-size:11px;letter-spacing:1.5px;color:var(--navy);
    font-weight:900;margin:15px 0 7px;
  }
  input[type=text]{
    width:100%;padding:11px 12px;border-radius:11px;
    border:2px solid var(--line);background:var(--blue-panel);color:var(--ink);
    font-size:14px;font-family:inherit;
  }
  input[type=text]:focus{outline:none;border-color:var(--turq)}

  select{
    width:100%;padding:11px 12px;border-radius:11px;
    border:2px solid var(--line);background:var(--blue-panel);color:var(--ink);
    font-size:14px;font-family:inherit;
  }
  select:focus{outline:none;border-color:var(--turq)}
  .drive-picker{
    border:2px solid var(--navy);border-radius:14px;padding:12px 13px;
    background:var(--turq-soft);margin-bottom:14px;
  }

  .file-drop{
    border:2px dashed #78B7C6;border-radius:16px;
    padding:22px 13px;text-align:center;color:var(--muted);font-size:12px;
    background:var(--blue-panel);position:relative;line-height:1.7;
  }
  .file-drop input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer}
  .file-drop.has-file{border-style:solid;color:var(--navy);border-color:var(--turq);background:var(--turq-soft)}
  .hint{font-size:11px;color:var(--muted);line-height:1.75;margin-top:7px}
  .hint b{color:var(--navy)}

  .preview-table{
    margin-top:12px;border:2px solid var(--navy);border-radius:14px;
    overflow:hidden;background:var(--blue-panel)
  }
  .preview-row{
    display:flex;justify-content:space-between;padding:8px 12px;
    border-bottom:1px solid var(--line);font-size:12px
  }
  .preview-row:last-child{border-bottom:none}
  .preview-row .cls{color:var(--navy);font-weight:900}
  .preview-row .cnt{color:var(--muted)}
  .error-box{
    background:var(--red-soft);border:1px solid var(--red);color:#98453E;
    padding:10px 12px;border-radius:11px;font-size:12px;margin-top:10px;line-height:1.7
  }

  .course-header{
    background:var(--navy);color:var(--white);
    border-radius:16px;padding:13px 15px;margin-bottom:10px;
    border:2px solid var(--navy);box-shadow:0 4px 0 rgba(18,58,90,.13);
  }
  .course-header .title{font-size:19px;font-weight:900}
  .course-header .sub{font-size:11px;color:#D7E8EF;margin-top:3px}

  .summary-bar{
    display:flex;justify-content:space-between;align-items:center;
    background:var(--blue-panel);border:2px solid var(--navy);
    border-radius:12px;padding:8px 12px;margin:10px 0;
    font-size:12px;color:var(--muted);box-shadow:0 2px 0 rgba(18,58,90,.07)
  }
  .summary-bar b{color:var(--navy);font-size:15px}
  .summary-bar span:last-child b{color:var(--red)}

  /* Keep the current class visible while scrolling, without adding a complex sidebar. */
  /* One student per row: large enough for teachers with presbyopia. */
  .student-grid{grid-template-columns:1fr;gap:7px}
  .student-chip{gap:11px;padding:9px 10px;min-height:52px}
  .student-chip .seat{width:42px;height:42px;border-radius:9px;font-size:17px}
  .student-chip .seat::after{font-size:15px}
  .student-chip .nm{font-size:19px;font-weight:800;line-height:1.25;overflow:visible;text-overflow:clip;white-space:normal}
  .student-chip.absent .seat::after{font-size:21px}

  /* One class = one compact horizontal card; students flow in a tight grid. */
  #classContainer{
    display:flex;
    flex-direction:column;
    gap:10px;
    padding-top:4px;
  }

  .class-block{
    margin:0;
    background:var(--blue-panel);
    border:2px solid var(--navy);
    border-radius:16px;
    padding:10px;
    box-shadow:0 3px 0 rgba(18,58,90,.09);
    min-width:0;
  }

  .class-title{
    display:flex;align-items:center;gap:7px;margin-bottom:8px;
    padding:6px 9px;border-radius:10px;
    background:var(--turq);color:var(--navy);
    font-size:14px;font-weight:900;
  }
  .class-title .dash{display:none}

  .student-grid{
    display:flex;
    flex-direction:column;
    gap:8px;
  }

  .student-chip{
    display:flex;align-items:center;gap:12px;
    border:2px solid var(--line);border-radius:12px;padding:10px 12px;
    background:var(--white);cursor:pointer;user-select:none;min-height:56px;
    transition:.1s ease;
  }
  .student-chip:active{transform:scale(.98)}

  .student-chip .seat{
    width:44px;height:44px;border-radius:9px;flex:none;
    display:flex;align-items:center;justify-content:center;
    font-size:18px;font-weight:900;border:2px solid var(--navy);
    color:var(--navy);background:var(--yellow)
  }
  .student-chip .seat::after{content:"✓";font-size:15px}
  .student-chip .nm{
    flex:1;min-width:0;
    font-size:19px;line-height:1.3;font-weight:700;
    overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--ink)
  }
  .student-chip.absent{border-color:#E0A09A;background:var(--red-soft)}
  .student-chip.absent .seat{
    border-color:var(--red);color:var(--white);background:var(--red)
  }
  .student-chip.absent .seat::after{content:"×";font-size:22px}
  .student-chip.absent .nm{color:#98453E;text-decoration:line-through}

  .edit-roster-toggle.active{background:var(--red);color:var(--white);border-color:var(--navy)}
  .student-chip.edit-mode{cursor:default}
  .student-chip.edit-mode:active{transform:none}
  .student-chip .del-student{
    display:none;flex:none;width:36px;height:36px;border-radius:9px;
    border:2px solid var(--red);background:var(--red-soft);color:var(--red);
    font-size:18px;font-weight:900;align-items:center;justify-content:center;cursor:pointer;
  }
  .student-chip.edit-mode .del-student{display:flex}

  .session-status{
    padding:0 2px;margin:6px 0 2px;font-size:11px;line-height:1.7;
    font-weight:400;border:none;background:none;
  }
  .session-status.waiting{color:var(--muted)}
  .session-status.ready{color:var(--navy);font-weight:700}

  .period-picker{
    display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;
    background:var(--yellow-soft);border:3px solid var(--red);border-radius:14px;
    padding:12px 14px;margin-bottom:10px;
    box-shadow:0 4px 0 rgba(18,58,90,.15);
  }
  .period-picker .ppk-label{font-size:12px;font-weight:900;color:var(--navy);letter-spacing:.5px}
  .period-picker .ppk-row{display:flex;align-items:center;justify-content:center;gap:12px}
  .period-picker .ppk-value{font-size:24px;font-weight:900;color:var(--red);min-width:78px;text-align:center}
  .period-picker .ppk-btn{
    width:38px;height:38px;border-radius:9px;border:2px solid var(--navy);
    background:var(--white);color:var(--navy);font-size:18px;font-weight:900;
    font-family:inherit;cursor:pointer;
  }
  .period-picker .ppk-btn:active{transform:scale(.94)}

  .sticky-actions{
    position:sticky;bottom:0;padding:11px 0 4px;margin-top:11px;
    background:linear-gradient(to top,var(--blue-bg) 72%,transparent)
  }

  .stamp{position:relative;margin:5px auto 14px;width:fit-content}
  .stamp .seal{
    display:inline-block;border:2px solid var(--red);color:var(--red);
    background:var(--blue-panel);padding:5px 15px;border-radius:9px;
    font-size:13px;font-weight:900;transform:rotate(-3deg);letter-spacing:1px
  }
  .message-box{
    background:var(--blue-panel);color:var(--ink);
    border:2px solid var(--navy);border-radius:16px;
    padding:15px 14px;font-size:13px;line-height:1.9;white-space:pre-wrap;
    font-family:'Noto Sans TC','Klee One',sans-serif;
    box-shadow:0 4px 0 rgba(18,58,90,.10)
  }

  .toast{
    position:fixed;left:50%;bottom:24px;transform:translateX(-50%);
    background:var(--navy);color:var(--yellow);padding:9px 16px;
    border:1px solid var(--yellow);border-radius:25px;font-size:12px;font-weight:900;
    z-index:50;opacity:0;pointer-events:none;transition:opacity .2s
  }
  .toast.show{opacity:1}

  .loading-overlay{
    position:fixed;inset:0;z-index:80;
    display:flex;align-items:center;justify-content:center;
    background:rgba(18,58,90,.55);padding:20px;
  }
  .loading-overlay[hidden]{display:none}
  .loading-overlay .loading-box{
    background:var(--white);border:3px solid var(--navy);border-radius:18px;
    padding:26px 24px;text-align:center;max-width:320px;
    box-shadow:0 6px 0 rgba(18,58,90,.25);
  }
  .loading-overlay .loading-spinner{
    width:34px;height:34px;margin:0 auto 12px;border-radius:50%;
    border:5px solid var(--turq-soft);border-top-color:var(--navy);
    animation:spin .8s linear infinite;
  }
  @keyframes spin{to{transform:rotate(360deg)}}
  .loading-overlay .loading-text{font-size:18px;font-weight:900;color:var(--navy);line-height:1.6}
  a.line-link{text-decoration:none}


  .current-class-bar{
    position:sticky;top:0;z-index:25;
    display:flex;align-items:center;justify-content:space-between;gap:10px;
    width:100%;
    background:var(--red);color:var(--white);
    border:2px solid var(--navy);border-radius:12px;
    padding:8px 12px;margin:10px 0 10px;
    font-size:14px;font-weight:900;
    box-shadow:0 3px 0 rgba(18,58,90,.16);
  }
  .current-class-main{display:flex;align-items:baseline;gap:8px;min-width:0}
  .current-class-bar .label{font-size:11px;color:#FBECE9;font-weight:900;white-space:nowrap}
  .current-class-bar .value{font-size:20px;letter-spacing:.5px;white-space:nowrap}
  .current-class-meta{display:flex;align-items:center;gap:8px;white-space:nowrap;font-size:13px;font-weight:900}
  .current-class-meta .time{font-variant-numeric:tabular-nums}
  .current-class-meta .period{font-size:10px;font-weight:400;opacity:.8}
  .class-title{background:var(--turq);}
  .auto-course-name{
    width:100%;padding:12px 13px;border-radius:12px;border:2px solid var(--navy);
    background:var(--yellow-soft);color:var(--navy);font-size:16px;font-weight:900;
    font-family:inherit;
  }
  .auto-course-name::placeholder{color:var(--muted);font-weight:700}
  .google-settings-card .big-sub{font-size:12px;line-height:1.8}
  .google-settings-card .big-sub b{color:var(--navy)}
  .guide-step{display:grid;grid-template-columns:30px 1fr;gap:9px;align-items:start;margin-top:13px}
  .guide-num{width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;
    background:var(--navy);color:var(--white);font-weight:900;font-size:14px}
  .guide-step b{font-size:13px;color:var(--navy)}
  .guide-text{font-size:11px;color:var(--muted);line-height:1.75;margin-top:2px}
  .setup-details{margin-top:14px;border-top:2px solid var(--line);padding-top:10px}
  .setup-details summary{cursor:pointer;color:var(--navy);font-weight:900;font-size:13px}
  .setup-body{padding-top:2px}
  .guide-step.compact{margin-top:11px}
  .small-btn{margin-top:12px;padding:9px 11px;font-size:12px}
  .guide-code-note{font-size:11px;color:var(--muted);margin-top:5px}
  .apps-script-code{display:none;white-space:pre-wrap;max-height:180px;overflow:auto;margin:9px 0 0;padding:10px;
    border-radius:10px;background:#16384E;color:#F9FCFD;font:10px/1.5 ui-monospace,SFMono-Regular,Consolas,monospace}
  .setup-details.show-code .apps-script-code{display:block}
  .auto-google-note{font-size:11px;color:var(--muted);text-align:center;line-height:1.6}
  .auto-google-note.ready{color:var(--navy);font-weight:800}
  .auto-google-note.off{color:var(--muted)}

  .settings-card{
    background:var(--blue-panel);border:2px solid var(--navy);border-radius:16px;
    padding:13px 14px;margin-bottom:12px;
  }
  .settings-title{font-size:18px;font-weight:900;color:var(--navy);margin-bottom:4px}
  .settings-sub{font-size:11px;color:var(--muted);line-height:1.7}
  .period-row{
    display:grid;grid-template-columns:48px 1fr 12px 1fr;align-items:center;gap:7px;
    margin-top:8px;
  }
  .period-row .pname{font-size:14px;font-weight:900;color:var(--navy)}
  .period-row input[type=time]{
    width:100%;padding:8px 7px;border-radius:9px;border:2px solid var(--line);
    background:var(--white);color:var(--ink);font:700 14px inherit;
  }
  .period-row input[type=time]:focus{outline:none;border-color:var(--turq)}
  .settings-divider{height:1px;background:var(--line);margin:14px 0}
  .google-note{font-size:11px;line-height:1.7;color:var(--muted);margin-top:6px}
  .google-note b{color:var(--navy)}
  .account-warning{margin-top:10px;padding:11px 12px;border-radius:11px;background:var(--red-soft);border:2px solid var(--red);color:#7F3833;font-size:12px;line-height:1.75}
  .google-test-result{margin-top:8px;font-size:12px;line-height:1.7;font-weight:800}
  .google-test-result.ok{color:#237A55}
  .google-test-result.err{color:#A33F38}
  .save-google-btn.saving{opacity:.65}

  @media(max-width:560px){
    .wrap{padding:10px 10px 6px}
    header.top{min-height:64px;border-radius:16px;padding:10px 12px}
    header.top .brand{font-size:22px}
    .row2{flex-direction:row}
    .student-chip .nm{font-size:18px}
    .current-class-bar{padding:8px 10px;margin:8px 0}
    .current-class-bar .value{font-size:18px}
    .current-class-meta{gap:5px;font-size:12px}
    .current-class-meta .period{font-size:11px}
  }
</style>

<script>
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').catch(console.error);
  });
}
</script>
</head>
<body>
<div class="wrap">

  <header class="top">
    <div class="brand">跨群選修/本土語 點名</div>
    <div style="display:flex;gap:7px;align-items:center"><button class="icon-btn" id="settingsBtn">設定</button><button class="icon-btn" id="homeBtn">課程列表</button></div>
  </header>

  <!-- HOME -->
  <section class="screen" id="screen-home">
    <div id="courseList"></div>
    <div class="stack">
      <button class="primary-btn" id="newCourseBtn">＋ 新增課程（上傳或下拉選擇名單）</button>
    </div>
  </section>

  <!-- UPLOAD -->
  <section class="screen" id="screen-upload">
    <div class="course-header">
      <div class="title">建立課程名單</div>
      <div class="sub">上傳試算表，設定課程與班級座號名單</div>
    </div>

    <div class="drive-picker" id="drivePickerBox" hidden>
      <label class="field-label">從雲端硬碟選擇課程名單</label>
      <div class="hint" style="margin-top:-3px">此區為期初教務處統一釋出的名單</div>
      <select id="driveFileSelect">
        <option value="">請選擇檔案…</option>
      </select>
      <div class="hint" id="drivePickerHint">讀取中…</div>
    </div>

    <label class="field-label">或上傳試算表檔案（.xlsx / .csv）</label>
    <div class="hint" style="margin-top:-3px">此區可自行上傳，和雲端的名單互不干擾</div>
    <div class="file-drop" id="fileDrop">
      <div id="fileDropText">點擊或拖曳檔案到這裡上傳</div>
      <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" />
    </div>
    <div class="hint">
      表格需包含欄位：<b>班級</b>、<b>座號</b>（可選填 <b>姓名</b>）。<br>
      不同班級可各自有不重複的座號，系統會依班級自動分組。
    </div>

    <div id="uploadError"></div>
    <div id="uploadPreview"></div>

    <label class="field-label">課程名稱</label>
    <input type="text" class="auto-course-name" id="courseNameInput" placeholder="選好上面的來源後，會自動帶入名稱" />
    <div class="hint">例如：<b>多媒體概論.xlsx</b> → 課程名稱會自動變成「多媒體概論」，也可以自行修改成其他名稱。</div>

    <div class="stack">
      <button class="primary-btn" id="saveCourseBtn" disabled>儲存課程並開始點名</button>
      <button class="primary-btn ghost" id="cancelUploadBtn">取消</button>
    </div>
  </section>

  <!-- SETTINGS -->
  <section class="screen" id="screen-settings">
    <div class="course-header">
      <div class="title">設定</div>
      <div class="sub">設定上課節次與 Google 試算表，資料會保存在這台裝置。</div>
    </div>

    <div class="settings-card">
      <div class="settings-title">訊息開頭文字</div>
      <div class="settings-sub">送出最終點名訊息時，會自動加在整段訊息最前面，可依需要自行修改或清空。</div>
      <input type="text" id="messagePrefixInput" placeholder="例如：@程程" />
    </div>

    <div class="settings-card">
      <div class="settings-title">進階同步設定</div>
      <div class="settings-sub">Google 試算表同步、課程名單雲端硬碟——這兩項設定好一次之後就不太需要再調整。</div>
      <button type="button" class="primary-btn ghost small-btn" id="openAdvancedSyncBtn" style="margin-top:10px">前往進階同步設定 →</button>
    </div>

    <div class="settings-card">
      <div class="settings-title">上課節次時間</div>
      <div class="settings-sub">程式會依現在時間，自動判斷目前是第 1～8 節。若目前時間不在任何區間，會顯示「非上課時段」。</div>
      <div id="periodSettings"></div>
    </div>

    <div class="stack">
      <button class="primary-btn" id="saveSettingsBtn">儲存設定</button>
      <button class="primary-btn ghost" id="backFromSettingsBtn">← 返回</button>
    </div>
  </section>

  <!-- SETTINGS: ADVANCED SYNC -->
  <section class="screen" id="screen-settings-sync">
    <div class="course-header">
      <div class="title">進階同步設定</div>
      <div class="sub">Google 試算表同步與課程名單雲端硬碟，設定好一次之後平常不太需要再調整。</div>
    </div>

    <div class="settings-card google-settings-card">
      <div class="settings-title">Google 試算表（選填）</div>
      <div class="settings-sub big-sub">如果你希望每次點名（含第一節、第二節）送出時，也自動把點名結果寫進 Google 試算表，請完成下面設定。<br><b>不設定也可以正常點名與產生 LINE 訊息。</b></div>

      <div class="account-warning"><b>⚠️ 請務必用你「私人的 Google 帳號」（自己的 Gmail，不是學校配發的帳號）來開這份試算表、部署 Apps Script。</b><br>學校帳號通常是 Google Workspace 網域帳號，校方管理員多半會限制或直接關閉「部署為任何人都能存取的網頁應用程式」這個功能，就算整個步驟都照著做，最後同步還是會沒有反應、抓不出原因。用自己的私人帳號做這件事才會成功。</div>
      <div class="account-warning"><b>💻 以下設定步驟請用電腦操作，手機無法完成部署。</b><br>第一次設定大約只要 5 分鐘，且只需要做一次；設定好以後，平常用手機點名、產生訊息完全不受影響，一樣正常同步。</div>

      <div class="guide-step">
        <div class="guide-num">1</div>
        <div><b>用「私人 Google 帳號」在電腦上開一份新的空白試算表</b><div class="guide-text">到 <a href="https://sheets.google.com" target="_blank" rel="noopener">sheets.google.com</a> 開一份空白試算表，當作點名紀錄本。之後每次點名會在右邊新增一欄，不會覆蓋以前的紀錄。一定要從這份試算表裡面開 Apps Script（下一步），不要另外跑去 script.google.com 開一個獨立專案——獨立專案抓不到「目前是哪一份試算表」，程式會直接執行失敗。</div></div>
      </div>

      <label class="field-label">① Google 試算表網址</label>
      <input type="text" id="googleSheetUrlInput" placeholder="https://docs.google.com/spreadsheets/d/..." />

      <details class="setup-details">
        <summary>第一次使用？按這裡看完整設定方法（電腦操作，約 5 分鐘）</summary>
        <div class="setup-body">
          <div class="guide-step compact"><div class="guide-num">2</div><div><b>在剛剛那份試算表裡，上方選單點「擴充功能」→「Apps Script」</b><div class="guide-text">會開啟一個程式碼編輯畫面。</div></div></div>
          <div class="guide-step compact"><div class="guide-num">3</div><div><b>把畫面裡原本的內容全部刪掉</b><div class="guide-text">改貼上下面這一段程式碼（可以按「複製程式碼」按鈕）。</div></div></div>
          <button type="button" class="primary-btn ghost small-btn" id="copyAppsScriptBtn">複製 Google 寫入程式</button>
          <div class="guide-code-note">複製後，貼到 Apps Script，按上面工具列的「儲存」（磁片圖示），隨便取個名字。</div>
          <pre id="appsScriptCode" class="apps-script-code"></pre>
          <div class="guide-step compact"><div class="guide-num">4</div><div><b>按「部署」→「新增部署作業」</b><div class="guide-text">「選取類型」點齒輪圖示，選「網頁應用程式」。</div></div></div>
          <div class="guide-step compact"><div class="guide-num">5</div><div><b>「具有存取權的使用者」選「所有人」（Anyone），按「部署」</b><div class="guide-text">執行身分維持預設的「我」即可，不用更改。</div></div></div>
          <div class="guide-step compact"><div class="guide-num">6</div><div><b>如果跳出「需要授權」的畫面</b><div class="guide-text">選你自己的 Google 帳號 → 如果看到「Google 尚未驗證這個應用程式」的警告，點「進階」→「前往（專案名稱）（不安全）」→「允許」。這個警告是因為程式是你自己貼的、還沒有申請 Google 官方認證，不是真的有問題，自己貼的程式可以放心允許。</div></div></div>
          <div class="guide-step compact"><div class="guide-num">7</div><div><b>複製部署後顯示的「網頁應用程式」網址</b><div class="guide-text">網址通常會長得像：<code>https://script.google.com/macros/s/.../exec</code></div></div></div>
          <label class="field-label">② Google 寫入服務網址（只需設定一次）</label>
          <input type="text" id="googleLogUrlInput" placeholder="https://script.google.com/macros/s/.../exec" />
          <button type="button" class="primary-btn ghost small-btn" id="testGoogleBtn">🟢 測試 Google 連線</button>
          <div class="google-test-result" id="googleTestResult"></div>
          <div class="guide-text" style="margin-top:6px">按「測試連線」——如果看到一段確認文字，代表部署成功；如果看到錯誤畫面，回到第 4～6 步檢查部署設定（尤其是「具有存取權的使用者」是不是選「所有人」）。</div>
          <div class="google-note"><b>設定完成後：</b>平常點名時不需要再手動處理，每節點名送出時（含第一節、第二節），系統都會自動同步進 Google 試算表，不用再手動觸發。</div>
        </div>
      </details>
    </div>

    <div class="settings-card google-settings-card">
      <div class="settings-title">課程名單雲端硬碟（選填）</div>
      <div class="settings-sub big-sub">把各課程的名單檔案放在同一個雲端硬碟資料夾裡，設定好之後，「新增課程」畫面就能直接用下拉選單選課程，不用每次手動上傳檔案。這組設定跟上面的 Google 試算表同步是分開的，也可以把同一組網址分享給其他老師一起使用同一個資料夾。<br><b>已經預設填好一組共用的資料夾與讀取服務網址，可以直接使用，也可以換成自己的一組。不設定也可以正常用檔案上傳的方式建立課程。</b></div>

      <div class="account-warning"><b>⚠️ 一樣請用你「私人的 Google 帳號」</b>來建立資料夾、部署這支 Apps Script，理由跟前面的 Google 試算表同步相同——學校帳號常會被管理員限制，導致部署後沒有反應。</div>
      <div class="account-warning"><b>💻 以下設定步驟請用電腦操作，手機無法完成部署。</b>設定好以後，平常用手機選課程完全不受影響。</div>

      <div class="guide-step">
        <div class="guide-num">1</div>
        <div><b>在雲端硬碟建一個資料夾，把課程名單都放進去</b><div class="guide-text">名單檔案可以是 <b>Google 試算表</b>（推薦，讀取最穩定）或 <b>.xlsx 檔案</b>，內容格式跟本機上傳完全一樣：第一列要有「班級」「座號」欄位，姓名欄位可選填。</div></div>
      </div>

      <label class="field-label">① 雲端硬碟資料夾網址</label>
      <input type="text" id="driveFolderUrlInput" placeholder="https://drive.google.com/drive/folders/..." />

      <details class="setup-details" id="driveSetupDetails">
        <summary>第一次使用？按這裡看完整設定方法（電腦操作）</summary>
        <div class="setup-body">
          <div class="guide-step compact"><div class="guide-num">2</div><div><b>到 <a href="https://script.google.com" target="_blank" rel="noopener">script.google.com</a>，按「新增專案」</b><div class="guide-text">這次不用從某份試算表裡面開，直接開一個獨立專案就可以。</div></div></div>
          <div class="guide-step compact"><div class="guide-num">3</div><div><b>把畫面裡原本的內容全部刪掉</b><div class="guide-text">改貼上下面這一段程式碼（可以按「複製程式碼」按鈕）。</div></div></div>
          <button type="button" class="primary-btn ghost small-btn" id="copyDriveScriptBtn">複製雲端硬碟讀取程式</button>
          <div class="guide-code-note">複製後，貼到 Apps Script，按上面工具列的「儲存」（磁片圖示），隨便取個名字。</div>
          <pre id="driveAppsScriptCode" class="apps-script-code"></pre>
          <div class="guide-step compact"><div class="guide-num">4</div><div><b>資料夾裡如果有 .xlsx 檔案，要多加一個服務</b><div class="guide-text">左側選單「服務」旁的 ＋ 號 → 選「Drive API」→ 版本選「v3」→ 按「加入」。如果資料夾裡全部都是 Google 試算表，這一步可以跳過。</div></div></div>
          <div class="guide-step compact"><div class="guide-num">5</div><div><b>按「部署」→「新增部署作業」→「網頁應用程式」</b><div class="guide-text">「具有存取權的使用者」選「所有人」，按「部署」，並依畫面指示完成帳號授權（跟前面 Google 試算表同步的授權步驟一樣）。</div></div></div>
          <div class="guide-step compact"><div class="guide-num">6</div><div><b>複製部署後顯示的「網頁應用程式」網址</b><div class="guide-text">網址通常會長得像：<code>https://script.google.com/macros/s/.../exec</code></div></div></div>
          <label class="field-label">② 雲端硬碟讀取服務網址（只需設定一次）</label>
          <input type="text" id="driveLogUrlInput" placeholder="https://script.google.com/macros/s/.../exec" />
          <button type="button" class="primary-btn ghost small-btn" id="testDriveBtn">🟢 測試雲端硬碟連線</button>
          <div class="google-test-result" id="driveTestResult"></div>
          <div class="guide-text" style="margin-top:6px">按「測試連線」——如果看到資料夾名稱與檔案數量，代表部署成功；如果看到錯誤畫面，回到第 4～5 步檢查部署設定。</div>
          <div class="google-note"><b>設定完成後：</b>到「新增課程」畫面，就會多一個「從雲端硬碟選擇」的下拉選單可以直接選課程名單。</div>
        </div>
      </details>
    </div>

    <div class="stack">
      <button class="primary-btn" id="saveAdvancedSyncBtn">儲存設定</button>
      <button class="primary-btn ghost" id="backFromAdvancedSyncBtn">← 返回設定</button>
    </div>
  </section>

  <!-- ATTENDANCE -->
  <section class="screen" id="screen-attend">
    <div class="course-header">
      <div style="display:flex;align-items:center;justify-content:space-between;gap:10px">
        <div class="title" id="attendCourseName">—</div>
        <button type="button" class="icon-btn edit-roster-toggle" id="editRosterBtn">休學簡易刪除</button>
      </div>
      <div class="sub" id="attendSub">點擊學生標記為未出席，預設為全員出席；每節課都要各別點名一次，兩節點完才會送出紀錄</div>
    </div>
    <div class="summary-bar">
      <span>共 <b id="totalCount">0</b> 人</span>
      <span>未出席 <b id="absentCount" style="color:var(--ink-red-soft)">0</b> 人</span>
    </div>
    <div class="current-class-bar" id="currentClassBar" aria-live="polite">
      <div class="current-class-main">
        <span class="label">目前班級</span>
        <span class="value" id="currentClassName">—</span>
      </div>
      <div class="current-class-meta">
        <span class="time" id="liveClock">--:--</span>
        <span class="period" id="livePeriod">（自動偵測：第?節）</span>
      </div>
    </div>
    <div class="session-status" id="sessionStatusBar" hidden></div>
    <div id="classContainer" aria-label="依班級整合的學生名單"></div>

    <div class="sticky-actions stack">
      <div class="period-picker">
        <span class="ppk-label">✅ 送出前最後確認：本次登記節次</span>
        <div class="ppk-row">
          <button type="button" class="ppk-btn" id="periodPrevBtn" aria-label="往前一節">－</button>
          <span class="ppk-value" id="periodPickerValue">—</span>
          <button type="button" class="ppk-btn" id="periodNextBtn" aria-label="往後一節">＋</button>
        </div>
      </div>
      <button class="primary-btn" id="genMsgBtn">完成本節點名</button>
      <div class="auto-google-note" id="autoGoogleNote">已設定 Google 時，每節點名都會自動記錄。</div>
    </div>
  </section>

  <!-- RESULT -->
  <section class="screen" id="screen-result">
    <div class="stamp"><span class="seal">已產生 · 待送出</span></div>
    <div class="message-box" id="messageBox"></div>
    <div class="stack">
      <a class="line-link" id="lineShareLink" target="_blank" rel="noopener">
        <button class="primary-btn line-send-btn">傳送到 LINE 群組</button>
      </a>
      <button class="primary-btn ghost" id="copyBtn">複製訊息文字</button>
      <button class="primary-btn ghost" id="backToAttendBtn">← 返回修改點名</button>
    </div>
    <div class="hint" style="margin-top:14px;">
      「傳送到 LINE 群組」會開啟 LINE 並帶入訊息內容，需要你手動選擇要傳送的群組或聯絡人送出——網頁本身無法未經你操作就自動送出訊息給指定對象。
    </div>
  </section>

</div>
<div class="toast" id="toast"></div>
<div class="loading-overlay" id="loadingOverlay" hidden>
  <div class="loading-box">
    <div class="loading-spinner"></div>
    <div class="loading-text" id="loadingOverlayText">請稍等，雲端硬碟資料載入中…</div>
  </div>
</div>

<script>
(function(){
  const $ = (id) => document.getElementById(id);
  const screens = ['home','upload','attend','result','settings','settings-sync'];
  function showScreen(name){
    screens.forEach(s => $('screen-'+s).classList.toggle('active', s===name));
    window.scrollTo(0,0);
  }
  function toast(msg){
    const t = $('toast'); t.textContent = msg; t.classList.add('show');
    setTimeout(()=>t.classList.remove('show'), 1800);
  }

  // 讀取雲端硬碟時跳出的明顯提示視窗，讀取完成（不管成功或失敗）後要記得呼叫 hideLoading() 關掉。
  function showLoading(msg){
    $('loadingOverlayText').textContent = msg || '請稍等，資料載入中…';
    $('loadingOverlay').hidden = false;
  }
  function hideLoading(){
    $('loadingOverlay').hidden = true;
  }

  let currentCourseId = null;
  let currentRoster = null; // { name, classes: { className: [{seat,name}] }, pendingSession }
  let parsedUpload = null;  // classes object staged from file, before save
  let pendingCourseName = '';
  let rosterEditMode = false;
  let manualPeriod = null; // null = 跟隨自動偵測的節次；0~8 = 老師手動調整過的節次（0代表「非上課時段」）
  let lastCompletedSessions = null; // 最近一次「完成兩節點名」時，兩節各自的節次與缺席紀錄，供「返回修改點名」還原用
  let pendingEditSecond = null; // 從結果畫面按「返回修改點名」後，暫存第二節原本的紀錄，等第一節改完再自動帶入
  let lastSeenDate = null; // 用來偵測「跨到新的一天」，避免舊資料留到下次點名

  // ---------- storage helpers ----------
  // Browser persistence: IndexedDB. This works on HTTPS/PWA and keeps
  // each teacher's courses on their own device.
  const DB_NAME = 'attendance-ledger-db';
  const DB_VERSION = 1;
  const STORE = 'kv';

  function openDB(){
    return new Promise((resolve,reject)=>{
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        if(!req.result.objectStoreNames.contains(STORE)){
          req.result.createObjectStore(STORE);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function storageGet(key){
    try{
      const db = await openDB();
      return await new Promise((resolve,reject)=>{
        const tx = db.transaction(STORE,'readonly');
        const req = tx.objectStore(STORE).get(key);
        req.onsuccess = () => resolve(req.result ?? null);
        req.onerror = () => reject(req.error);
      });
    }catch(e){
      console.error(e);
      return null;
    }
  }

  async function storageSet(key, value){
    try{
      const db = await openDB();
      return await new Promise((resolve,reject)=>{
        const tx = db.transaction(STORE,'readwrite');
        tx.objectStore(STORE).put(value,key);
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => reject(tx.error);
      });
    }catch(e){
      console.error(e);
      toast('儲存失敗，請再試一次');
      return false;
    }
  }

  async function storageDelete(key){
    try{
      const db = await openDB();
      return await new Promise((resolve,reject)=>{
        const tx = db.transaction(STORE,'readwrite');
        tx.objectStore(STORE).delete(key);
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => reject(tx.error);
      });
    }catch(e){
      console.error(e);
      return false;
    }
  }

  async function getIndex(){
    try{
      const v = await storageGet('course-index');
      return v ? JSON.parse(v) : [];
    }catch(e){ return []; }
  }

  async function setIndex(list){
    await storageSet('course-index', JSON.stringify(list));
  }

  async function getRoster(id){
    try{
      const v = await storageGet('roster:'+id);
      return v ? JSON.parse(v) : null;
    }catch(e){ return null; }
  }

  async function setRoster(id, data){
    return await storageSet('roster:'+id, JSON.stringify(data));
  }

  async function deleteRoster(id){
    await storageDelete('roster:'+id);
  }

  // ---------- home screen ----------
  async function renderHome(){
    const list = await getIndex();
    const container = $('courseList');
    if(!list.length){
      container.innerHTML = '<div class="empty-note">還沒有任何課程。<br>點擊下方按鈕上傳第一份名單試算表。</div>';
      return;
    }
    container.innerHTML = '';
    list.slice().reverse().forEach(item => {
      const card = document.createElement('div');
      card.className = 'course-card';
      card.innerHTML = `
        <div style="flex:1; min-width:0;">
          <div class="name">${escapeHtml(item.name)}</div>
          <div class="meta">${item.classCount} 個班級・共 ${item.studentCount} 人</div>
        </div>
        <button type="button" class="del" data-id="${item.id}">刪除</button>
        <div class="go">›</div>
      `;
      card.addEventListener('click', (e) => {
        if(e.target.closest('.del')) return;
        openCourse(item.id);
      });
      card.querySelector('.del').addEventListener('click', async (e) => {
        e.preventDefault();
        e.stopPropagation();
        const ok = window.confirm(`確定要刪除課程「${item.name}」的名單嗎？\n\n刪除後將無法在此裝置上復原。`);
        if(!ok) return;
        const deleted = await deleteRoster(item.id);
        if(deleted === false){ toast('刪除失敗，請再試一次'); return; }
        const idx = await getIndex();
        await setIndex(idx.filter(c => c.id !== item.id));
        await renderHome();
        toast('課程已刪除');
      });
      container.appendChild(card);
    });
  }

  async function openCourse(id){
    const roster = await getRoster(id);
    if(!roster){ toast('讀取名單失敗'); return; }
    currentCourseId = id;
    currentRoster = roster;
    rosterEditMode = false;
    manualPeriod = null;
    lastCompletedSessions = null;
    pendingEditSecond = null;
    await checkDayRollover(); // 如果距離上次點名已經跨到新的一天，先把舊的暫存資料清乾淨
    // reset attendance state: everyone present by default
    Object.values(currentRoster.classes).forEach(students => {
      students.forEach(s => { s.absent = false; });
    });
    // 如果今天已經對目前這個節次輸入過紀錄（例如中途切出去又回來），直接帶入，避免看起來像資料不見了。
    const auto = await getCurrentPeriod();
    loadPeriodForDisplay(currentPickerValue(auto) || null);
    renderAttendance();
    showScreen('attend');
  }

  $('homeBtn').addEventListener('click', () => { renderHome(); showScreen('home'); });
  $('newCourseBtn').addEventListener('click', () => { initUploadScreen(); showScreen('upload'); });

  // ---------- upload screen ----------
  function resetUploadScreen(){
    pendingCourseName = '';
    $('courseNameInput').value = '';
    $('fileInput').value = '';
    $('fileDropText').textContent = '點擊或拖曳檔案到這裡上傳';
    $('fileDrop').classList.remove('has-file');
    $('uploadError').innerHTML = '';
    $('uploadPreview').innerHTML = '';
    $('saveCourseBtn').disabled = true;
    parsedUpload = null;
    if($('driveFileSelect')) $('driveFileSelect').value = '';
  }

  // 如果設定頁已經設好雲端硬碟資料夾，就抓取檔案清單填進下拉選單；沒設定就隱藏這個區塊。
  async function initUploadScreen(){
    resetUploadScreen();
    const box = $('drivePickerBox');
    const select = $('driveFileSelect');
    const hint = $('drivePickerHint');
    const folderUrl = (await getSetting(DRIVE_FOLDER_KEY, DEFAULT_DRIVE_FOLDER_URL)).trim();
    const webAppUrl = (await getSetting(DRIVE_LOG_KEY, DEFAULT_DRIVE_LOG_URL)).trim();
    if(!folderUrl || !webAppUrl){
      box.hidden = true;
      return;
    }
    box.hidden = false;
    select.innerHTML = '<option value="">讀取中…</option>';
    select.disabled = true;
    hint.textContent = '⏳ 正在讀取雲端硬碟資料夾內容…';
    try{
      const data = await jsonpRequest(webAppUrl, {action:'listFolder', folderUrl});
      if(data && data.ok){
        const files = data.files || [];
        select.innerHTML = '<option value="">請選擇檔案…</option>' + files.map(f =>
          `<option value="${escapeHtml(f.id)}" data-mime="${escapeHtml(f.mimeType)}" data-name="${escapeHtml(f.name)}">${escapeHtml(f.name)}</option>`
        ).join('');
        select.disabled = false;
        hint.textContent = files.length
          ? `資料夾「${data.folderName||''}」共 ${files.length} 個檔案，選好後仍可自行修改課程名稱。`
          : `資料夾「${data.folderName||''}」目前沒有找到課程名單檔案。`;
      }else{
        select.innerHTML = '<option value="">讀取失敗</option>';
        hint.textContent = '❌ 讀取失敗：'+(data && data.message ? data.message : '請到設定頁檢查雲端硬碟網址。');
      }
    }catch(e){
      select.innerHTML = '<option value="">讀取失敗</option>';
      hint.textContent = '❌ 無法連到雲端硬碟讀取服務，請到設定頁檢查設定或稍後再試。';
    }
  }

  $('cancelUploadBtn').addEventListener('click', () => { renderHome(); showScreen('home'); });

  function seatSortKey(seat){
    const m = String(seat).match(/\d+/);
    return m ? parseInt(m[0],10) : 999999;
  }

  function findCol(headerRow, keywords){
    for(let i=0;i<headerRow.length;i++){
      const cell = String(headerRow[i] ?? '').trim();
      if(keywords.some(k => cell.includes(k))) return i;
    }
    return -1;
  }

  function parseRows(rows){
    if(!rows.length) throw new Error('檔案內容是空的。');
    const header = rows[0];
    const clsCol = findCol(header, ['班級','班別','class']);
    const seatCol = findCol(header, ['座號','學號','seat','no']);
    const nameCol = findCol(header, ['姓名','名字','name']);
    if(clsCol === -1 || seatCol === -1){
      throw new Error('找不到「班級」或「座號」欄位，請確認試算表第一列是否為標題列。');
    }
    const classes = {};
    for(let r=1; r<rows.length; r++){
      const row = rows[r];
      if(!row || row.every(c => c===undefined || c==='')) continue;
      const cls = String(row[clsCol] ?? '').trim();
      const seat = String(row[seatCol] ?? '').trim();
      if(!cls || !seat) continue;
      const name = nameCol !== -1 ? String(row[nameCol] ?? '').trim() : '';
      if(!classes[cls]) classes[cls] = [];
      classes[cls].push({ seat, name, absent:false });
    }
    Object.values(classes).forEach(list => list.sort((a,b)=> seatSortKey(a.seat)-seatSortKey(b.seat)));
    if(!Object.keys(classes).length) throw new Error('沒有解析到任何學生資料，請確認內容是否正確。');
    return classes;
  }

  function handleFile(file){
    const reader = new FileReader();
    reader.onload = (e) => {
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:''});
        const classes = parseRows(rows);
        parsedUpload = classes;
        pendingCourseName = String(file.name || '').replace(/\.[^.]+$/, '').trim() || '未命名課程';
        $('courseNameInput').value = pendingCourseName;
        $('uploadError').innerHTML = '';
        if($('driveFileSelect')) $('driveFileSelect').value = '';
        renderPreview(classes);
        $('fileDropText').textContent = '✓ ' + file.name;
        $('fileDrop').classList.add('has-file');
        updateSaveEnabled();
      }catch(err){
        parsedUpload = null;
        $('uploadPreview').innerHTML = '';
        $('uploadError').innerHTML = `<div class="error-box">解析失敗：${escapeHtml(err.message)}</div>`;
        updateSaveEnabled();
      }
    };
    reader.readAsArrayBuffer(file);
  }

  $('driveFileSelect').addEventListener('change', async () => {
    const select = $('driveFileSelect');
    const fileId = select.value;
    const hint = $('drivePickerHint');
    if(!fileId) return;
    const opt = select.options[select.selectedIndex];
    const mimeType = opt.dataset.mime;
    const fileName = opt.dataset.name || opt.textContent;
    const webAppUrl = (await getSetting(DRIVE_LOG_KEY, DEFAULT_DRIVE_LOG_URL)).trim();
    hint.textContent = '⏳ 請等待雲端硬碟載入資料…';
    showLoading('請稍等，雲端硬碟資料載入中…');
    $('uploadError').innerHTML = '';
    try{
      const data = await jsonpRequest(webAppUrl, {action:'readFile', fileId, mimeType});
      if(data && data.ok){
        const classes = parseRows(data.rows || []);
        parsedUpload = classes;
        pendingCourseName = String(fileName || '').replace(/\.[^.]+$/, '').trim() || '未命名課程';
        $('courseNameInput').value = pendingCourseName;
        $('fileInput').value = '';
        $('fileDropText').textContent = '點擊或拖曳檔案到這裡上傳';
        $('fileDrop').classList.remove('has-file');
        renderPreview(classes);
        hint.textContent = `已載入「${fileName}」，可視需要修改課程名稱後儲存。`;
        updateSaveEnabled();
      }else{
        $('uploadError').innerHTML = `<div class="error-box">讀取失敗：${escapeHtml((data && data.message) || '無法讀取檔案內容。')}</div>`;
      }
    }catch(e){
      $('uploadError').innerHTML = `<div class="error-box">讀取失敗：無法連到雲端硬碟讀取服務，請稍後再試。</div>`;
    } finally {
      hideLoading();
    }
  });

  function renderPreview(classes){
    const box = $('uploadPreview');
    let total = 0;
    let html = '<div class="preview-table">';
    Object.keys(classes).forEach(cls => {
      total += classes[cls].length;
      html += `<div class="preview-row"><span class="cls">${escapeHtml(cls)}</span><span class="cnt">${classes[cls].length} 人</span></div>`;
    });
    html += '</div>';
    html += `<div class="hint">共 ${Object.keys(classes).length} 個班級、${total} 位學生</div>`;
    box.innerHTML = html;
  }

  $('fileInput').addEventListener('change', (e) => {
    if(e.target.files[0]) handleFile(e.target.files[0]);
  });
  $('courseNameInput').addEventListener('input', () => {
    pendingCourseName = $('courseNameInput').value;
    updateSaveEnabled();
  });
  function updateSaveEnabled(){
    $('saveCourseBtn').disabled = !(parsedUpload && $('courseNameInput').value.trim());
  }

  $('saveCourseBtn').addEventListener('click', async () => {
    const name = $('courseNameInput').value.trim();
    if(!name || !parsedUpload) return;
    const id = 'c' + Date.now();
    const rosterData = { name, classes: parsedUpload };
    const ok = await setRoster(id, rosterData);
    if(!ok){ toast('儲存失敗，請再試一次'); return; }
    const idx = await getIndex();
    const studentCount = Object.values(parsedUpload).reduce((n,l)=>n+l.length,0);
    const classCount = Object.keys(parsedUpload).length;
    const existingPos = idx.findIndex(c => c.id === id);
    const entry = { id, name, studentCount, classCount };
    if(existingPos >= 0) idx[existingPos] = entry; else idx.push(entry);
    await setIndex(idx);
    openCourse(id);
  });

  // ---------- attendance screen ----------
  let classScrollHandler = null;
  let classResizeHandler = null;

  function renderAttendance(){
    if(classScrollHandler) window.removeEventListener('scroll', classScrollHandler);
    if(classResizeHandler) window.removeEventListener('resize', classResizeHandler);
    $('attendCourseName').textContent = currentRoster.name;
    updateLiveClock();
    refreshSessionUI();

    const editBtn = $('editRosterBtn');
    editBtn.textContent = rosterEditMode ? '完成編輯' : '休學簡易刪除';
    editBtn.classList.toggle('active', rosterEditMode);

    const container = $('classContainer');
    container.innerHTML = '';
    Object.keys(currentRoster.classes)
      .sort((a,b) => String(a).localeCompare(String(b),'zh-Hant',{numeric:true}))
      .forEach(cls => {
        const block = document.createElement('div');
        block.className = 'class-block';
        block.dataset.className = cls;
        block.innerHTML = `<div class="class-title">${escapeHtml(cls)}<div class="dash"></div></div>
          <div class="student-grid"></div>`;
        const grid = block.querySelector('.student-grid');

        currentRoster.classes[cls].forEach(stu => {
          const chip = document.createElement('div');
          chip.className = 'student-chip' + (stu.absent ? ' absent' : '') + (rosterEditMode ? ' edit-mode' : '');
          chip.innerHTML = `<span class="seat">${escapeHtml(stu.seat)}</span><span class="nm">${escapeHtml(stu.name || '座號'+stu.seat)}</span><button type="button" class="del-student" aria-label="刪除這位學生">✕</button>`;
          chip.addEventListener('click', () => {
            if(rosterEditMode) return;
            stu.absent = !stu.absent;
            chip.classList.toggle('absent', stu.absent);
            updateSummary();
            if(currentCourseId && currentRoster) setRoster(currentCourseId, currentRoster);
          });
          chip.querySelector('.del-student').addEventListener('click', (e) => {
            e.stopPropagation();
            const label = stu.name ? `${cls} ${stu.seat}號 ${stu.name}` : `${cls} ${stu.seat}號`;
            const ok = window.confirm(`確定要從名單中刪除「${label}」嗎？\n\n這個動作無法復原（例如學生休學時使用）。`);
            if(!ok) return;
            removeStudent(cls, stu.seat);
          });
          grid.appendChild(chip);
        });
        container.appendChild(block);
      });

    const blocks = [...container.querySelectorAll('.class-block')];
    const setCurrentClass = (name) => { $('currentClassName').textContent = name || '—'; };

    function refreshCurrentClass(){
      if(!blocks.length) return;
      const bar = $('currentClassBar');
      const barBottom = bar.getBoundingClientRect().bottom + 8;
      let current = blocks[0];

      for(const block of blocks){
        if(block.getBoundingClientRect().top <= barBottom) current = block;
        else break;
      }
      setCurrentClass(current.dataset.className);
    }

    if(blocks.length){
      setCurrentClass(blocks[0].dataset.className);
      classScrollHandler = refreshCurrentClass;
      window.addEventListener('scroll', classScrollHandler, {passive:true});
      classResizeHandler = refreshCurrentClass;
      window.addEventListener('resize', classResizeHandler);
      refreshCurrentClass();
    }

    updateSummary();
  }

  async function removeStudent(cls, seat){
    if(!currentRoster || !currentRoster.classes[cls]) return;
    currentRoster.classes[cls] = currentRoster.classes[cls].filter(s => s.seat !== seat);
    if(!currentRoster.classes[cls].length) delete currentRoster.classes[cls];
    await setRoster(currentCourseId, currentRoster);
    const idx = await getIndex();
    const pos = idx.findIndex(c => c.id === currentCourseId);
    if(pos >= 0){
      idx[pos].studentCount = Object.values(currentRoster.classes).reduce((n,l)=>n+l.length,0);
      idx[pos].classCount = Object.keys(currentRoster.classes).length;
      await setIndex(idx);
    }
    renderAttendance();
    toast('已刪除該名學生');
  }

  $('editRosterBtn').addEventListener('click', () => {
    rosterEditMode = !rosterEditMode;
    renderAttendance();
  });

  function updateSummary(){
    let total=0, absent=0;
    Object.values(currentRoster.classes).forEach(list => {
      total += list.length;
      absent += list.filter(s=>s.absent).length;
    });
    $('totalCount').textContent = total;
    $('absentCount').textContent = absent;
  }


  // ---------- Settings / period / Google logging ----------
  const GOOGLE_SHEET_KEY = 'google-sheet-url';
  const GOOGLE_LOG_KEY = 'google-log-url';
  const DRIVE_FOLDER_KEY = 'drive-folder-url';
  const DRIVE_LOG_KEY = 'drive-log-url';
  const DEFAULT_DRIVE_FOLDER_URL = 'https://drive.google.com/drive/folders/1hCLx-IBTBdEbXhGz4NiXZmIcyr89WR53';
  const DEFAULT_DRIVE_LOG_URL = 'https://script.google.com/macros/s/AKfycbxNZpJtEQam_3LI2DQlIMm5bIE6YuLqpFiGu25cRZqRWU-CrUvI_6wJ_dxccAP6LUpisQ/exec';
  const PERIODS_KEY = 'attendance-periods';
  const MESSAGE_PREFIX_KEY = 'message-prefix';
  const DEFAULT_MESSAGE_PREFIX = '@程程';

  // 用 <script> 標籤呼叫 Apps Script 網頁應用程式並取回資料（JSONP 方式），
  // 因為 Apps Script 網頁應用程式的回應不支援一般 fetch 的跨網域讀取。
  function jsonpRequest(baseUrl, params){
    return new Promise((resolve, reject) => {
      if(!baseUrl){ reject(new Error('缺少服務網址')); return; }
      const cb = 'jsonp_cb_' + Date.now() + '_' + Math.floor(Math.random()*1e6);
      const script = document.createElement('script');
      const timer = setTimeout(() => { cleanup(); reject(new Error('連線逾時')); }, 12000);
      function cleanup(){ clearTimeout(timer); delete window[cb]; script.remove(); }
      window[cb] = (data) => { cleanup(); resolve(data); };
      script.onerror = () => { cleanup(); reject(new Error('連線失敗')); };
      const qs = Object.keys(params||{}).map(k => encodeURIComponent(k)+'='+encodeURIComponent(params[k])).join('&');
      script.src = baseUrl + (baseUrl.includes('?') ? '&' : '?') + qs + (qs?'&':'') + 'callback=' + cb;
      document.body.appendChild(script);
    });
  }
  const DEFAULT_PERIODS = [
    ['08:05','08:55'],
    ['09:10','10:00'],
    ['10:10','11:00'],
    ['11:10','12:00'],
    ['13:00','13:50'],
    ['14:00','14:50'],
    ['15:05','15:55'],
    ['16:05','16:55']
  ];

  const GOOGLE_APPS_SCRIPT_CODE = "/**\n * 跨群選修點名｜Google Apps Script\n *\n * 用途：讓 index.html 把點名結果寫進指定的 Google 試算表。\n * Web App 以「執行身分：我」部署後，網頁把試算表網址與點名資料 POST 過來。\n *\n * 試算表格式：\n * A 欄：班級\n * B 欄：座號\n * C 欄：姓名\n * D 欄開始：每次點名一個新欄，第一列寫「日期｜第X節」\n * 例如：D1 = 2026/09/06｜第3節\n *       D2:Dn = 出席／未出席\n */\n\nfunction doGet() {\n  return ContentService\n    .createTextOutput('跨群選修點名 Google 紀錄服務正常運作')\n    .setMimeType(ContentService.MimeType.TEXT);\n}\n\nfunction doPost(e) {\n  const lock = LockService.getScriptLock();\n  lock.waitLock(30000);\n  try {\n    if (!e || !e.postData || !e.postData.contents) {\n      return jsonOutput({ok:false, message:'沒有收到資料'});\n    }\n\n    const data = JSON.parse(e.postData.contents);\n    const sheetUrl = String(data.sheetUrl || '').trim();\n    if (!sheetUrl) return jsonOutput({ok:false, message:'缺少 Google 試算表網址'});\n\n    const ss = SpreadsheetApp.openByUrl(sheetUrl);\n    const sheet = ss.getSheets()[0];\n\n    const rows = Array.isArray(data.rows) ? data.rows : [];\n    if (!rows.length) return jsonOutput({ok:false, message:'沒有學生資料'});\n\n    const date = String(data.date || '').trim() || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');\n    const period = String(data.period || '').trim() || '非上課時段';\n    const periodLabel = period === '非上課時段' ? period : '第' + period + '節';\n    const header = date + '｜' + periodLabel;\n\n    // 確保前三欄標題存在。\n    if (sheet.getLastRow() === 0) {\n      sheet.getRange(1, 1, 1, 3).setValues([['班級', '座號', '姓名']]);\n    } else {\n      const first = sheet.getRange(1, 1, 1, 3).getValues()[0];\n      if (!first[0] && !first[1] && !first[2]) {\n        sheet.getRange(1, 1, 1, 3).setValues([['班級', '座號', '姓名']]);\n      }\n    }\n\n    // 每次點名都新增一欄，不覆蓋既有紀錄。\n    const newCol = Math.max(sheet.getLastColumn() + 1, 4);\n    sheet.getRange(1, newCol).setValue(header);\n\n    // 讀取既有學生資料，依「班級 + 座號」找學生所在列。\n    const lastRow = Math.max(sheet.getLastRow(), 2);\n    const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];\n    const rowMap = new Map();\n    existing.forEach((r, i) => {\n      const key = makeKey(r[0], r[1]);\n      if (key) rowMap.set(key, i + 2);\n    });\n\n    // 若新名單出現原本沒有的學生，就加到最下面。\n    const newStudents = [];\n    rows.forEach(item => {\n      const key = makeKey(item.className, item.seat);\n      if (key && !rowMap.has(key)) {\n        newStudents.push([\n          String(item.className || ''),\n          String(item.seat || ''),\n          String(item.name || '')\n        ]);\n      }\n    });\n\n    if (newStudents.length) {\n      const startRow = Math.max(sheet.getLastRow() + 1, 2);\n      sheet.getRange(startRow, 1, newStudents.length, 3).setValues(newStudents);\n      newStudents.forEach((r, i) => {\n        const key = makeKey(r[0], r[1]);\n        rowMap.set(key, startRow + i);\n      });\n    }\n\n    // 寫入本次點名狀態。\n    rows.forEach(item => {\n      const key = makeKey(item.className, item.seat);\n      if (!key) return;\n      const row = rowMap.get(key);\n      if (!row) return;\n      sheet.getRange(row, newCol).setValue(String(item.status || ''));\n    });\n\n    // 簡單格式化，讓期末查看比較清楚。\n    sheet.getRange(1, 1, 1, newCol).setFontWeight('bold');\n    sheet.getRange(1, newCol).setWrap(true);\n    sheet.setFrozenRows(1);\n    sheet.setFrozenColumns(3);\n    sheet.autoResizeColumns(1, 3);\n\n    return jsonOutput({ok:true, message:'已新增一欄點名紀錄', column:newCol, header:header});\n  } catch (err) {\n    return jsonOutput({ok:false, message:String(err && err.message ? err.message : err)});\n  } finally {\n    lock.releaseLock();\n  }\n}\n\nfunction makeKey(cls, seat) {\n  const c = String(cls || '').trim();\n  const s = String(seat || '').trim();\n  return c && s ? c + '|' + s : '';\n}\n\nfunction jsonOutput(obj) {\n  return ContentService\n    .createTextOutput(JSON.stringify(obj))\n    .setMimeType(ContentService.MimeType.JSON);\n}\n";

  const DRIVE_APPS_SCRIPT_CODE = "/**\n * 課程名單雲端硬碟讀取服務｜Google Apps Script\n *\n * 用途：讓點名 App 讀取你指定的 Google 雲端硬碟資料夾裡的課程名單檔案，\n * 在「新增課程」畫面用下拉選單直接選課程，不用每次手動上傳檔案。\n *\n * 這是一支「獨立」的 Apps Script 專案（跟「Google 試算表同步」用的那一支\n * 是分開的兩個部署，各自有各自的網址）。部署方式：\n *\n * 1. 用你的私人 Google 帳號到 script.google.com，點「新增專案」\n *    （這次不需要從某一份試算表裡面開，直接開一個獨立專案就可以）。\n * 2. 把這個檔案的內容整段貼進去，取代原本的內容。\n * 3. 如果你的資料夾裡「只有 Google 試算表、沒有 .xlsx 檔案」，可以跳過這一步；\n *    如果資料夾裡「有 .xlsx 檔案」，需要多加一個服務才能讀取：\n *    左側選單「服務」旁的 ＋ 號 → 選「Drive API」→ 版本選「v3」→ 按「加入」。\n * 4. 按「部署」→「新增部署作業」→ 類型選「網頁應用程式」→\n *    「具有存取權的使用者」選「所有人」→ 部署 → 依需要完成帳號授權。\n * 5. 複製部署後的網址，貼到 App 設定頁「課程名單雲端硬碟」的\n *    「雲端硬碟讀取服務網址」欄位。\n *\n * 支援的檔案格式：\n * - Google 試算表（推薦，讀取最穩定、不需要額外設定）\n * - .xlsx 檔案（會先自動轉成暫存的 Google 試算表讀取，讀完立刻刪除暫存檔；\n *   需要完成上面第 3 步的 Drive API 進階服務設定）\n *\n * 名單檔案內容格式：跟本機上傳完全相同——第一列是標題列，\n * 需包含「班級」「座號」欄位，姓名欄位可選填。\n */\n\nconst XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';\n\nfunction doGet(e) {\n  const p = (e && e.parameter) ? e.parameter : {};\n  const callback = String(p.callback || '').replace(/[^a-zA-Z0-9_$.]/g, '');\n\n  let result;\n  try {\n    if (p.action === 'listFolder') {\n      result = listFolder(String(p.folderUrl || p.folderId || '').trim());\n    } else if (p.action === 'readFile') {\n      result = readFile(String(p.fileId || '').trim(), String(p.mimeType || '').trim());\n    } else {\n      result = { ok: true, message: '課程名單雲端硬碟讀取服務正常運作' };\n    }\n  } catch (err) {\n    result = { ok: false, message: String(err && err.message ? err.message : err) };\n  }\n\n  const text = JSON.stringify(result);\n  if (callback) {\n    return ContentService\n      .createTextOutput(callback + '(' + text + ');')\n      .setMimeType(ContentService.MimeType.JAVASCRIPT);\n  }\n  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);\n}\n\n// 從網址或直接貼的 ID 裡取出雲端硬碟資料夾的 ID。\nfunction extractDriveId(input) {\n  const s = String(input || '').trim();\n  const m = s.match(/[-\\w]{20,}/);\n  return m ? m[0] : s;\n}\n\n// 列出資料夾裡的課程名單檔案（Google 試算表 + .xlsx）。\nfunction listFolder(folderUrlOrId) {\n  const folderId = extractDriveId(folderUrlOrId);\n  if (!folderId) return { ok: false, message: '缺少雲端硬碟資料夾網址' };\n\n  const folder = DriveApp.getFolderById(folderId);\n  const files = [];\n\n  const sheetIter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);\n  while (sheetIter.hasNext()) {\n    const f = sheetIter.next();\n    files.push({ id: f.getId(), name: f.getName(), mimeType: MimeType.GOOGLE_SHEETS });\n  }\n\n  const xlsxIter = folder.getFilesByType(XLSX_MIME_TYPE);\n  while (xlsxIter.hasNext()) {\n    const f = xlsxIter.next();\n    files.push({ id: f.getId(), name: f.getName(), mimeType: XLSX_MIME_TYPE });\n  }\n\n  files.sort(function (a, b) { return a.name.localeCompare(b.name, 'zh-Hant'); });\n  return { ok: true, folderName: folder.getName(), files: files };\n}\n\n// 讀取單一檔案的內容，回傳跟本機上傳解析出來一樣的「整張表格」資料（二維陣列）。\nfunction readFile(fileId, mimeType) {\n  if (!fileId) return { ok: false, message: '缺少檔案 ID' };\n\n  let ss;\n  let tempFileId = null;\n  try {\n    if (mimeType === XLSX_MIME_TYPE) {\n      // .xlsx 檔案：先轉成暫存的 Google 試算表才能讀取內容，用完即刪。\n      const blob = DriveApp.getFileById(fileId).getBlob();\n      const converted = Drive.Files.create(\n        { name: '__attendance_temp_convert__' + fileId, mimeType: MimeType.GOOGLE_SHEETS },\n        blob\n      );\n      tempFileId = converted.id;\n      ss = SpreadsheetApp.openById(converted.id);\n    } else {\n      ss = SpreadsheetApp.openById(fileId);\n    }\n\n    const sheet = ss.getSheets()[0];\n    const values = sheet.getDataRange().getValues();\n    return { ok: true, rows: values };\n  } catch (err) {\n    if (String(err).indexOf('Drive is not defined') !== -1) {\n      return { ok: false, message: '讀取 .xlsx 檔案需要先啟用「Drive API」進階服務，請參考程式最上面的說明完成設定。' };\n    }\n    throw err;\n  } finally {\n    if (tempFileId) {\n      try { Drive.Files.delete(tempFileId); } catch (e2) { /* 刪不掉暫存檔也沒關係，不影響結果 */ }\n    }\n  }\n}\n";

  async function getSetting(key, fallback=''){ return (await storageGet(key)) ?? fallback; }

  function buildPeriodSettings(){
    const box = $('periodSettings');
    box.innerHTML = '';
    for(let i=0;i<8;i++){
      const row=document.createElement('div'); row.className='period-row';
      row.innerHTML=`<div class="pname">第 ${i+1} 節</div>
        <input type="time" id="periodStart${i}" aria-label="第 ${i+1} 節開始時間">
        <div>～</div>
        <input type="time" id="periodEnd${i}" aria-label="第 ${i+1} 節結束時間">`;
      box.appendChild(row);
    }
  }

  function buildGoogleGuideCode(){
    const box=$('appsScriptCode');
    if(box) box.textContent=GOOGLE_APPS_SCRIPT_CODE;
  }

  function buildDriveGuideCode(){
    const box=$('driveAppsScriptCode');
    if(box) box.textContent=DRIVE_APPS_SCRIPT_CODE;
  }

  async function loadSettings(){
    const savedPeriods = await getSetting(PERIODS_KEY, null);
    const periods = savedPeriods ? JSON.parse(savedPeriods) : DEFAULT_PERIODS;
    periods.forEach((p,i)=>{
      if($('periodStart'+i)) $('periodStart'+i).value=p[0]||'';
      if($('periodEnd'+i)) $('periodEnd'+i).value=p[1]||'';
    });
    $('googleSheetUrlInput').value = await getSetting(GOOGLE_SHEET_KEY,'');
    $('googleLogUrlInput').value = await getSetting(GOOGLE_LOG_KEY,'');
    $('driveFolderUrlInput').value = await getSetting(DRIVE_FOLDER_KEY, DEFAULT_DRIVE_FOLDER_URL);
    $('driveLogUrlInput').value = await getSetting(DRIVE_LOG_KEY, DEFAULT_DRIVE_LOG_URL);
    $('messagePrefixInput').value = await getSetting(MESSAGE_PREFIX_KEY, DEFAULT_MESSAGE_PREFIX);
    updateGoogleNote();
  }

  async function saveSettings(){
    const periods=[];
    for(let i=0;i<8;i++){
      periods.push([$('periodStart'+i).value,$('periodEnd'+i).value]);
    }
    await storageSet(PERIODS_KEY, JSON.stringify(periods));
    await storageSet(GOOGLE_SHEET_KEY, $('googleSheetUrlInput').value.trim());
    await storageSet(GOOGLE_LOG_KEY, $('googleLogUrlInput').value.trim());
    await storageSet(DRIVE_FOLDER_KEY, $('driveFolderUrlInput').value.trim());
    await storageSet(DRIVE_LOG_KEY, $('driveLogUrlInput').value.trim());
    await storageSet(MESSAGE_PREFIX_KEY, $('messagePrefixInput').value);
    updateCurrentPeriod();
    updateGoogleNote();
    toast('設定已儲存');
  }

  async function getPeriods(){
    const saved = await getSetting(PERIODS_KEY, null);
    return saved ? JSON.parse(saved) : DEFAULT_PERIODS;
  }

  function hmToMinutes(hm){
    if(!/^\d{2}:\d{2}$/.test(hm||'')) return null;
    const [h,m]=hm.split(':').map(Number); return h*60+m;
  }

  function taipeiMinutesNow(){
    const parts=new Intl.DateTimeFormat('zh-TW',{timeZone:'Asia/Taipei',hour:'2-digit',minute:'2-digit',hour12:false}).formatToParts(new Date());
    const h=Number(parts.find(x=>x.type==='hour')?.value||0);
    const m=Number(parts.find(x=>x.type==='minute')?.value||0);
    return h*60+m;
  }

  async function getCurrentPeriod(){
    const periods=await getPeriods();
    const minutes=taipeiMinutesNow();
    for(let i=0;i<periods.length;i++){
      const start=hmToMinutes(periods[i][0]), end=hmToMinutes(periods[i][1]);
      if(start!==null && end!==null && start<=minutes && minutes<=end) return i+1;
    }
    return null;
  }

  function formatTaipeiTime(){
    return new Intl.DateTimeFormat('zh-TW',{
      timeZone:'Asia/Taipei',hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false
    }).format(new Date());
  }

  async function updateLiveClock(){
    const clock = $('liveClock');
    const periodEl = $('livePeriod');
    if(!clock || !periodEl) return;
    clock.textContent = formatTaipeiTime();
    const p = await getCurrentPeriod();
    periodEl.textContent = p ? `（自動偵測：第${p}節，僅供參考）` : '（自動偵測：非上課時段，僅供參考）';
  }

  async function updateCurrentPeriod(){
    await checkDayRollover();
    await updateLiveClock();
    if(currentRoster) await refreshSessionUI();
  }

  $('settingsBtn').addEventListener('click', async ()=>{
    buildPeriodSettings();
    buildGoogleGuideCode();
    buildDriveGuideCode();
    await loadSettings();
    showScreen('settings');
  });
  $('backFromSettingsBtn').addEventListener('click', ()=>{ updateCurrentPeriod(); showScreen(currentRoster ? 'attend' : 'home'); });
  $('saveSettingsBtn').addEventListener('click', saveSettings);
  $('openAdvancedSyncBtn').addEventListener('click', async ()=>{
    buildGoogleGuideCode();
    buildDriveGuideCode();
    await loadSettings();
    showScreen('settings-sync');
  });
  $('backFromAdvancedSyncBtn').addEventListener('click', ()=>{ showScreen('settings'); });
  $('saveAdvancedSyncBtn').addEventListener('click', saveSettings);

  function taipeiDateStr(){
    return new Intl.DateTimeFormat('zh-TW',{timeZone:'Asia/Taipei',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date()).replace(/\//g,'/');
  }

  async function buildGooglePayload(periodOverride){
    const sheetUrl = (await getSetting(GOOGLE_SHEET_KEY,'')).trim();
    const period = periodOverride !== undefined ? periodOverride : await getCurrentPeriod();
    const now = new Date();
    const dateStr = taipeiDateStr();
    const rows=[];
    Object.keys(currentRoster.classes).forEach(cls=>currentRoster.classes[cls].forEach(stu=>rows.push({
      className:cls, seat:stu.seat, name:stu.name||'', status:stu.absent?'未出席':'出席'
    })));
    return {sheetUrl, course:currentRoster.name, date:dateStr, period:period?String(period):'非上課時段', timestamp:now.toISOString(), rows};
  }

  async function recordToGoogle(periodOverride){
    const sheetUrl = (await getSetting(GOOGLE_SHEET_KEY,'')).trim();
    const webAppUrl = (await getSetting(GOOGLE_LOG_KEY,'')).trim();
    if(!sheetUrl || !webAppUrl) return {configured:false};
    try{
      const payload=await buildGooglePayload(periodOverride);
      await fetch(webAppUrl,{method:'POST',mode:'no-cors',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload)});
      return {configured:true, sent:true};
    }catch(e){
      console.error(e);
      return {configured:true, sent:false};
    }
  }

  async function updateGoogleNote(){
    const note=$('autoGoogleNote');
    if(!note) return;
    const sheetUrl=(await getSetting(GOOGLE_SHEET_KEY,'')).trim();
    const webAppUrl=(await getSetting(GOOGLE_LOG_KEY,'')).trim();
    const ready=!!(sheetUrl && webAppUrl);
    note.textContent=ready ? '✓ 已設定 Google：每一節點名送出時都會自動記錄。' : 'Google 試算表未設定；仍可正常點名與產生訊息。';
    note.classList.toggle('ready',ready);
    note.classList.toggle('off',!ready);
  }

  function testGoogleConnection(){
    return new Promise(async (resolve)=>{
      const sheetUrl=(await getSetting(GOOGLE_SHEET_KEY,'')).trim() || $('googleSheetUrlInput').value.trim();
      const webAppUrl=(await getSetting(GOOGLE_LOG_KEY,'')).trim() || $('googleLogUrlInput').value.trim();
      const result=$('googleTestResult');
      if(!sheetUrl || !webAppUrl){
        result.className='google-test-result err';
        result.textContent='❌ 請先填好「① Google 試算表網址」及「② Google 寫入服務網址」。';
        return resolve(false);
      }
      result.className='google-test-result';
      result.textContent='測試中…';
      const cb='googleTestCallback_'+Date.now();
      const script=document.createElement('script');
      const timer=setTimeout(()=>{
        cleanup(); result.className='google-test-result err';
        result.textContent='❌ 測試逾時。請確認 Apps Script 已部署為「網頁應用程式」，且執行身分為「我」。';
        resolve(false);
      },10000);
      function cleanup(){ clearTimeout(timer); delete window[cb]; script.remove(); }
      window[cb]=(data)=>{
        cleanup();
        if(data && data.ok){
          result.className='google-test-result ok';
          result.textContent=`✅ 連線成功！已找到試算表：「${data.sheetName || ''}」。現在點名送出時就會自動記錄。`;
          resolve(true);
        }else{
          result.className='google-test-result err';
          result.textContent='❌ 連線失敗：'+(data && data.message ? data.message : '無法確認 Google 權限。');
          resolve(false);
        }
      };
      script.onerror=()=>{ cleanup(); result.className='google-test-result err'; result.textContent='❌ 無法連到 Google 寫入服務。請檢查網址與部署權限。'; resolve(false); };
      script.src=webAppUrl+(webAppUrl.includes('?')?'&':'?')+'action=test&sheetUrl='+encodeURIComponent(sheetUrl)+'&callback='+cb;
      document.body.appendChild(script);
    });
  }

  $('testGoogleBtn').addEventListener('click', async ()=>{
    await storageSet(GOOGLE_SHEET_KEY, $('googleSheetUrlInput').value.trim());
    await storageSet(GOOGLE_LOG_KEY, $('googleLogUrlInput').value.trim());
    await testGoogleConnection();
    updateGoogleNote();
  });

  $('copyAppsScriptBtn').addEventListener('click', async () => {
    const code=$('appsScriptCode').textContent;
    try{
      await navigator.clipboard.writeText(code);
      $('screen-settings').querySelector('.setup-details').classList.add('show-code');
      toast('已複製 Google 寫入程式');
    }catch(e){
      toast('複製失敗，請直接選取下方程式');
      $('screen-settings').querySelector('.setup-details').classList.add('show-code');
    }
  });

  async function testDriveConnection(){
    const folderUrl=(await getSetting(DRIVE_FOLDER_KEY, DEFAULT_DRIVE_FOLDER_URL)).trim() || $('driveFolderUrlInput').value.trim();
    const webAppUrl=(await getSetting(DRIVE_LOG_KEY, DEFAULT_DRIVE_LOG_URL)).trim() || $('driveLogUrlInput').value.trim();
    const result=$('driveTestResult');
    if(!folderUrl || !webAppUrl){
      result.className='google-test-result err';
      result.textContent='❌ 請先填好「① 雲端硬碟資料夾網址」及「② 雲端硬碟讀取服務網址」。';
      return false;
    }
    result.className='google-test-result';
    result.textContent='測試中…';
    try{
      const data = await jsonpRequest(webAppUrl, {action:'listFolder', folderUrl});
      if(data && data.ok){
        const count = (data.files||[]).length;
        result.className='google-test-result ok';
        result.textContent=`✅ 連線成功！資料夾「${data.folderName||''}」裡找到 ${count} 個課程名單檔案。現在「新增課程」畫面就能直接選取。`;
        return true;
      }else{
        result.className='google-test-result err';
        result.textContent='❌ 連線失敗：'+(data && data.message ? data.message : '無法讀取資料夾。');
        return false;
      }
    }catch(e){
      result.className='google-test-result err';
      result.textContent='❌ 無法連到雲端硬碟讀取服務。請確認 Apps Script 已部署為「網頁應用程式」，且存取權為「所有人」。';
      return false;
    }
  }

  $('testDriveBtn').addEventListener('click', async ()=>{
    await storageSet(DRIVE_FOLDER_KEY, $('driveFolderUrlInput').value.trim());
    await storageSet(DRIVE_LOG_KEY, $('driveLogUrlInput').value.trim());
    await testDriveConnection();
  });

  $('copyDriveScriptBtn').addEventListener('click', async () => {
    const code=$('driveAppsScriptCode').textContent;
    try{
      await navigator.clipboard.writeText(code);
      $('driveSetupDetails').classList.add('show-code');
      toast('已複製雲端硬碟讀取程式');
    }catch(e){
      toast('複製失敗，請直接選取下方程式');
      $('driveSetupDetails').classList.add('show-code');
    }
  });

  // ---------- two-period roll call / message generation ----------
  // 每門課都是雙節課：第一節點完名先記錄下來，等第二節上課時間開始後
  // 再點一次名，兩節課的紀錄才會合併成一則訊息送出。
  function periodLabel(p){ return p ? `第${p}節` : '非上課時段'; }
  function statusLabel(seats){ return (seats && seats.length) ? `缺席${seats.join('、')}號` : '全員到齊'; }

  function snapshotAbsences(){
    const records = {};
    Object.keys(currentRoster.classes).forEach(cls => {
      records[cls] = currentRoster.classes[cls].filter(s=>s.absent).map(s=>s.seat);
    });
    return records;
  }

  function resetAttendanceToPresent(){
    Object.values(currentRoster.classes).forEach(list => list.forEach(s => { s.absent = false; }));
  }

  // 依「班級 → 缺席座號陣列」的紀錄，還原目前畫面上每位學生的出席狀態（供返回修改時使用）。
  function restoreRecords(records){
    Object.keys(currentRoster.classes).forEach(cls => {
      const absentSeats = new Set((records && records[cls]) || []);
      currentRoster.classes[cls].forEach(s => { s.absent = absentSeats.has(s.seat); });
    });
  }

  // ---------- 節次草稿：讓老師可以自由切換節次調整器，每個節次各自保留當時輸入的紀錄 ----------
  // 存放在 currentRoster.periodDrafts = { date, records: { '<節次或none>': {班級: [缺席座號]} } }
  function draftKey(period){ return period === null || period === undefined ? 'none' : String(period); }

  function ensureDraftsToday(){
    const today = taipeiDateStr();
    if(!currentRoster.periodDrafts || currentRoster.periodDrafts.date !== today){
      currentRoster.periodDrafts = { date: today, records: {} };
    }
    return currentRoster.periodDrafts;
  }

  function saveDraftForPeriod(period){
    const drafts = ensureDraftsToday();
    drafts.records[draftKey(period)] = snapshotAbsences();
  }

  // 找出某個節次「當初輸入過的紀錄」：優先看草稿，再看目前記錄中、已完成送出的第一節/第二節紀錄。
  // 都找不到就回傳 null，代表這個節次還沒有人點過名，畫面上應該顯示全員出席。
  function getRecordsForPeriod(period){
    const today = taipeiDateStr();
    const drafts = currentRoster.periodDrafts;
    if(drafts && drafts.date === today && drafts.records[draftKey(period)]){
      return drafts.records[draftKey(period)];
    }
    const pending = currentRoster.pendingSession;
    if(pending && pending.date === today && pending.period === period) return pending.records;
    if(lastCompletedSessions && lastCompletedSessions.date === today){
      if(lastCompletedSessions.first.period === period) return lastCompletedSessions.first.records;
      if(lastCompletedSessions.second.period === period) return lastCompletedSessions.second.records;
    }
    return null;
  }

  // 把畫面切換去顯示某個節次：有紀錄就帶入，沒有就顯示全員出席（代表這節還沒點過名）。
  function loadPeriodForDisplay(period){
    const records = getRecordsForPeriod(period);
    if(records) restoreRecords(records);
    else resetAttendanceToPresent();
  }

  // 偵測「跨到新的一天」：只要日期變了（例如老師把瀏覽器分頁一直開著沒關，隔了一天甚至隔週才又點開），
  // 就把所有跟「今天點名」有關的暫存資料全部清空，並強制把畫面重設為全新、乾淨、預設「出席」的狀態，
  // 確保絕對不會把之前點名的紀錄誤帶到新的一次點名裡。
  async function checkDayRollover(){
    const today = taipeiDateStr();
    if(lastSeenDate === null){ lastSeenDate = today; return false; }
    if(lastSeenDate === today) return false;
    lastSeenDate = today;
    lastCompletedSessions = null;
    pendingEditSecond = null;
    manualPeriod = null;
    if(currentRoster){
      currentRoster.pendingSession = null;
      currentRoster.periodDrafts = null;
      resetAttendanceToPresent();
      if(currentCourseId) await setRoster(currentCourseId, currentRoster);
      if($('screen-attend').classList.contains('active')) renderAttendance();
    }
    return true;
  }

  async function buildMessage(first, second){
    const prefix = await getSetting(MESSAGE_PREFIX_KEY, DEFAULT_MESSAGE_PREFIX);
    const classes = Object.keys(currentRoster.classes).sort((a,b) => String(a).localeCompare(String(b),'zh-Hant',{numeric:true}));
    const lines = [];
    classes.forEach(cls => {
      lines.push(cls);
      lines.push(`　${periodLabel(first.period)} ${statusLabel(first.records[cls])}`);
      lines.push(`　${periodLabel(second.period)} ${statusLabel(second.records[cls])}`);
    });
    const body = `${currentRoster.name}\n${lines.join('\n')}`;
    return prefix ? `${prefix}\n${body}` : body;
  }

  // 節次調整器：預設跟隨自動偵測的節次，但老師可以手動往前/往後調整
  // （例如忘記點名、事後補登，或想連續完成兩次點名時使用）。
  function currentPickerValue(auto){
    return manualPeriod !== null ? manualPeriod : (auto || 0);
  }

  async function refreshPeriodPicker(){
    const valueEl = $('periodPickerValue');
    if(!valueEl) return;
    const auto = await getCurrentPeriod();
    const val = currentPickerValue(auto);
    valueEl.textContent = periodLabel(val || null);
  }

  // 切換節次調整器時：先把目前畫面上（正在輸入、還沒送出）的狀態存成這個節次的草稿，
  // 再把新節次「當初輸入過的紀錄」帶進畫面——這樣老師可以自由來回切換節次核對／修改，
  // 不會因為換了節次，畫面內容跟原本輸入的對不起來。
  async function switchPickerPeriod(newManualPeriod){
    if(!currentRoster) { manualPeriod = newManualPeriod; refreshPeriodPicker(); return; }
    await checkDayRollover();
    const auto = await getCurrentPeriod();
    const leavingPeriod = currentPickerValue(auto) || null;
    saveDraftForPeriod(leavingPeriod);
    manualPeriod = newManualPeriod;
    const enteringPeriod = newManualPeriod || null;
    loadPeriodForDisplay(enteringPeriod);
    await setRoster(currentCourseId, currentRoster);
    renderAttendance();
  }

  $('periodPrevBtn').addEventListener('click', async () => {
    const auto = await getCurrentPeriod();
    await switchPickerPeriod(Math.max(0, currentPickerValue(auto) - 1));
  });
  $('periodNextBtn').addEventListener('click', async () => {
    const auto = await getCurrentPeriod();
    await switchPickerPeriod(Math.min(8, currentPickerValue(auto) + 1));
  });

  // 判斷目前點名進度，更新畫面上的提示文字（僅供提醒，不會鎖住按鈕）。
  async function refreshSessionUI(){
    const bar = $('sessionStatusBar');
    const btn = $('genMsgBtn');
    if(!bar || !btn || !currentRoster) return;
    await refreshPeriodPicker();
    const today = taipeiDateStr();
    const auto = await getCurrentPeriod();
    const pending = currentRoster.pendingSession;

    if(!pending || pending.date !== today){
      bar.hidden = true;
      btn.textContent = '完成本節點名';
      btn.disabled = false;
      return;
    }

    bar.hidden = false;
    const stillSamePeriod = auto !== null && auto === pending.period;
    if(stillSamePeriod){
      bar.className = 'session-status waiting';
      bar.textContent = `✅ 已完成${periodLabel(pending.period)}的點名。如果現在還是同一節，可以先等一下；如果是要補點其他節，可以在下方調整節次後直接送出。`;
    }else{
      bar.className = 'session-status ready';
      bar.textContent = `🔔 已可以進行第二次點名，請確認下方節次無誤，完成點名後按下方按鈕送出這兩節課的紀錄。`;
    }
    btn.textContent = '完成第二次點名並送出紀錄';
    btn.disabled = false;
  }

  async function handleGenMsgClick(){
    await checkDayRollover();
    const btn = $('genMsgBtn');
    const today = taipeiDateStr();
    const auto = await getCurrentPeriod();
    const period = manualPeriod !== null ? (manualPeriod || null) : auto;
    const pending = currentRoster.pendingSession;
    const alreadyStartedToday = pending && pending.date === today;

    // 兩次都選了同一節——多半是誤按，先跟老師確認一次，但不強制擋下來。
    if(alreadyStartedToday && period === pending.period){
      const ok = window.confirm(`兩次點名都是${periodLabel(period)}，確定要合併送出嗎？\n\n如果其實是要等下一節再點，可以按「取消」。`);
      if(!ok) return;
    }

    btn.disabled = true;
    btn.textContent = '處理中…';
    try{
      if(!alreadyStartedToday){
        // ---- 第一節：先記錄下來，重置畫面準備點第二次 ----
        const records = snapshotAbsences();
        const googleResult = await recordToGoogle(period);
        currentRoster.pendingSession = { date: today, period, records };
        saveDraftForPeriod(period);
        if(pendingEditSecond){
          // ---- 從「返回修改點名」進來：改完第一節後，自動帶入原本第二節的紀錄繼續修改 ----
          manualPeriod = pendingEditSecond.period;
          restoreRecords(pendingEditSecond.records);
          pendingEditSecond = null;
        }else{
          // 第一次點名送出後，直接把節次調整器往前推一節（接下來幾乎都是點下一節的名），
          // 如果這個節次之前已經輸入過紀錄（例如老師剛剛自己切換節次核對過），就帶入那份紀錄。
          const nextPeriodValue = Math.min(8, (period || 0) + 1);
          manualPeriod = nextPeriodValue;
          loadPeriodForDisplay(nextPeriodValue || null);
        }
        await setRoster(currentCourseId, currentRoster);
        renderAttendance();
        if(googleResult.configured && googleResult.sent) toast('已記錄第一次點名，Google 紀錄已送出，可繼續點第二次');
        else if(googleResult.configured && !googleResult.sent) toast('已記錄第一次點名，但 Google 紀錄送出失敗');
        else toast('已記錄第一次點名，請於下一節再點一次名');
      }else{
        // ---- 第二節：合併兩節紀錄，產生並送出訊息 ----
        const secondRecords = snapshotAbsences();
        const googleResult = await recordToGoogle(period);
        saveDraftForPeriod(period);
        const msg = await buildMessage(pending, { period, records: secondRecords });
        lastCompletedSessions = {
          date: today,
          first: { period: pending.period, records: pending.records },
          second: { period, records: secondRecords }
        };
        currentRoster.pendingSession = null;
        manualPeriod = null;
        await setRoster(currentCourseId, currentRoster);
        $('messageBox').textContent = msg;
        $('lineShareLink').href = 'https://line.me/R/msg/text/?' + encodeURIComponent(msg);
        showScreen('result');
        if(googleResult.configured && googleResult.sent) toast('訊息已產生，Google 點名紀錄已送出');
        else if(googleResult.configured && !googleResult.sent) toast('訊息已產生，但 Google 紀錄送出失敗');
        else toast('訊息已產生');
      }
    } finally {
      btn.disabled = false;
      await refreshSessionUI();
    }
  }

  $('genMsgBtn').addEventListener('click', handleGenMsgClick);

  $('backToAttendBtn').addEventListener('click', async () => {
    await checkDayRollover();
    if(lastCompletedSessions && currentRoster){
      // 回到「第一個節次」重新開始修改，並暫存第二節的原始紀錄，等第一節改完後自動帶入。
      restoreRecords(lastCompletedSessions.first.records);
      currentRoster.pendingSession = null;
      manualPeriod = lastCompletedSessions.first.period;
      pendingEditSecond = lastCompletedSessions.second;
      await setRoster(currentCourseId, currentRoster);
      renderAttendance();
    }
    showScreen('attend');
  });

  $('copyBtn').addEventListener('click', async () => {
    const msg = $('messageBox').textContent;
    try{
      await navigator.clipboard.writeText(msg);
      toast('已複製訊息文字');
    }catch(e){
      const ta = document.createElement('textarea');
      ta.value = msg; document.body.appendChild(ta); ta.select();
      document.execCommand('copy'); document.body.removeChild(ta);
      toast('已複製訊息文字');
    }
  });

  function escapeHtml(str){
    return String(str).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  // ---------- init ----------
  buildPeriodSettings();
  buildGoogleGuideCode();
  buildDriveGuideCode();
  loadSettings();
  updateGoogleNote();
  setInterval(updateCurrentPeriod, 30000);
  renderHome();
  showScreen('home');
})();
</script>
</body>
</html>
