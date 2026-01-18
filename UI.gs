// UI.gs - ไฟล์จัดการส่วนติดต่อผู้ใช้
function showFormCreationResult(form, spreadsheet) {
  try {
    var htmlTemplate = HtmlService.createTemplateFromFile('ResultPage');
    
    htmlTemplate.formUrl = form.getPublishedUrl();
    htmlTemplate.editUrl = form.getEditUrl();
    htmlTemplate.spreadsheetUrl = spreadsheet.getUrl();
    htmlTemplate.formTitle = form.getTitle();
    htmlTemplate.createdDate = new Date().toLocaleDateString('th-TH', {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
    
    var htmlOutput = htmlTemplate.evaluate()
      .setWidth(850)
      .setHeight(700)
      .setTitle('✅ สร้างแบบฟอร์มสำเร็จแล้ว');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'สร้างแบบฟอร์มสำเร็จ');
    
  } catch (error) {
    Logger.log('Error showing result: ' + error.toString());
    showError(error);
  }
}

function showError(error) {
  try {
    var errorHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Sarabun', sans-serif;
            padding: 20px;
            background-color: #f8f9fa;
          }
          .error-container {
            max-width: 600px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
          }
          .error-header {
            background-color: #dc3545;
            color: white;
            padding: 20px;
            text-align: center;
          }
          .error-content {
            padding: 30px;
          }
          .error-details {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            border-radius: 5px;
            padding: 15px;
            margin: 20px 0;
            font-family: monospace;
            font-size: 12px;
            overflow: auto;
            max-height: 200px;
          }
          .button {
            background-color: #6c757d;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            margin: 5px;
          }
          .button-primary {
            background-color: #2c7a3e;
          }
        </style>
        <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
      </head>
      <body>
        <div class="error-container">
          <div class="error-header">
            <h2>⚠️ เกิดข้อผิดพลาด</h2>
          </div>
          <div class="error-content">
            <p>ไม่สามารถสร้างแบบฟอร์มได้เนื่องจากเกิดข้อผิดพลาด:</p>
            <div class="error-details">
              ${error.toString()}
            </div>
            <p>กรุณาลองใหม่อีกครั้งหรือติดต่อผู้ดูแลระบบ</p>
            <div style="text-align: center; margin-top: 30px;">
              <button class="button button-primary" onclick="retryCreateForm()">ลองใหม่อีกครั้ง</button>
              <button class="button" onclick="google.script.host.close()">ปิด</button>
            </div>
          </div>
        </div>
        
        <script>
          function retryCreateForm() {
            google.script.run.createAnimalDiseaseForm();
            google.script.host.close();
          }
        </script>
      </body>
      </html>
    `;
    
    var htmlOutput = HtmlService.createHtmlOutput(errorHtml)
      .setWidth(650)
      .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ข้อผิดพลาด');
    
  } catch (uiError) {
    // หากเกิดข้อผิดพลาดในการแสดง UI แสดง alert แทน
    SpreadsheetApp.getUi().alert('เกิดข้อผิดพลาด: ' + error.toString());
  }
}

function showMainMenu() {
  try {
    var htmlTemplate = HtmlService.createTemplateFromFile('MainMenu');
    var htmlOutput = htmlTemplate.evaluate()
      .setWidth(900)
      .setHeight(700)
      .setTitle('ระบบแจ้งโรคระบาดสัตว์');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'เมนูหลัก');
    
  } catch (error) {
    Logger.log('Error showing main menu: ' + error.toString());
    showError(error);
  }
}

function showInstructions() {
  var html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: 'Sarabun', sans-serif;
          padding: 20px;
          line-height: 1.6;
          background-color: #f8f9fa;
        }
        .container {
          max-width: 800px;
          margin: 0 auto;
          background: white;
          border-radius: 10px;
          box-shadow: 0 2px 10px rgba(0,0,0,0.1);
          overflow: hidden;
        }
        .header {
          background-color: #2c7a3e;
          color: white;
          padding: 20px;
          text-align: center;
        }
        .content {
          padding: 30px;
        }
        .step {
          background-color: #f8f9fa;
          padding: 20px;
          margin: 15px 0;
          border-radius: 8px;
          border-left: 4px solid #2c7a3e;
        }
        .step-number {
          background-color: #2c7a3e;
          color: white;
          width: 30px;
          height: 30px;
          border-radius: 50%;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          margin-right: 10px;
          font-weight: bold;
        }
        .button {
          background-color: #2c7a3e;
          color: white;
          padding: 10px 20px;
          border: none;
          border-radius: 5px;
          cursor: pointer;
          margin-top: 20px;
        }
      </style>
      <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h2>📋 คู่มือการใช้งาน</h2>
          <p>ระบบแจ้งโรคระบาดสัตว์ กรมปศุสัตว์</p>
        </div>
        <div class="content">
          <div class="step">
            <div><span class="step-number">1</span> <strong>การสร้างแบบฟอร์ม</strong></div>
            <p>คลิกที่เมนู "🚨 แจ้งโรคระบาดสัตว์" → "📝 สร้างแบบฟอร์มใหม่" เพื่อสร้างแบบฟอร์มสำหรับรับแจ้งโรค</p>
          </div>
          <div class="step">
            <div><span class="step-number">2</span> <strong>การแจกจ่ายแบบฟอร์ม</strong></div>
            <p>คัดลอก URL ที่สร้างขึ้นและส่งให้เกษตรกรหรือผู้ที่ต้องการแจ้งโรคระบาดสัตว์</p>
          </div>
          <div class="step">
            <div><span class="step-number">3</span> <strong>การดูข้อมูล</strong></div>
            <p>ข้อมูลที่กรอกจะถูกบันทึกใน Google Sheet อัตโนมัติ สามารถดูได้จากลิงก์ที่แสดงหลังสร้างแบบฟอร์ม</p>
          </div>
          <div class="step">
            <div><span class="step-number">4</span> <strong>การจัดการ</strong></div>
            <p>ใช้เมนู "📊 ดูรายงาน" เพื่อดูสถิติและการตอบกลับทั้งหมด</p>
          </div>
          <div style="text-align: center;">
            <button class="button" onclick="google.script.host.close()">เข้าใจแล้ว</button>
            <button class="button" onclick="createForm()" style="background-color: #28a745;">สร้างแบบฟอร์มเลย</button>
          </div>
        </div>
      </div>
      <script>
        function createForm() {
          google.script.run.createAnimalDiseaseForm();
          google.script.host.close();
        }
      </script>
    </body>
    </html>
  `;
  
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(700)
    .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'คู่มือการใช้งาน');
}

function viewReports() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var formData = scriptProperties.getProperty('latest_form');
  
  var html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: 'Sarabun', sans-serif;
          padding: 20px;
          background-color: #f8f9fa;
        }
        .container {
          max-width: 800px;
          margin: 0 auto;
          background: white;
          border-radius: 10px;
          box-shadow: 0 2px 10px rgba(0,0,0,0.1);
          overflow: hidden;
        }
        .header {
          background-color: #2c7a3e;
          color: white;
          padding: 20px;
          text-align: center;
        }
        .content {
          padding: 30px;
        }
        .info-card {
          background-color: #f8f9fa;
          padding: 20px;
          margin: 15px 0;
          border-radius: 8px;
          border-left: 4px solid #2c7a3e;
        }
        .url-box {
          background-color: white;
          border: 1px solid #ddd;
          border-radius: 5px;
          padding: 15px;
          margin: 10px 0;
          word-break: break-all;
          font-family: monospace;
          font-size: 14px;
        }
        .button {
          background-color: #2c7a3e;
          color: white;
          padding: 10px 20px;
          border: none;
          border-radius: 5px;
          cursor: pointer;
          margin: 5px;
        }
      </style>
      <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h2>📊 รายงานและสถิติ</h2>
          <p>ระบบแจ้งโรคระบาดสัตว์</p>
        </div>
        <div class="content">
  `;
  
  if (formData) {
    var data = JSON.parse(formData);
    html += `
          <div class="info-card">
            <h3>แบบฟอร์มล่าสุด</h3>
            <p><strong>ชื่อ:</strong> ${data.title}</p>
            <p><strong>สร้างเมื่อ:</strong> ${new Date(data.created).toLocaleDateString('th-TH')}</p>
          </div>
          <div class="info-card">
            <h3>ลิงก์แบบฟอร์ม</h3>
            <div class="url-box">
              <a href="${data.formUrl}" target="_blank">${data.formUrl}</a>
            </div>
          </div>
          <div class="info-card">
            <h3>ข้อมูลการตอบกลับ</h3>
            <div class="url-box">
              <a href="${data.spreadsheetUrl}" target="_blank">${data.spreadsheetUrl}</a>
            </div>
          </div>
          <div style="text-align: center; margin-top: 30px;">
            <button class="button" onclick="openForm()">เปิดแบบฟอร์ม</button>
            <button class="button" onclick="openSpreadsheet()">เปิดข้อมูล</button>
            <button class="button" onclick="createNewForm()">สร้างแบบฟอร์มใหม่</button>
            <button class="button" onclick="google.script.host.close()">ปิด</button>
          </div>
          <script>
            function openForm() {
              window.open('${data.formUrl}', '_blank');
            }
            function openSpreadsheet() {
              window.open('${data.spreadsheetUrl}', '_blank');
            }
            function createNewForm() {
              google.script.run.createAnimalDiseaseForm();
              google.script.host.close();
            }
          </script>
    `;
  } else {
    html += `
          <div class="info-card">
            <h3>ยังไม่มีแบบฟอร์ม</h3>
            <p>ยังไม่มีแบบฟอร์มที่สร้างขึ้น กรุณาสร้างแบบฟอร์มใหม่เพื่อเริ่มใช้งาน</p>
          </div>
          <div style="text-align: center; margin-top: 30px;">
            <button class="button" onclick="createNewForm()">สร้างแบบฟอร์มใหม่</button>
            <button class="button" onclick="google.script.host.close()">ปิด</button>
          </div>
          <script>
            function createNewForm() {
              google.script.run.createAnimalDiseaseForm();
              google.script.host.close();
            }
          </script>
    `;
  }
  
  html += `
        </div>
      </div>
    </body>
    </html>
  `;
  
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(800)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'รายงาน');
}

function showSettings() {
  var html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: 'Sarabun', sans-serif;
          padding: 20px;
          background-color: #f8f9fa;
        }
        .container {
          max-width: 600px;
          margin: 0 auto;
          background: white;
          border-radius: 10px;
          box-shadow: 0 2px 10px rgba(0,0,0,0.1);
          overflow: hidden;
        }
        .header {
          background-color: #2c7a3e;
          color: white;
          padding: 20px;
          text-align: center;
        }
        .content {
          padding: 30px;
        }
        .setting-item {
          margin: 20px 0;
        }
        .button {
          background-color: #2c7a3e;
          color: white;
          padding: 10px 20px;
          border: none;
          border-radius: 5px;
          cursor: pointer;
          margin: 5px;
        }
      </style>
      <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h2>⚙️ การตั้งค่า</h2>
          <p>ระบบแจ้งโรคระบาดสัตว์</p>
        </div>
        <div class="content">
          <div class="setting-item">
            <h3>ข้อมูลระบบ</h3>
            <p><strong>ชื่อระบบ:</strong> ${CONFIG.APP_NAME}</p>
            <p><strong>เวอร์ชัน:</strong> ${CONFIG.VERSION}</p>
            <p><strong>สถานะ:</strong> <span style="color: #28a745;">พร้อมใช้งาน</span></p>
          </div>
          <div style="text-align: center; margin-top: 30px;">
            <button class="button" onclick="clearCache()">ล้างแคช</button>
            <button class="button" onclick="resetSystem()">รีเซ็ตระบบ</button>
            <button class="button" onclick="google.script.host.close()">ปิด</button>
          </div>
        </div>
      </div>
      <script>
        function clearCache() {
          if (confirm('ต้องการล้างแคชใช่หรือไม่?')) {
            google.script.run.clearCache();
            alert('ล้างแคชสำเร็จแล้ว');
          }
        }
        function resetSystem() {
          if (confirm('ต้องการรีเซ็ตระบบทั้งหมดใช่หรือไม่? การดำเนินการนี้ไม่สามารถย้อนกลับได้')) {
            google.script.run.resetSystem();
            alert('รีเซ็ตระบบสำเร็จแล้ว');
            google.script.host.close();
          }
        }
      </script>
    </body>
    </html>
  `;
  
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(650)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'การตั้งค่า');
}

// ฟังก์ชันอรรถประโยชน์
function clearCache() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert('ล้างแคชสำเร็จแล้ว');
}

function resetSystem() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert('รีเซ็ตระบบสำเร็จแล้ว');
}

function getRecentForms() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var formData = scriptProperties.getProperty('latest_form');
  
  if (formData) {
    var data = JSON.parse(formData);
    return [{
      title: data.title,
      createdDate: new Date(data.created).toLocaleDateString('th-TH'),
      url: data.formUrl,
      editUrl: data.editUrl,
      responseCount: 0 // สามารถดึงข้อมูลจริงจาก Form ได้
    }];
  }
  
  return [];
}

function getFormStats() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var formData = scriptProperties.getProperty('latest_form');
  
  if (formData) {
    var data = JSON.parse(formData);
    try {
      var form = FormApp.openById(data.formId);
      var responses = form.getResponses();
      
      return {
        totalResponses: responses.length,
        lastResponse: responses.length > 0 ? responses[responses.length - 1].getTimestamp() : null,
        formUrl: data.formUrl,
        spreadsheetUrl: data.spreadsheetUrl,
        created: data.created
      };
    } catch (e) {
      return { error: 'ไม่พบแบบฟอร์ม' };
    }
  }
  
  return { error: 'ยังไม่มีแบบฟอร์ม' };
}