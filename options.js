document.addEventListener('DOMContentLoaded', function() {
  // 加载保存的设置
  loadSavedSettings();
  
  // 绑定保存按钮事件
  document.getElementById('save-btn').addEventListener('click', saveSettings);
  
  // 绑定免费申请按钮事件
  document.getElementById('apply-free-btn').addEventListener('click', function() {
    // 在新标签页中打开SiliconFlow网站
    chrome.tabs.create({ url: 'https://cloud.siliconflow.cn/i/imSYdTsF' });
  });
});

// 加载保存的设置
function loadSavedSettings() {
  chrome.storage.local.get(['apiKey', 'userData'], function(data) {
    if (data.apiKey) {
      document.getElementById('api-key').value = data.apiKey;
    }
    
    if (data.userData) {
      document.getElementById('user-name').value = data.userData.name || '';
      document.getElementById('user-contact').value = data.userData.contact || '';
    }
  });
}

// 保存设置
function saveSettings() {
  const apiKey = document.getElementById('api-key').value.trim();
  const userName = document.getElementById('user-name').value.trim();
  const userContact = document.getElementById('user-contact').value.trim();
  
  // 保存API密钥
  chrome.storage.local.set({ 
    apiKey: apiKey,
    userData: {
      name: userName,
      contact: userContact
    }
  }, function() {
    showStatusMessage('设置已保存', 'success');
  });
}

// 显示状态消息
function showStatusMessage(message, type) {
  const statusElement = document.getElementById('status');
  statusElement.textContent = message;
  statusElement.className = 'status-message ' + type;
  statusElement.style.display = 'block';
  
  // 3秒后自动隐藏消息
  setTimeout(function() {
    statusElement.style.display = 'none';
  }, 3000);
}