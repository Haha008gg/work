document.addEventListener('DOMContentLoaded', function() {
  // 添加自动保存功能
  const jobTitleInput = document.getElementById('job-title');
  const jobDescriptionTextarea = document.getElementById('job-description');
  
  jobTitleInput.addEventListener('input', function() {
    chrome.storage.local.set({ 'savedJobTitle': this.value });
  });
  
  jobDescriptionTextarea.addEventListener('input', function() {
    chrome.storage.local.set({ 'savedJobDescription': this.value });
  });
  
  // 加载之前保存的招聘信息和简历内容
  loadSavedData();
  
  // 绑定按钮点击事件
  document.getElementById('generate-btn').addEventListener('click', generateResume);
  document.getElementById('download-btn').addEventListener('click', downloadResume);
  document.getElementById('copy-btn').addEventListener('click', copyResumeContent);
  
  // 添加设置API Key的功能
  addSettingsSection();
});

// 加载保存的数据函数，用于初始化和页面可见性变化时调用
function loadSavedData() {
  chrome.storage.local.get(['savedJobTitle', 'savedJobDescription', 'userData', 'savedResumePreview', 'savedResumeContent'], function(data) {
    // 恢复岗位信息
    if (data.savedJobTitle) {
      document.getElementById('job-title').value = data.savedJobTitle;
    }
    
    if (data.savedJobDescription) {
      document.getElementById('job-description').value = data.savedJobDescription;
    }
    
    // 同时加载用户数据
    if (data.userData) {
      document.getElementById('user-name').value = data.userData.name || '';
      document.getElementById('user-contact').value = data.userData.contact || '';
    }
    
    // 恢复之前生成的预览内容
    if (data.savedResumePreview && data.savedResumeContent) {
      // 移除```markdown字符
      let formattedPreview = data.savedResumePreview;
      formattedPreview = formattedPreview.replace(/```markdown/g, '');
      formattedPreview = formattedPreview.replace(/```/g, '');
      document.getElementById('resume-preview').innerHTML = formattedPreview;
      document.getElementById('result-section').style.display = 'block';
      window.resumeContent = data.savedResumeContent;
      
      // 使用保存的岗位名称
      window.jobTitle = data.savedJobTitle || '';
      
      // 显示下载按钮
      document.getElementById('download-btn').style.display = 'inline-block';
      
      // 显示成功状态消息
      showStatusMessage('已恢复上次生成的简历', 'success');
    }
  });
}

// 监听popup页面可见性变化，确保在切换标签页后返回时数据仍然存在
document.addEventListener('visibilitychange', function() {
  if (document.visibilityState === 'visible') {
    // 当页面变为可见时，重新加载保存的数据
    loadSavedData();
  }
});

// 监听窗口关闭事件
window.addEventListener('beforeunload', function() {
  // 在窗口关闭前保存当前状态
  const jobDescription = document.getElementById('job-description').value;
  if (jobDescription) {
    chrome.storage.local.set({ 'savedJobDescription': jobDescription });
  }
});

// 生成简历
async function generateResume() {
  const jobTitle = document.getElementById('job-title').value.trim();
  const jobDescription = document.getElementById('job-description').value.trim();
  
  if (!jobDescription) {
    showStatusMessage('请粘贴岗位信息', 'error');
    return;
  }
  
  // 获取用户数据
  const userData = {
    name: document.getElementById('user-name').value.trim(),
    contact: document.getElementById('user-contact').value.trim()
  };
  
  // 保存用户数据
  saveUserData(userData);
  
  // 保存当前招聘信息
  chrome.storage.local.set({ 
    'savedJobTitle': jobTitle,
    'savedJobDescription': jobDescription 
  });
  
  // 显示加载状态
  showStatusMessage('正在生成简历，请稍候...', 'loading');
  
  // 显示进度条并开始模拟进度
  showProgressBar(true);
  simulateProgress();
  
  try {
    // 提交到background.js处理
    const response = await chrome.runtime.sendMessage({
      action: 'generateResume',
      jobDescription: jobTitle ? `岗位名称: ${jobTitle}\n\n${jobDescription}` : jobDescription,
      userData: userData
    });
    
    if (response.success) {
      // 完成进度条
      completeProgress();
      
      // 显示生成的简历，并移除```markdown字符
      let formattedResume = response.resume.replace(/\n/g, '<br>');
      // 移除```markdown字符
      formattedResume = formattedResume.replace(/```markdown/g, '');
      formattedResume = formattedResume.replace(/```/g, '');
      document.getElementById('resume-preview').innerHTML = formattedResume;
      document.getElementById('result-section').style.display = 'block';
      showStatusMessage('简历已生成', 'success');
      
      // 保存简历内容作为全局变量
      window.resumeContent = response.resume;
      
      // 使用输入的岗位名称或从描述中提取
      window.jobTitle = jobTitle || '';
      if (!window.jobTitle) {
        const jobTitleMatch = jobDescription.match(/(?:职位|岗位|招聘)[:：]\s*([^\n,，。]+)/);
        window.jobTitle = jobTitleMatch ? jobTitleMatch[1].trim() : '';
      }
      
      // 保存预览内容，以便在UI界面重新打开时恢复
      chrome.storage.local.set({ 
        'savedResumePreview': formattedResume,
        'savedResumeContent': response.resume
      });
      
      // 显示下载按钮
      document.getElementById('download-btn').style.display = 'inline-block';
    } else {
      // 隐藏进度条
      showProgressBar(false);
      showStatusMessage('生成失败: ' + response.error, 'error');
    }
  } catch (error) {
    // 隐藏进度条
    showProgressBar(false);
    showStatusMessage('发生错误: ' + error.message, 'error');
    console.error(error);
  }
}

// 下载Word文档
// 免费生成功能 - 跳转到mujicv.com
function downloadResume() {
  // 直接跳转到mujicv.com网站，不再检查是否已生成简历
  chrome.tabs.create({ url: 'https://mujicv.com' });
  
  // 不再要求先生成简历，直接复制当前内容（如果有）
  if (window.resumeContent) {
    // 创建临时textarea元素来复制内容
    const textarea = document.createElement('textarea');
    textarea.value = window.resumeContent;
    document.body.appendChild(textarea);
    textarea.select();
    
    try {
      // 执行复制命令
      const successful = document.execCommand('copy');
      if (successful) {
        showStatusMessage('简历内容已复制到剪贴板', 'success');
      }
    } catch (err) {
      console.error('复制失败: ' + err);
    }
    
    // 移除临时元素
    document.body.removeChild(textarea);
  }
}


// 显示状态消息

// 显示进度条
function showProgressBar(show = true) {
  const progressContainer = document.getElementById('progress-container');
  progressContainer.style.display = show ? 'block' : 'none';
}

// 更新进度条
function updateProgress(percent) {
  const progressBar = document.getElementById('progress-bar');
  const progressText = document.getElementById('progress-text');
  progressBar.style.width = percent + '%';
  progressText.textContent = percent + '%';
}

// 模拟进度更新
function simulateProgress() {
  let progress = 0;
  showProgressBar(true);
  updateProgress(progress);
  
  window.progressInterval = setInterval(() => {
    if (progress >= 90) {
      clearInterval(window.progressInterval);
      return;
    }
    
    // 非线性增长，使进度条在开始时快速增长，然后减慢
    if (progress < 30) {
      progress += 5;
    } else if (progress < 60) {
      progress += 3;
    } else {
      progress += 1;
    }
    
    updateProgress(progress);
  }, 300);
}

// 完成进度
function completeProgress() {
  if (window.progressInterval) {
    clearInterval(window.progressInterval);
  }
  updateProgress(100);
  setTimeout(() => {
    showProgressBar(false);
  }, 1000);
}


function showStatusMessage(message, type) {
  const statusElement = document.getElementById('status-msg');
  statusElement.textContent = message;
  statusElement.className = 'status-message ' + type;
  statusElement.style.display = 'block';
}

// 保存用户数据
function saveUserData(userData) {
  chrome.storage.local.set({ userData: userData });
}



// 添加设置API Key的部分
function addSettingsSection() {
  const settingsSection = document.createElement('div');
  settingsSection.className = 'settings-section';
  settingsSection.innerHTML = `
    <h3>SiliconFlow API设置</h3>
    <div class="form-group">
      <label for="api-key">API Key:</label>
      <input type="password" id="api-key">
    </div>
    <button id="save-api-key">保存API Key</button>
    <button id="apply-free-btn" style="margin-left: 10px; background-color: #27ae60;">点我免费申请</button>
  `;
  
  // 插入到个人信息部分后面
  const personalInfoSection = document.querySelector('.settings-section');
  personalInfoSection.parentNode.insertBefore(settingsSection, personalInfoSection.nextSibling);
  
  // 绑定事件
  document.getElementById('save-api-key').addEventListener('click', saveApiKey);
  document.getElementById('apply-free-btn').addEventListener('click', function() {
    window.open('https://account.siliconflow.cn/zh/login?redirect=https%3A%2F%2Fcloud.siliconflow.cn&invitation=imSYdTsF', '_blank');
  });
  
  // 加载已保存的API Key
  chrome.storage.local.get('apiKey', function(data) {
    if (data.apiKey) {
      document.getElementById('api-key').value = data.apiKey;
    }
  });
}

// 保存API Key
function saveApiKey() {
  const apiKey = document.getElementById('api-key').value.trim();
  if (!apiKey) {
    showStatusMessage('请输入有效的API Key', 'error');
    return;
  }
  
  chrome.storage.local.set({ apiKey: apiKey });
  showStatusMessage('API Key已保存', 'success');
}

// 复制简历内容
function copyResumeContent() {
  const resumePreview = document.getElementById('resume-preview');
  if (!resumePreview || !resumePreview.innerHTML) {
    showStatusMessage('没有可复制的内容', 'error');
    return;
  }
  
  // 创建一个临时textarea元素来复制HTML内容
  const textarea = document.createElement('textarea');
  textarea.value = window.resumeContent || resumePreview.innerText.replace(/<br>/g, '\n');
  document.body.appendChild(textarea);
  textarea.select();
  
  try {
    // 执行复制命令
    const successful = document.execCommand('copy');
    if (successful) {
      showStatusMessage('简历内容已复制到剪贴板', 'success');
    } else {
      showStatusMessage('复制失败，请手动选择并复制', 'error');
    }
  } catch (err) {
    showStatusMessage('复制失败: ' + err, 'error');
  }
  
  // 移除临时元素
  document.body.removeChild(textarea);
}