// 监听来自popup.js的消息
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === "generateResume") {
    generateResume(request.jobDescription, request.userData)
      .then(result => sendResponse({ success: true, resume: result }))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true; // 异步响应
  }
  // 删除以下代码块
  // else if (request.action === "openFileLocation") {
  //   // 打开文件位置
  //   const filePath = request.filePath;
  //   if (filePath) {
  //     chrome.tabs.create({ url: `file:///${filePath.replace(/\\/g, '/')}` }, (tab) => {
  //       sendResponse({ success: true });
  //     });
  //     return true;
  //   } else {
  //     sendResponse({ success: false, error: "无效的文件路径" });
  //   }
  // } else if (request.action === "openSaveFolder") {
  //   // 打开下载文件夹
  //   chrome.downloads.showDefaultFolder();
  //   sendResponse({ success: true });
  //   return true;
  // }
});

// 调用SiliconFlow API生成简历
async function generateResume(jobDescription, userData) {
  // 从存储中获取API密钥
  const data = await chrome.storage.local.get('apiKey');
  const apiKey = data.apiKey;
  
  if (!apiKey) {
    throw new Error('请先在设置中配置API密钥');
  }
  
  const prompt = `基于以下招聘信息，请为${userData.name || '求职者'}创建一份专业美观的简历。
  
招聘信息:
${jobDescription}

请创建一份针对此职位量身定制的专业简历，严格按照以下格式和结构：

# ${userData.name || '求职者'} - [职位名称] - [技术/专业领域]

::: left
icon:info [性别]/[出生年月]
icon:weixin [微信号]
:::

::: right
[icon:blog 家庭住址在](博客URL)
[icon:phone 手机号](GitHub URL)
:::

## 教育背景

:::left
**[学校名称] - [专业名称]**
:::
:::right
**[入学年月] - [毕业年月]**
:::
[学习成绩、获奖情况、重要项目经历等]

## 工作/实习经验

:::left
**[公司名称] - [部门名称] - [职位名称]**
:::
:::right
**[开始年月] - [结束年月]**
:::

- [工作职责和成就，使用量化数据展示成果]
- [技术难点攻克，使用具体技术名称]
- [项目管理或团队协作经验]

## 项目经验

### [项目名称]

\`[技术1]\` \`[技术2]\` \`[技术3]\` \`[技术4]\` \`[技术5]\`

- **项目描述**：
    [简洁描述项目的目标、规模和你的角色]
- **工作内容**：
  - [具体工作内容，强调技术实现和解决方案]
  - [性能优化，包含具体的优化指标]
  - [架构设计或重构经验]
  - [使用的框架、工具和方法论]

## 技能

- **[技能类别1]**：[详细描述掌握的技能，包括熟练程度]
- **[技能类别2]**：[详细描述掌握的技能，包括熟练程度]
- **[技能类别3]**：[详细描述掌握的技能，包括熟练程度]

请根据招聘信息中的要求，填充上述模板中的所有内容，确保内容真实、专业、有针对性。使用markdown格式，保持布局美观。请直接输出简历内容，不要有其他解释。`;

  try {
    const response = await fetch('https://api.siliconflow.cn/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: "deepseek-ai/DeepSeek-R1",
        messages: [
          {
            role: "system",
            content: "你是一位专业的招聘顾问和简历撰写专家。你擅长创建结构清晰、布局美观的简历。请严格按照提供的格式模板生成简历内容，保持左右对齐的布局结构，正确使用图标标记和Markdown格式。确保生成的内容专业、有针对性，并且视觉上美观易读。"
          },
          {
            role: "user",
            content: prompt
          }
        ],
        temperature: 0.7
      })
    });
    
    const data = await response.json();
    
    if (data.error) {
      throw new Error(`API错误: ${data.error.message}`);
    }
    
    return data.choices[0].message.content;
  } catch (error) {
    console.error('API调用失败:', error);
    throw new Error('简历生成失败: ' + error.message);
  }
}

// 创建并下载Word文档
async function createAndDownloadDocx(resumeContent, jobTitle) {
  await loadDocxLibrary();
  
  try {
    // 直接使用全局docx对象，而不是window.docx
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle } = docx;
    
    // 将简历内容转换为段落
    const lines = resumeContent.split('\n');
    const docElements = [];
    
    let inHeading = false;
    
    lines.forEach(line => {
      const trimmedLine = line.trim();
      
      if (trimmedLine === '') {
        // 空行
        docElements.push(new Paragraph({}));
      } else if (/^#{1,3}/.test(trimmedLine) || /^[一二三四五六七八]、/.test(trimmedLine) || 
                (trimmedLine.length < 20 && (trimmedLine.endsWith(':') || trimmedLine.endsWith('：')))) {
        // 标题样式
        inHeading = true;
        docElements.push(
          new Paragraph({
            text: trimmedLine.replace(/^#{1,3} /, ''),
            heading: HeadingLevel.HEADING_2,
            thematicBreak: true,
            spacing: {
              before: 240,
              after: 120
            }
          })
        );
      } else {
        // 普通段落或列表项
        const isListItem = trimmedLine.startsWith('-') || trimmedLine.startsWith('*') || /^[0-9]+\./.test(trimmedLine);
        
        docElements.push(
          new Paragraph({
            text: trimmedLine,
            bullet: isListItem ? {level: 0} : undefined,
            indent: isListItem ? {left: 720} : undefined,
            spacing: {
              line: 360
            }
          })
        );
        inHeading = false;
      }
    });
    
    // 创建文档
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            text: "智能生成的简历",
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: {
              after: 200
            }
          }),
          ...docElements
        ],
      }],
    });
    
    // 生成文件内容 - 使用浏览器兼容的方式
    let blob;
    try {
      // 尝试使用各种可能的方法，按优先级排序
      if (typeof Packer.toBlob === 'function') {
        // 1. 首选：直接使用toBlob方法（最新版本docx.js支持）
        blob = await Packer.toBlob(doc);
      } else if (typeof Packer.toArrayBuffer === 'function') {
        // 2. 次选：使用ArrayBuffer（大多数浏览器环境支持）
        const arrayBuffer = await Packer.toArrayBuffer(doc);
        blob = new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      } else if (typeof Packer.toBase64String === 'function') {
        // 3. 备选：使用Base64字符串（几乎所有环境都支持）
        const base64 = await Packer.toBase64String(doc);
        const binaryString = atob(base64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      }
    } catch (error) {
      console.error('创建文档Blob失败:', error);
      throw new Error('创建Word文档失败: 浏览器环境不支持此操作，请尝试更新docx.js库');
    }
    
    // 提取职位名称作为文件名的一部分
    let jobTitleText = "智能简历";
    if (jobTitle && jobTitle.trim()) {
      jobTitleText = jobTitle.trim().replace(/[\/:*?"<>|]/g, '_').substring(0, 20);
    }
    
    // 生成文件名:智能简历_职位名称_日期.docx
    const date = new Date();
    const dateStr = `${date.getFullYear()}${(date.getMonth()+1).toString().padStart(2, '0')}${date.getDate().toString().padStart(2, '0')}`;
    const filename = `智能简历_${jobTitleText}_${dateStr}.docx`;
    
    // 使用saveAs让用户自主选择保存位置和文件名
    const result = await saveToFile(blob, filename);
    
    return result;
  } catch (error) {
    console.error('创建Word文档失败:', error);
    throw new Error('创建Word文档失败: ' + error.message);
  }
}

// 在Service Worker初始化时预加载docx.js库
try {
  importScripts(chrome.runtime.getURL('docx.js'));
  console.log('docx.js库已成功预加载');
} catch (error) {
  console.error('预加载docx.js库失败:', error);
}

// 加载docx.js库
async function loadDocxLibrary() {
  // 在后台脚本中不能使用window对象，直接导入docx.js
  if (typeof docx !== 'undefined') return;
  
  return new Promise((resolve, reject) => {
    try {
      // 检查docx是否已经被预加载
      if (typeof docx !== 'undefined') {
        resolve();
        return;
      }
      
      // 如果预加载失败，尝试使用fetch API加载docx.js
      fetch(chrome.runtime.getURL('docx.js'))
        .then(response => response.text())
        .then(scriptText => {
          // 使用Function构造函数执行脚本
          new Function(scriptText)();
          if (typeof docx !== 'undefined') {
            resolve();
          } else {
            reject(new Error('加载docx.js库失败: 库加载后未定义docx对象'));
          }
        })
        .catch(error => {
          console.error('加载docx.js库失败:', error);
          reject(new Error('加载docx.js库失败: ' + error.message));
        });
    } catch (error) {
      console.error('加载docx.js库失败:', error);
      reject(new Error('加载docx.js库失败: ' + error.message));
    }
  });
}

// 使用Chrome下载API保存文件
async function saveToFile(blob, filename) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    chrome.downloads.download({
      url: url,
      filename: filename,
      saveAs: true, // 设置为true，让用户自主选择保存位置和文件名
      conflictAction: 'uniquify'
    }, (downloadId) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        // 监听下载完成事件，获取文件路径
        chrome.downloads.onChanged.addListener(function downloadListener(downloadDelta) {
          if (downloadDelta.id === downloadId && downloadDelta.state && downloadDelta.state.current === "complete") {
            chrome.downloads.search({id: downloadId}, function(downloads) {
              if (downloads && downloads.length > 0) {
                const filePath = downloads[0].filename;
                chrome.storage.local.set({ 'lastResumeFilePath': filePath });
                chrome.downloads.onChanged.removeListener(downloadListener);
                resolve({filename: filename, filePath: filePath});
              } else {
                resolve(filename);
              }
            });
          }
        });
      }
    });
  });
}