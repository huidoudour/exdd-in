/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("create-template").onclick = createIPTemplate;
    
    // 设置默认值
    document.getElementById("start-ip").value = "192.168.1.1";
  }
});

export async function createIPTemplate() {
  try {
    // 显示加载状态
    showFeedbackMessage('正在创建IP对应表模板...', 'loading');
    
    const rowsCount = parseInt(document.getElementById("rows-count").value) || 10;
    const startIP = document.getElementById("start-ip").value.trim();
    const style = document.getElementById("column-style").value;
    const includeMac = document.getElementById("include-mac").checked;
    const includeDesc = document.getElementById("include-desc").checked;
    const includeStripe = document.getElementById("include-stripe")?.checked || true; // 默认启用隔行变色

    // 验证输入
    if (!startIP) {
      showFeedbackMessage('请输入起始IP地址', 'error');
      return;
    }

    if (!isValidIP(startIP)) {
      showFeedbackMessage('IP地址格式不正确，请输入有效的IP地址（如：192.168.1.1）', 'error');
      return;
    }

    await Excel.run(async (context) => {
      // 获取活动工作表
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // 设置工作表名称
      sheet.name = "IP对应表";
      
      // 清除现有内容（从A1开始）
      const clearRange = sheet.getRange("A1:Z100");
      clearRange.clear();

      // 设置列标题
      const headers = ["序号", "IP地址"];
      if (includeMac) headers.push("MAC地址");
      if (includeDesc) headers.push("描述");

      // 写入标题行
      const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
      headerRange.values = [headers];
      
      // 应用高级标题样式
      applyHeaderStyle(headerRange, style);

      // 生成IP数据
      const ipData = generateIPData(startIP, rowsCount);
      
      // 准备数据
      const data = [];
      for (let i = 0; i < rowsCount; i++) {
        const row = [i + 1, ipData[i]];
        if (includeMac) row.push(""); // MAC地址列留空，用户填写
        if (includeDesc) row.push(""); // 描述列留空，用户填写
        data.push(row);
      }

      // 写入数据
      if (data.length > 0) {
        const dataRange = sheet.getRangeByIndexes(1, 0, data.length, headers.length);
        dataRange.values = data;
        
        // 设置数据样式
        applyDataStyle(dataRange, style);
        
        // 设置高级表格样式
        const tableRange = sheet.getRangeByIndexes(0, 0, data.length + 1, headers.length);
        
        // 应用高级边框样式
        applyBorderStyle(tableRange, style);
        
        // 自动调整列宽
        sheet.getUsedRange().format.autofitColumns();
        
        // 设置列宽
        sheet.getRange("A:A").format.columnWidth = 8; // 序号列
        sheet.getRange("B:B").format.columnWidth = 18; // IP地址列
        if (includeMac) {
          sheet.getRange("C:C").format.columnWidth = 22; // MAC地址列
        }
        if (includeDesc) {
          const descColumn = sheet.getRangeByIndexes(0, headers.length - 1, data.length + 1, 1);
          descColumn.horizontalAlignment = "left";
          descColumn.format.columnWidth = 30; // 描述列
        }
        
        // 添加表格筛选功能
        tableRange.applyFilters();
        
        // 添加条件格式（突出显示特定IP范围）
        addConditionalFormatting(dataRange, style);
      }

      await context.sync();
      console.log("IP对应表模板创建成功");
      
      showFeedbackMessage(`IP对应表模板创建成功！共 ${rowsCount} 行数据。`, 'success');
    });
  } catch (error) {
    console.error(error);
    showFeedbackMessage("创建模板时发生错误：" + error.message, 'error');
  }
}

// 显示反馈消息
function showFeedbackMessage(message, type) {
  const feedbackSection = document.getElementById('feedback-message');
  
  // 清除现有内容
  feedbackSection.innerHTML = '';
  
  // 设置消息样式
  const messageDiv = document.createElement('div');
  messageDiv.className = `feedback-message feedback-${type}`;
  messageDiv.textContent = message;
  
  // 添加到页面
  feedbackSection.appendChild(messageDiv);
  
  // 如果是成功或错误消息，3秒后自动消失
  if (type === 'success' || type === 'error') {
    setTimeout(() => {
      feedbackSection.innerHTML = '';
    }, 3000);
  }
}

// 应用边框样式
function applyBorderStyle(range, style) {
  // 基础边框设置
  range.format.borders.getItem("EdgeTop").lineStyle = "thin";
  range.format.borders.getItem("EdgeBottom").lineStyle = "thin";
  range.format.borders.getItem("EdgeLeft").lineStyle = "thin";
  range.format.borders.getItem("EdgeRight").lineStyle = "thin";
  range.format.borders.getItem("InsideHorizontal").lineStyle = "thin";
  range.format.borders.getItem("InsideVertical").lineStyle = "thin";
  
  // 根据主题设置边框颜色
  let borderColor = "CCCCCC";
  switch (style) {
    case 'blue':
      borderColor = "B6D0F6";
      break;
    case 'green':
      borderColor = "C3E6CB";
      break;
    case 'purple':
      borderColor = "D6BBFB";
      break;
    case 'modern':
      borderColor = "0078D4";
      // 现代风格使用更细的边框
      range.format.borders.getItem("EdgeTop").lineStyle = "hair";
      range.format.borders.getItem("EdgeBottom").lineStyle = "hair";
      range.format.borders.getItem("EdgeLeft").lineStyle = "hair";
      range.format.borders.getItem("EdgeRight").lineStyle = "hair";
      break;
    case 'classic':
      borderColor = "000000";
      // 经典风格使用更粗的外边框
      range.format.borders.getItem("EdgeTop").lineStyle = "medium";
      range.format.borders.getItem("EdgeBottom").lineStyle = "medium";
      range.format.borders.getItem("EdgeLeft").lineStyle = "medium";
      range.format.borders.getItem("EdgeRight").lineStyle = "medium";
      break;
  }
  
  // 应用边框颜色
  range.format.borders.getItem("EdgeTop").color = borderColor;
  range.format.borders.getItem("EdgeBottom").color = borderColor;
  range.format.borders.getItem("EdgeLeft").color = borderColor;
  range.format.borders.getItem("EdgeRight").color = borderColor;
  range.format.borders.getItem("InsideHorizontal").color = borderColor;
  range.format.borders.getItem("InsideVertical").color = borderColor;
}

// 添加条件格式
function addConditionalFormatting(range, style) {
  try {
    // 为IP地址列添加条件格式（假设是第二列）
    const ipColumn = range.getColumn(1); // 索引从0开始
    
    // 添加数据条格式
    const dataBarFormat = ipColumn.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    
    // 根据主题设置数据条颜色
    let dataBarColor = "4472C4";
    switch (style) {
      case 'blue':
        dataBarColor = "4472C4";
        break;
      case 'green':
        dataBarColor = "70AD47";
        break;
      case 'purple':
        dataBarColor = "7030A0";
        break;
      case 'modern':
        dataBarColor = "0078D4";
        break;
      case 'classic':
        dataBarColor = "201F1E";
        break;
    }
    
    dataBarFormat.dataBar.barColor = dataBarColor;
    dataBarFormat.dataBar.showValue = true;
    dataBarFormat.dataBar.axisColor = "FFFFFF";
  } catch (error) {
    // 条件格式功能可能在某些Excel版本不支持，忽略错误
    console.log("条件格式设置失败，可能不支持该功能：", error);
  }
}

// 验证IP地址格式
function isValidIP(ip) {
  const ipRegex = /^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$/;
  return ipRegex.test(ip);
}

// 生成IP地址列表
function generateIPData(startIP, count) {
  const parts = startIP.split('.');
  const baseIP = parts.slice(0, 3).join('.');
  const lastOctet = parseInt(parts[3]);
  
  const ips = [];
  for (let i = 0; i < count; i++) {
    const currentOctet = lastOctet + i;
    if (currentOctet > 255) {
      // 如果超过255，重置为1并增加第三个八位组
      const thirdOctet = parseInt(parts[2]) + Math.floor(currentOctet / 256);
      const finalOctet = currentOctet % 256;
      if (thirdOctet > 255) {
        // 超出范围，停止生成
        break;
      }
      ips.push(`${parts[0]}.${parts[1]}.${thirdOctet}.${finalOctet}`);
    } else {
      ips.push(`${baseIP}.${currentOctet}`);
    }
  }
  return ips;
}

// 应用标题样式
function applyHeaderStyle(range, style) {
  // 基础样式设置
  range.format.font.bold = true;
  range.format.font.size = 12;
  range.format.font.name = "微软雅黑";
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  
  // 设置标题行高度
  range.format.rowHeight = 24;
  
  // 根据主题设置颜色和高级样式
  switch (style) {
    case 'blue':
      range.format.fill.color = "4472C4";
      range.format.font.color = "FFFFFF";
      // 添加蓝色主题特有的边框效果
      range.format.borders.getItem("EdgeBottom").lineStyle = "medium";
      range.format.borders.getItem("EdgeBottom").color = "2E5984";
      break;
    case 'green':
      range.format.fill.color = "70AD47";
      range.format.font.color = "FFFFFF";
      // 添加绿色主题特有的边框效果
      range.format.borders.getItem("EdgeBottom").lineStyle = "medium";
      range.format.borders.getItem("EdgeBottom").color = "4B7F2E";
      break;
    case 'purple':
      range.format.fill.color = "7030A0";
      range.format.font.color = "FFFFFF";
      // 添加紫色主题特有的边框效果
      range.format.borders.getItem("EdgeBottom").lineStyle = "medium";
      range.format.borders.getItem("EdgeBottom").color = "4A206F";
      break;
    case 'modern':
      // 现代简约风格
      range.format.fill.color = "0078D4";
      range.format.font.color = "FFFFFF";
      range.format.font.size = 13;
      range.format.borders.getItem("EdgeBottom").lineStyle = "medium";
      range.format.borders.getItem("EdgeBottom").color = "005A9E";
      break;
    case 'classic':
      // 经典商务风格
      range.format.fill.color = "201F1E";
      range.format.font.color = "FFFFFF";
      range.format.font.name = "Arial";
      range.format.borders.getItem("EdgeBottom").lineStyle = "thick";
      range.format.borders.getItem("EdgeBottom").color = "FFFFFF";
      break;
    default:
      range.format.fill.color = "D9D9D9";
      range.format.font.color = "000000";
      range.format.borders.getItem("EdgeBottom").lineStyle = "medium";
      range.format.borders.getItem("EdgeBottom").color = "A6A6A6";
  }
}

// 应用数据样式
function applyDataStyle(range, style) {
  // 基础数据样式设置
  range.format.font.name = "微软雅黑";
  range.format.font.size = 11;
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  
  // 设置数据行高度
  range.format.rowHeight = 22;
  
  // 添加隔行变色效果
  const rowCount = range.rowCount;
  for (let i = 0; i < rowCount; i++) {
    const rowRange = range.getRow(i);
    
    // 偶数行使用不同的背景色
    if (i % 2 === 1) {
      switch (style) {
        case 'blue':
          rowRange.format.fill.color = "F0F7FF";
          break;
        case 'green':
          rowRange.format.fill.color = "F4F9F4";
          break;
        case 'purple':
          rowRange.format.fill.color = "FDF5FF";
          break;
        case 'modern':
          rowRange.format.fill.color = "F5F9FF";
          break;
        case 'classic':
          rowRange.format.fill.color = "F8F8F8";
          break;
        default:
          rowRange.format.fill.color = "F5F5F5";
      }
    } else {
      // 奇数行背景色
      switch (style) {
        case 'blue':
          rowRange.format.fill.color = "FFFFFF";
          break;
        case 'green':
          rowRange.format.fill.color = "FFFFFF";
          break;
        case 'purple':
          rowRange.format.fill.color = "FFFFFF";
          break;
        case 'modern':
          rowRange.format.fill.color = "FFFFFF";
          break;
        case 'classic':
          rowRange.format.fill.color = "FFFFFF";
          break;
        default:
          rowRange.format.fill.color = "FFFFFF";
      }
    }
  }
  
  // 添加行悬停效果的提示（通过注释说明）
  // 注意：Excel JavaScript API 目前不支持直接添加行悬停效果