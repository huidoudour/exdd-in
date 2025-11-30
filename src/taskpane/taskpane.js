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
    const includeStripe = document.getElementById("include-stripe")?.checked || false; // 默认不启用隔行变色

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
          
          // 设置高级表格样式
          const tableRange = sheet.getRangeByIndexes(0, 0, data.length + 1, headers.length);
          
          // 应用高级边框样式（内外边框都需要）
          applyBorderStyle(tableRange, style);
          
          // 设置列宽（先设置固定宽度，再自动调整）
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
          
          // 应用我们自定义的边框样式到表格范围
          applyBorderStyle(tableRange, style);
          
          // 设置数据样式（在表格创建前应用，确保行高设置生效）
          applyDataStyle(dataRange, style, includeStripe);
          
          // 创建表格对象来启用筛选功能，但使用无样式表格
          const table = sheet.tables.add(tableRange, true);
          table.style = "TableStyleNone"; // 使用无样式表格，避免默认隔行色
          
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
  const first = parseInt(parts[0]);
  const second = parseInt(parts[1]);
  const third = parseInt(parts[2]);
  const fourth = parseInt(parts[3]);
  
  const ips = [];
  let a = first;
  let b = second;
  let c = third;
  let d = fourth;
  
  for (let i = 0; i < count; i++) {
    ips.push(`${a}.${b}.${c}.${d}`);
    
    // 递增IP地址
    d++;
    if (d > 255) {
      d = 0;
      c++;
      if (c > 255) {
        c = 0;
        b++;
        if (b > 255) {
          b = 0;
          a++;
          if (a > 255) {
            break; // IP地址耗尽
          }
        }
      }
    }
  }
  
  return ips;
}

// 应用标题样式
function applyHeaderStyle(range, style) {
  // 基础样式
  range.format.font.bold = true;
  range.format.font.size = 12;
  range.format.font.name = "微软雅黑";
  range.horizontalAlignment = "center";
  range.verticalAlignment = "center";
  range.format.rowHeight = 25;
  
  // 设置标题背景和字体颜色
  let backgroundColor = "EEEEEE";
  let fontColor = "000000";
  
  switch (style) {
    case 'blue':
      backgroundColor = "4472C4";
      fontColor = "FFFFFF";
      break;
    case 'green':
      backgroundColor = "70AD47";
      fontColor = "FFFFFF";
      break;
    case 'purple':
      backgroundColor = "7030A0";
      fontColor = "FFFFFF";
      break;
    case 'modern':
      backgroundColor = "0078D4";
      fontColor = "FFFFFF";
      // 现代风格使用较大字号
      range.format.font.size = 14;
      break;
    case 'classic':
      backgroundColor = "201F1E";
      fontColor = "FFFFFF";
      // 经典风格使用粗体和下划线
      range.format.font.underline = Excel.RangeUnderlineStyle.single;
      break;
  }
  
  range.format.fill.color = backgroundColor;
  range.format.font.color = fontColor;
}

// 应用数据样式
function applyDataStyle(range, style, includeStripe) {
  // 基础数据样式
  range.format.font.name = "微软雅黑";
  range.format.font.size = 11;
  range.horizontalAlignment = "center";
  range.verticalAlignment = "center";
  range.format.rowHeight = 22;
  
  // 设置数据字体颜色
  let fontColor = "000000";
  
  switch (style) {
    case 'blue':
    case 'modern':
      fontColor = "1A1A1A";
      break;
    case 'green':
      fontColor = "1A1A1A";
      break;
    case 'purple':
      fontColor = "1A1A1A";
      break;
    case 'classic':
      fontColor = "000000";
      break;
  }
  
  range.format.font.color = fontColor;
  
  // 应用隔行变色 - 使用条件格式代替循环访问rowCount
  if (includeStripe) {
    try {
      // 使用Excel的条件格式功能来实现隔行变色
      // 这样避免了需要加载rowCount属性
      const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      
      // 设置公式：=MOD(ROW(),2)=0 表示偶数行（Excel行号从1开始）
      conditionalFormat.customRule.formula = "=MOD(ROW(),2)=0";
      
      // 根据样式设置填充颜色
      let fillColor = "F5F5F5";
      switch (style) {
        case 'blue':
          fillColor = "F2F7FB";
          break;
        case 'green':
          fillColor = "F4F9F4";
          break;
        case 'purple':
          fillColor = "FDF5FF";
          break;
        case 'modern':
          fillColor = "F5F9FF";
          break;
        case 'classic':
          fillColor = "F8F8F8";
          break;
      }
      
      conditionalFormat.format.fill.color = fillColor;
    } catch (error) {
      // 如果条件格式设置失败，使用备用方案
      console.log("条件格式设置失败，使用备用方案：", error);
      
      // 备用方案：使用简单的隔行变色，但不依赖rowCount
      // 这里我们暂时不实现备用方案，避免复杂的异步操作
      // 条件格式失败时，表格仍然可以正常创建，只是没有隔行变色效果
    }
  }
}