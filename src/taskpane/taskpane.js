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
    const rowsCount = parseInt(document.getElementById("rows-count").value) || 10;
    const startIP = document.getElementById("start-ip").value.trim();
    const style = document.getElementById("column-style").value;
    const includeMac = document.getElementById("include-mac").checked;
    const includeDesc = document.getElementById("include-desc").checked;

    // 验证输入
    if (!startIP) {
      Office.context.ui.displayDialogAsync(
        "请输入起始IP地址",
        { height: 30, width: 300 }
      );
      return;
    }

    if (!isValidIP(startIP)) {
      Office.context.ui.displayDialogAsync(
        "IP地址格式不正确，请输入有效的IP地址（如：192.168.1.1）",
        { height: 30, width: 400 }
      );
      return;
    }

    await Excel.run(async (context) => {
      // 获取活动工作表
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
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
      
      // 设置标题样式
      headerRange.format.font.bold = true;
      headerRange.format.font.size = 12;
      
      // 根据选择的主题设置颜色
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
        
        // 自动调整列宽
        sheet.getUsedRange().columns.autoFit();
        
        // 设置边框
        const tableRange = sheet.getRangeByIndexes(0, 0, data.length + 1, headers.length);
        tableRange.format.borders.lineStyle = "thin";
        tableRange.format.borders.color = "CCCCCC";
        
        // 居中对齐
        tableRange.horizontalAlignment = "center";
        if (includeDesc) {
          // 描述列左对齐
          const descColumn = sheet.getRangeByIndexes(0, headers.length - 1, data.length + 1, 1);
          descColumn.horizontalAlignment = "left";
        }
      }

      await context.sync();
      console.log("IP对应表模板创建成功");
      
      Office.context.ui.displayDialogAsync(
        `IP对应表模板创建成功！共 ${rowsCount} 行数据。`,
        { height: 30, width: 350 }
      );
    });
  } catch (error) {
    console.error(error);
    Office.context.ui.displayDialogAsync(
      "创建模板时发生错误：" + error.message,
      { height: 30, width: 400 }
    );
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
  switch (style) {
    case 'blue':
      range.format.fill.color = "4472C4";
      range.format.font.color = "FFFFFF";
      break;
    case 'green':
      range.format.fill.color = "70AD47";
      range.format.font.color = "FFFFFF";
      break;
    case 'purple':
      range.format.fill.color = "7030A0";
      range.format.font.color = "FFFFFF";
      break;
    default:
      range.format.fill.color = "D9D9D9";
      range.format.font.color = "000000";
  }
}

// 应用数据样式
function applyDataStyle(range, style) {
  switch (style) {
    case 'blue':
      range.format.fill.color = "E7F1FF";
      break;
    case 'green':
      range.format.fill.color = "E8F5E8";
      break;
    case 'purple':
      range.format.fill.color = "F3E8FF";
      break;
    default:
      range.format.fill.color = "FFFFFF";
  }
}
