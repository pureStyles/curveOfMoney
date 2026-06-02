import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import ReactECharts from 'echarts-for-react';

export default function TradeDashboard() {
  // 从 LocalStorage 初始化数据，防止刷新页面丢失
  const [dailyData, setDailyData] = useState(() => {
    const saved = localStorage.getItem('curve_of_money_data');
    return saved ? JSON.parse(saved) : [];
  });

  const [metrics, setMetrics] = useState({ winRate: 0, plRatio: 0, totalFee: 0 });

  // 监听数据变化，实时同步到本地缓存并重新计算综合指标
  useEffect(() => {
    localStorage.setItem('curve_of_money_data', JSON.stringify(dailyData));
    calculateMetrics(dailyData);
  }, [dailyData]);

  // 解析上传的 Excel
  const handleFileUpload = async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;

    console.log("=== 开始解析文件，总计:", files.length);
    const updatedData = [...dailyData];

    for (const file of files) {
      try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        
        let date = '';
        let balance = 0;
        let fee = 0;
        let dayTrades = [];

        // 1. 解析【客户交易结算日报】页面
        const infoSheet = workbook.Sheets['客户交易结算日报'];
        if (infoSheet) {
          const json = XLSX.utils.sheet_to_json(infoSheet, { header: 1 });
          json.forEach(row => {
            if (!row || row.length === 0) return;
            
            // 全体转成字符串并过滤掉空单元格，防止跨多列时产生的连续空值干扰索引
            const cleanRow = row.map(cell => cell !== undefined && cell !== null ? String(cell).trim() : "").filter(c => c !== "");
            
            // 匹配：交易日期
            const dateIdx = cleanRow.findIndex(cell => cell.includes('交易日期'));
            if (dateIdx !== -1 && cleanRow[dateIdx + 1]) {
              date = cleanRow[dateIdx + 1];
            }

            // 匹配：客户权益
            const eqIdx = cleanRow.findIndex(cell => cell.includes('客户权益'));
            if (eqIdx !== -1 && cleanRow[eqIdx + 1]) {
              balance = parseFloat(cleanRow[eqIdx + 1]) || 0;
            }

            // 匹配：当日手续费
            const feeIdx = cleanRow.findIndex(cell => cell.includes('当日手续费'));
            if (feeIdx !== -1 && cleanRow[feeIdx + 1]) {
              fee = parseFloat(cleanRow[feeIdx + 1]) || 0;
            }
          });
        }

        // 2. 解析【平仓明细】页面
        const closeSheet = workbook.Sheets['平仓明细'];
        if (closeSheet) {
          const json = XLSX.utils.sheet_to_json(closeSheet, { header: 1 });
          let profitIdx = -1;
          let startParsing = false;

          json.forEach(row => {
            if (!row || row.length === 0) return;
            
            const rowStr = row.map(c => String(c)).join(',');
            
            // 模糊匹配表头行，只要包含合约和平仓盈亏即可
            if (rowStr.includes('平仓盈亏') && rowStr.includes('合约')) {
              // 寻找真实的列索引
              profitIdx = row.findIndex(cell => String(cell).includes('平仓盈亏'));
              startParsing = true;
              return;
            }

            // 数据提取行：避开合计和开头描述
            if (startParsing) {
              const firstCell = String(row[0] || '').trim();
              if (firstCell && !firstCell.includes('合计') && !firstCell.includes('基本资料')) {
                const val = parseFloat(row[profitIdx]);
                if (!isNaN(val)) {
                  dayTrades.push(val);
                }
              }
            }
          });
        }

        console.log(`[单日解析成功] 日期: ${date}, 权益: ${balance}, 手续费: ${fee}, 平仓单数: ${dayTrades.length}`);

        // 只要能成功提取到日期和账户权益就视作有效记录
        if (date && !isNaN(balance)) {
          const existingIndex = updatedData.findIndex(d => d.date === date);
          const dayRecord = { date, balance, fee, trades: dayTrades };

          if (existingIndex !== -1) {
            updatedData[existingIndex] = dayRecord; // 重复上传则更新覆盖
          } else {
            updatedData.push(dayRecord); // 新数据追加
          }
        } else {
          console.warn(`⚠️ 文件 ${file.name} 解析出来的日期或权益无效: Date=${date}, Balance=${balance}`);
        }

      } catch (err) {
        console.error(`❌ 文件 ${file.name} 解析过程中崩溃:`, err);
      }
    }

    // 重新按时间升序排列，确保曲线连续
    const sortedData = updatedData.sort((a, b) => new Date(a.date) - new Date(b.date));
    setDailyData(sortedData);
  };

  // 核心指标计算逻辑
  const calculateMetrics = (allDays) => {
    if (!allDays || allDays.length === 0) {
      setMetrics({ winRate: 0, plRatio: 0, totalFee: 0 });
      return;
    }

    const allTrades = allDays.flatMap(d => d.trades || []);
    const totalFee = allDays.reduce((sum, d) => sum + (d.fee || 0), 0);

    if (allTrades.length === 0) {
      setMetrics({ winRate: "0.00", plRatio: "0.00", totalFee: totalFee.toFixed(2) });
      return;
    }

    const wins = allTrades.filter(p => p > 0);
    const losses = allTrades.filter(p => p < 0);
    
    // 计算胜率
    const winRate = (wins.length / allTrades.length) * 100;
    
    // 计算盈亏比 (平均盈利 / 平均亏损)
    const avgWin = wins.length ? (wins.reduce((a, b) => a + b, 0) / wins.length) : 0;
    const avgLoss = losses.length ? Math.abs(losses.reduce((a, b) => a + b, 0) / losses.length) : 0;
    
    let plRatio = 0;
    if (avgLoss > 0) {
      plRatio = avgWin / avgLoss;
    } else if (avgWin > 0) {
      plRatio = avgWin; // 如果全胜，盈亏比直接等于平均盈利
    }

    setMetrics({
      winRate: winRate.toFixed(2),
      plRatio: plRatio.toFixed(2),
      totalFee: totalFee.toFixed(2)
    });
  };

  // 一键清空缓存
  const handleClearData = () => {
    if (window.confirm("确认要清空本地保存的所有历史交易曲线数据吗？")) {
      setDailyData([]);
      localStorage.removeItem('curve_of_money_data');
    }
  };

  // ECharts 渲染配置项
  const getChartOption = () => {
    const dates = dailyData.map(d => d.date);
    const balances = dailyData.map(d => d.balance);
    
    let currentTotalFee = 0;
    const cumulativeFees = dailyData.map(d => {
      currentTotalFee += (d.fee || 0);
      return currentTotalFee.toFixed(2);
    });

    return {
      title: { text: '账户资产与累计手续费曲线', left: 'center', textStyle: { fontSize: 16 } },
      tooltip: { trigger: 'axis', shared: true },
      legend: { data: ['客户权益 (资产)', '累计手续费'], top: '8%' },
      grid: { top: '20%', bottom: '12%', left: '8%', right: '8%' },
      xAxis: { type: 'category', data: dates, boundaryGap: false },
      yAxis: [
        { type: 'value', name: '账户资金/权益', position: 'left', scale: true },
        { type: 'value', name: '累计手续费', position: 'right', scale: true, splitLine: { show: false } }
      ],
      series: [
        {
          name: '客户权益 (资产)', type: 'line', data: balances, smooth: true,
          itemStyle: { color: '#3f51b5' },
          areaStyle: {
            color: {
              type: 'linear', x: 0, y: 0, x2: 0, y2: 1,
              colorStops: [{ offset: 0, color: 'rgba(63,81,181,0.2)' }, { offset: 1, color: 'rgba(63,81,181,0)' }]
            }
          }
        },
        {
          name: '累计手续费', type: 'line', yAxisIndex: 1, data: cumulativeFees, smooth: true,
          itemStyle: { color: '#ff9800' }
        }
      ]
    };
  };

  return (
    <div style={{ padding: '24px', fontFamily: 'sans-serif', maxWidth: '1200px', margin: '0 auto' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
        <h2 style={{ margin: 0 }}>📊 curveOfMoney - 交易资产看板</h2>
        {dailyData.length > 0 && (
          <button onClick={handleClearData} style={clearBtnStyle}>🗑️ 清空历史本地数据</button>
        )}
      </div>
      
      {/* 上传区域 */}
      <div style={uploadBoxStyle}>
        <label style={{ cursor: 'pointer', display: 'block', width: '100%' }}>
          📂 <strong>点击或拖拽上传 Excel 结算日报</strong> (支持多选历史多天账单)
          <input type="file" multiple accept=".xlsx, .xls" onChange={handleFileUpload} style={{ display: 'none' }} />
        </label>
      </div>

      {/* 指标展示卡片 */}
      <div style={{ display: 'flex', gap: '20px', marginBottom: '30px' }}>
        <div style={cardStyle}>
          <span style={labelStyle}>📊 胜率 (Win Rate)</span>
          <p style={{ fontSize: '28px', color: '#4caf50', margin: '6px 0', fontWeight: 'bold' }}>{metrics.winRate}%</p>
        </div>
        <div style={cardStyle}>
          <span style={labelStyle}>⚖️ 盈亏比 (P/L Ratio)</span>
          <p style={{ fontSize: '28px', color: '#2196f3', margin: '6px 0', fontWeight: 'bold' }}>{metrics.plRatio}</p>
        </div>
        <div style={cardStyle}>
          <span style={labelStyle}>💰 累计消耗手续费</span>
          <p style={{ fontSize: '28px', color: '#ff9800', margin: '6px 0', fontWeight: 'bold' }}>¥{metrics.totalFee}</p>
        </div>
      </div>

      {/* 图表视图切换 */}
      {dailyData.length >= 2 ? (
        <div style={chartContainerStyle}>
          <ReactECharts option={getChartOption()} style={{ height: '500px' }} />
        </div>
      ) : dailyData.length === 1 ? (
        <div style={hintBoxStyle}>
          <p>💡 <strong>已成功录入 1 天的数据！</strong> 资产权益：¥{dailyData[0].balance}。曲线图至少需要 <strong>2 天</strong> 的历史账单数据才能拉出连线。请继续上传其他日期的结算单 Excel 吧！</p>
        </div>
      ) : (
        <p style={{ color: '#999', textAlign: 'center', marginTop: '60px' }}>暂无历史交易数据，请上传结算单文件。</p>
      )}
    </div>
  );
}

const uploadBoxStyle = { marginBottom: '24px', padding: '20px', border: '2px dashed #3f51b5', borderRadius: '8px', textAlign: 'center', background: '#f9f9ff' };
const cardStyle = { flex: 1, padding: '16px', background: '#f5f5f5', borderRadius: '8px', borderLeft: '5px solid #3f51b5' };
const labelStyle = { color: '#666', fontSize: '14px' };
const chartContainerStyle = { background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.08)' };
const hintBoxStyle = { padding: '20px', background: '#fff8e1', borderLeft: '5px solid #ffb300', borderRadius: '4px', color: '#b78103' };
const clearBtnStyle = { padding: '8px 14px', background: '#ffebee', color: '#c62828', border: '1px solid #ffcdd2', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' };