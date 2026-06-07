import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import ReactECharts from 'echarts-for-react';

const STORAGE_KEY = 'curve_of_money_data';
const NORMALIZED_BASE_URL = import.meta.env.BASE_URL.endsWith('/') ? import.meta.env.BASE_URL : `${import.meta.env.BASE_URL}/`;
const SHARED_DATA_URL = `${NORMALIZED_BASE_URL}data/trade-data.json`;

const parseDateValue = (value) => {
  if (value instanceof Date) return new Date(value);

  const text = String(value || '').trim();
  if (!text) return null;

  const compactMatch = text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (compactMatch) {
    const [, year, month, day] = compactMatch;
    return new Date(`${year}/${month}/${day}`);
  }

  const parsed = new Date(text.replace(/[.-]/g, '/'));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const getRangeFilteredData = (allDays, rangeKey) => {
  if (!allDays.length || rangeKey === 'ALL') return allDays;

  const latestDate = parseDateValue(allDays[allDays.length - 1].date);
  if (!latestDate) return allDays;

  const rangeStart = new Date(latestDate);

  if (rangeKey === 'WTD') {
    const dayOfWeek = rangeStart.getDay();
    const diffToMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    rangeStart.setDate(rangeStart.getDate() - diffToMonday);
  }

  if (rangeKey === 'MTD') {
    rangeStart.setDate(1);
  }

  rangeStart.setHours(0, 0, 0, 0);

  return allDays.filter((day) => {
    const currentDate = parseDateValue(day.date);
    return currentDate && currentDate >= rangeStart && currentDate <= latestDate;
  });
};

const mergeDailyRecords = (...sources) => {
  const merged = new Map();

  sources.flat().forEach((item) => {
    if (!item?.date) return;
    merged.set(item.date, {
      date: item.date,
      balance: Number(item.balance) || 0,
      fee: Number(item.fee) || 0,
      trades: Array.isArray(item.trades) ? item.trades : []
    });
  });

  return Array.from(merged.values()).sort((a, b) => parseDateValue(a.date) - parseDateValue(b.date));
};

const readLocalDailyData = () => {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    const parsed = saved ? JSON.parse(saved) : [];
    return Array.isArray(parsed) ? mergeDailyRecords(parsed) : [];
  } catch {
    return [];
  }
};

const normalizeSharedPayload = (payload) => {
  if (Array.isArray(payload)) {
    return { records: mergeDailyRecords(payload), updatedAt: '' };
  }

  if (payload && Array.isArray(payload.records)) {
    return {
      records: mergeDailyRecords(payload.records),
      updatedAt: payload.updatedAt || ''
    };
  }

  return { records: [], updatedAt: '' };
};

const getDerivedDailySeries = (allDays, fullDays = allDays) => {
  return allDays.map((day, index) => {
    const fullIndex = fullDays.findIndex((item) => item.date === day.date);
    const previousDay =
      fullIndex > 0
        ? fullDays[fullIndex - 1]
        : index > 0
          ? allDays[index - 1]
          : null;
    const previousBalance = previousDay ? Number(previousDay.balance) || 0 : 0;
    const balance = Number(day.balance) || 0;
    const profitAmount = previousDay ? balance - previousBalance : 0;
    const profitRate = previousDay && previousBalance !== 0 ? (profitAmount / previousBalance) * 100 : 0;

    return {
      ...day,
      profitAmount,
      profitRate
    };
  });
};

const formatCurrency = (value) => `¥${Number(value || 0).toFixed(2)}`;
const formatPercent = (value) => `${Number(value || 0).toFixed(2)}%`;
const formatDateLabel = (value) => {
  const date = parseDateValue(value);
  return date ? date.toLocaleDateString('zh-CN') : String(value || '');
};

export default function TradeDashboard() {
  // 从 LocalStorage 初始化数据，防止刷新页面丢失
  const [dailyData, setDailyData] = useState(() => readLocalDailyData());

  const [metrics, setMetrics] = useState({ winRate: 0, plRatio: 0, totalFee: 0 });
  const [selectedRange, setSelectedRange] = useState('ALL');
  const [isMobile, setIsMobile] = useState(() => window.innerWidth <= 768);
  const [sharedUpdatedAt, setSharedUpdatedAt] = useState('');
  const [hasHydratedSharedData, setHasHydratedSharedData] = useState(false);

  // 监听数据变化，实时同步到本地缓存
  useEffect(() => {
    if (!hasHydratedSharedData) return;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(dailyData));
  }, [dailyData, hasHydratedSharedData]);

  useEffect(() => {
    const handleResize = () => setIsMobile(window.innerWidth <= 768);
    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  useEffect(() => {
    const loadSharedData = async () => {
      const localData = readLocalDailyData();

      try {
        const response = await fetch(SHARED_DATA_URL, { cache: 'no-cache' });
        if (!response.ok) {
          throw new Error(`Failed to load shared data: ${response.status}`);
        }

        const payload = await response.json();
        const normalized = normalizeSharedPayload(payload);
        setSharedUpdatedAt(normalized.updatedAt);
        setDailyData(mergeDailyRecords(normalized.records, localData));
      } catch (error) {
        console.warn('共享数据文件加载失败，继续使用本地缓存。', error);
        setDailyData(localData);
      } finally {
        setHasHydratedSharedData(true);
      }
    };

    loadSharedData();
  }, []);

  const filteredDailyData = getRangeFilteredData(dailyData, selectedRange);
  const derivedDailyData = getDerivedDailySeries(filteredDailyData, dailyData);
  const latestDay = derivedDailyData[derivedDailyData.length - 1];

  // 数据或筛选区间变化时，重新计算综合指标
  useEffect(() => {
    calculateMetrics(filteredDailyData);
  }, [dailyData, selectedRange]);

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
    setDailyData(mergeDailyRecords(updatedData));
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
      localStorage.removeItem(STORAGE_KEY);
    }
  };

  const handleExportSharedData = () => {
    const payload = {
      updatedAt: new Date().toISOString(),
      records: mergeDailyRecords(dailyData)
    };

    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'trade-data.json';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  };

  // ECharts 渲染配置项
  const getChartOption = () => {
    const derivedData = derivedDailyData;
    const dates = derivedData.map(d => d.date);
    const balances = derivedData.map(d => Number(d.balance).toFixed(2));
    const profitAmounts = derivedData.map(d => d.profitAmount.toFixed(2));
    const profitRates = derivedData.map(d => d.profitRate.toFixed(2));

    return {
      title: { text: '权益资金曲线与每日收益表现', left: 'center', textStyle: { fontSize: 16 } },
      tooltip: {
        trigger: 'axis',
        axisPointer: { type: 'cross' },
        formatter: (params) => {
          const day = derivedData[params[0]?.dataIndex];
          if (!day) return '';

          return [
            `${day.date}`,
            `权益资金: ${formatCurrency(day.balance)}`,
            `每日收益金额: ${formatCurrency(day.profitAmount)}`,
            `每日收益率: ${formatPercent(day.profitRate)}`
          ].join('<br/>');
        },
        confine: true
      },
      legend: {
        data: ['权益资金曲线', '每日收益金额', '每日收益率'],
        top: isMobile ? '10%' : '8%',
        type: isMobile ? 'scroll' : 'plain'
      },
      grid: {
        top: isMobile ? '28%' : '20%',
        bottom: isMobile ? '18%' : '12%',
        left: isMobile ? '12%' : '8%',
        right: isMobile ? '20%' : '14%'
      },
      xAxis: {
        type: 'category',
        data: dates,
        boundaryGap: false,
        axisLabel: {
          fontSize: isMobile ? 10 : 12,
          hideOverlap: true
        }
      },
      yAxis: [
        {
          type: 'value',
          name: isMobile ? '权益' : '权益资金',
          position: 'left',
          scale: true,
          axisLabel: {
            formatter: (value) => isMobile ? `${(Number(value) / 10000).toFixed(1)}w` : formatCurrency(value)
          }
        },
        {
          type: 'value',
          name: isMobile ? '收益额' : '收益金额',
          position: 'right',
          scale: true,
          splitLine: { show: false },
          axisLabel: {
            formatter: (value) => isMobile ? `${(Number(value) / 10000).toFixed(1)}w` : formatCurrency(value)
          }
        },
        {
          type: 'value',
          name: isMobile ? '收益率' : '收益率',
          position: 'right',
          offset: isMobile ? 48 : 80,
          scale: true,
          splitLine: { show: false },
          axisLabel: { formatter: (value) => `${Number(value).toFixed(2)}%` }
        }
      ],
      series: [
        {
          name: '权益资金曲线',
          type: 'line',
          data: balances,
          smooth: true,
          itemStyle: { color: '#3f51b5' },
          areaStyle: {
            color: {
              type: 'linear', x: 0, y: 0, x2: 0, y2: 1,
              colorStops: [{ offset: 0, color: 'rgba(63,81,181,0.2)' }, { offset: 1, color: 'rgba(63,81,181,0)' }]
            }
          },
          lineStyle: { width: 3 }
        },
        {
          name: '每日收益金额',
          type: 'bar',
          yAxisIndex: 1,
          data: profitAmounts,
          itemStyle: {
            color: (params) => Number(params.value) >= 0 ? '#4caf50' : '#f44336'
          },
          barMaxWidth: 18
        },
        {
          name: '每日收益率',
          type: 'line',
          yAxisIndex: 2,
          data: profitRates,
          smooth: true,
          itemStyle: { color: '#ff9800' },
          lineStyle: { width: 2 },
          symbolSize: 7
        }
      ]
    };
  };
  const rangeOptions = [
    { key: 'ALL', label: '全部' },
    { key: 'WTD', label: 'WTD' },
    { key: 'MTD', label: 'MTD' }
  ];

  const pageStyle = {
    padding: isMobile ? '16px 12px 24px' : '24px',
    fontFamily: 'sans-serif',
    maxWidth: '1200px',
    margin: '0 auto'
  };
  const headerStyle = {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: isMobile ? 'stretch' : 'center',
    flexDirection: isMobile ? 'column' : 'row',
    gap: isMobile ? '12px' : 0,
    marginBottom: '20px'
  };
  const headerActionsStyle = {
    display: 'flex',
    gap: '10px',
    alignItems: 'center',
    flexDirection: isMobile ? 'column' : 'row',
    width: isMobile ? '100%' : 'auto'
  };
  const metricsWrapStyle = {
    display: 'flex',
    gap: isMobile ? '12px' : '20px',
    marginBottom: '30px',
    flexDirection: isMobile ? 'column' : 'row'
  };
  const summaryWrapStyle = {
    display: 'flex',
    gap: isMobile ? '12px' : '20px',
    marginBottom: '24px',
    flexDirection: isMobile ? 'column' : 'row'
  };

  return (
    <div style={pageStyle}>
      <div style={headerStyle}>
        <h2 style={{ margin: 0, fontSize: isMobile ? '22px' : '28px', lineHeight: 1.3 }}>📊 curveOfMoney - 交易资产看板</h2>
        <div style={headerActionsStyle}>
          {dailyData.length > 0 && (
            <button onClick={handleExportSharedData} style={isMobile ? mobileSecondaryBtnStyle : secondaryBtnStyle}>⬇️ 导出同步文件</button>
          )}
          {dailyData.length > 0 && (
            <button onClick={handleClearData} style={isMobile ? mobileClearBtnStyle : clearBtnStyle}>🗑️ 清空历史本地数据</button>
          )}
        </div>
      </div>
      
      {/* 上传区域 */}
      <div style={uploadBoxStyle}>
        <label style={{ cursor: 'pointer', display: 'block', width: '100%' }}>
          📂 <strong>点击或拖拽上传 Excel 结算日报</strong> (支持多选历史多天账单)
          <input type="file" multiple accept=".xlsx, .xls" onChange={handleFileUpload} style={{ display: 'none' }} />
        </label>
      </div>

      {dailyData.length > 0 && (
        <div style={isMobile ? mobileFilterBarStyle : filterBarStyle}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
            <span style={{ ...labelStyle, fontSize: '15px' }}>查看区间</span>
            <span style={{ color: '#7b879b', fontSize: '12px' }}>
              {sharedUpdatedAt ? `共享文件更新时间：${formatDateLabel(sharedUpdatedAt)}` : '当前使用本地缓存或尚未生成共享文件'}
            </span>
          </div>
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
            {rangeOptions.map((option) => (
              <button
                key={option.key}
                onClick={() => setSelectedRange(option.key)}
                style={selectedRange === option.key ? (isMobile ? mobileActiveFilterBtnStyle : activeFilterBtnStyle) : (isMobile ? mobileFilterBtnStyle : filterBtnStyle)}
              >
                {option.label}
              </button>
            ))}
            {filteredDailyData.length > 0 && (
              <span style={{ color: '#666', fontSize: '13px' }}>
                {formatDateLabel(filteredDailyData[0].date)} - {formatDateLabel(filteredDailyData[filteredDailyData.length - 1].date)}
              </span>
            )}
          </div>
        </div>
      )}

      {/* 指标展示卡片 */}
      <div style={metricsWrapStyle}>
        <div style={cardStyle}>
          <span style={labelStyle}>📊 胜率 (Win Rate)</span>
          <p style={{ fontSize: isMobile ? '24px' : '28px', color: '#4caf50', margin: '6px 0', fontWeight: 'bold' }}>{metrics.winRate}%</p>
        </div>
        <div style={cardStyle}>
          <span style={labelStyle}>⚖️ 盈亏比 (P/L Ratio)</span>
          <p style={{ fontSize: isMobile ? '24px' : '28px', color: '#2196f3', margin: '6px 0', fontWeight: 'bold' }}>{metrics.plRatio}</p>
        </div>
        <div style={cardStyle}>
          <span style={labelStyle}>💰 累计消耗手续费</span>
          <p style={{ fontSize: isMobile ? '24px' : '28px', color: '#ff9800', margin: '6px 0', fontWeight: 'bold' }}>¥{metrics.totalFee}</p>
        </div>
      </div>

      {derivedDailyData.length > 0 && (
        <div style={summaryWrapStyle}>
          <div style={cardStyle}>
            <span style={labelStyle}>📈 最新权益资金</span>
            <p style={{ fontSize: isMobile ? '24px' : '28px', color: '#3f51b5', margin: '6px 0', fontWeight: 'bold', wordBreak: 'break-word' }}>
              {formatCurrency(latestDay.balance)}
            </p>
          </div>
          <div style={cardStyle}>
            <span style={labelStyle}>🧾 最新日收益金额</span>
            <p style={{ fontSize: isMobile ? '24px' : '28px', color: latestDay.profitAmount >= 0 ? '#4caf50' : '#f44336', margin: '6px 0', fontWeight: 'bold', wordBreak: 'break-word' }}>
              {formatCurrency(latestDay.profitAmount)}
            </p>
          </div>
          <div style={cardStyle}>
            <span style={labelStyle}>📅 最新日收益率</span>
            <p style={{ fontSize: isMobile ? '24px' : '28px', color: latestDay.profitRate >= 0 ? '#ff9800' : '#f44336', margin: '6px 0', fontWeight: 'bold' }}>
              {formatPercent(latestDay.profitRate)}
            </p>
          </div>
        </div>
      )}

      {/* 图表视图切换 */}
      {filteredDailyData.length >= 2 ? (
        <div style={chartContainerStyle}>
          <ReactECharts option={getChartOption()} style={{ height: isMobile ? '380px' : '500px' }} />
        </div>
      ) : filteredDailyData.length === 1 ? (
        <div style={hintBoxStyle}>
          <p>💡 <strong>当前筛选区间只有 1 天数据。</strong> 当前权益资金：{formatCurrency(filteredDailyData[0].balance)}。收益金额和收益率需要至少 <strong>2 天</strong> 的历史账单数据才能计算；你可以继续上传更多账单，或切回“全部”查看完整曲线。</p>
        </div>
      ) : (
        <p style={{ color: '#999', textAlign: 'center', marginTop: '60px' }}>
          {dailyData.length === 0 ? '暂无历史交易数据，请上传结算单文件。' : '当前筛选区间暂无数据，请切换到其他区间查看。'}
        </p>
      )}

      <div style={syncTipStyle}>
        <strong>同步说明：</strong> GitHub Pages 不能在网页里直接改仓库文件。PC 端上传完新账单后，请点“导出同步文件”，然后把下载得到的 `trade-data.json` 替换到项目的 `public/data/trade-data.json` 并提交发布。手机端下次打开时会自动合并这份共享数据。
      </div>
    </div>
  );
}

const uploadBoxStyle = { marginBottom: '24px', padding: '20px', border: '2px dashed #3f51b5', borderRadius: '8px', textAlign: 'center', background: '#f9f9ff' };
const cardStyle = { flex: 1, minWidth: 0, padding: '16px', background: '#f5f5f5', borderRadius: '8px', borderLeft: '5px solid #3f51b5' };
const labelStyle = { color: '#666', fontSize: '14px' };
const filterBarStyle = { display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '16px', marginBottom: '24px', padding: '14px 16px', background: '#f6f8fc', border: '1px solid #dde4f0', borderRadius: '10px', flexWrap: 'wrap' };
const filterBtnStyle = { padding: '8px 14px', borderRadius: '999px', border: '1px solid #c7d2e5', background: '#fff', color: '#44516b', cursor: 'pointer', fontWeight: 600 };
const activeFilterBtnStyle = { ...filterBtnStyle, background: '#3f51b5', color: '#fff', border: '1px solid #3f51b5', boxShadow: '0 6px 14px rgba(63,81,181,0.18)' };
const mobileFilterBarStyle = { ...filterBarStyle, alignItems: 'stretch', padding: '12px', gap: '12px' };
const mobileFilterBtnStyle = { ...filterBtnStyle, minWidth: '72px', textAlign: 'center' };
const mobileActiveFilterBtnStyle = { ...activeFilterBtnStyle, minWidth: '72px', textAlign: 'center' };
const chartContainerStyle = { background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 4px 12px rgba(0,0,0,0.08)' };
const hintBoxStyle = { padding: '20px', background: '#fff8e1', borderLeft: '5px solid #ffb300', borderRadius: '4px', color: '#b78103' };
const clearBtnStyle = { padding: '8px 14px', background: '#ffebee', color: '#c62828', border: '1px solid #ffcdd2', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' };
const secondaryBtnStyle = { padding: '8px 14px', background: '#eef3ff', color: '#2f4aac', border: '1px solid #cdd9ff', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' };
const mobileClearBtnStyle = { ...clearBtnStyle, width: '100%', justifyContent: 'center' };
const mobileSecondaryBtnStyle = { ...secondaryBtnStyle, width: '100%' };
const syncTipStyle = { marginTop: '24px', padding: '14px 16px', borderRadius: '10px', background: '#f8f9fb', border: '1px solid #e1e6ef', color: '#5b667a', fontSize: '13px', lineHeight: 1.6 };
