/**
 * ExcelArşiv MCP Server & Decision Kernel Deterministic Test Suite
 * Iron Law of Verification PASS Checker
 */

import { spawn } from 'node:child_process';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const mcpServerPath = resolve(__dirname, 'excelarsiv-mcp.mjs');

let passCount = 0;
let failCount = 0;

function assert(condition, message) {
  if (condition) {
    passCount++;
    console.log(`  ✅ PASS: ${message}`);
  } else {
    failCount++;
    console.error(`  ❌ FAIL: ${message}`);
  }
}

async function runTest() {
  console.log('🚀 [ExcelArşiv MCP Test] Sunucu başlatılıyor...');

  const proc = spawn('node', [mcpServerPath], {
    stdio: ['pipe', 'pipe', 'inherit']
  });

  const responses = [];

  proc.stdout.on('data', (data) => {
    const lines = data.toString().split('\n');
    for (const line of lines) {
      if (!line.trim()) continue;
      try {
        responses.push(JSON.parse(line.trim()));
      } catch (err) {
        console.error('Yanıt parse hatası:', err, line);
      }
    }
  });

  function sendRequest(req) {
    proc.stdin.write(JSON.stringify(req) + '\n');
  }

  function waitForResponse(id, timeout = 4000) {
    const start = Date.now();
    return new Promise((resolve, reject) => {
      const interval = setInterval(() => {
        const found = responses.find((r) => r.id === id);
        if (found) {
          clearInterval(interval);
          resolve(found);
        } else if (Date.now() - start > timeout) {
          clearInterval(interval);
          reject(new Error(`Zaman aşımı: id ${id}`));
        }
      }, 50);
    });
  }

  try {
    // 1. Initialize
    console.log('\n--- Test 1: initialize ---');
    sendRequest({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: { protocolVersion: '2024-11-05' }
    });
    const initRes = await waitForResponse(1);
    assert(initRes.result.serverInfo.name === 'excelarsiv-mcp', 'Sunucu adı excelarsiv-mcp');
    assert(initRes.result.capabilities.tools !== undefined, 'Tools capability mevcut');

    // 2. Tools List
    console.log('\n--- Test 2: tools/list ---');
    sendRequest({
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/list',
      params: {}
    });
    const toolsRes = await waitForResponse(2);
    const tools = toolsRes.result.tools;
    assert(tools.length >= 6, `En az 6 MCP aracı tanımlı (bulunan: ${tools.length})`);
    const toolNames = tools.map((t) => t.name);
    assert(toolNames.includes('search_templates'), 'search_templates aracı mevcut');
    assert(toolNames.includes('simulate_cashflow'), 'simulate_cashflow aracı mevcut');
    assert(toolNames.includes('calculate_breakeven'), 'calculate_breakeven aracı mevcut');
    assert(toolNames.includes('calculate_loan_dscr'), 'calculate_loan_dscr aracı mevcut');

    // 3. Search Templates Tool Call
    console.log('\n--- Test 3: tools/call (search_templates) ---');
    sendRequest({
      jsonrpc: '2.0',
      id: 3,
      method: 'tools/call',
      params: {
        name: 'search_templates',
        arguments: { query: 'nakit' }
      }
    });
    const searchRes = await waitForResponse(3);
    const searchPayload = JSON.parse(searchRes.result.content[0].text);
    assert(searchPayload.count > 0, `Nakit araması sonuç döndürdü (${searchPayload.count} sonuç)`);

    // 4. Simulate Cashflow Tool Call (13-Week Stress Test)
    console.log('\n--- Test 4: tools/call (simulate_cashflow) ---');
    const weeklyInflows = [100000, 120000, 90000, 110000, 80000, 95000, 105000, 115000, 130000, 125000, 140000, 135000, 150000];
    const weeklyOutflows = [90000, 85000, 95000, 100000, 110000, 90000, 85000, 95000, 100000, 95000, 105000, 110000, 100000];
    sendRequest({
      jsonrpc: '2.0',
      id: 4,
      method: 'tools/call',
      params: {
        name: 'simulate_cashflow',
        arguments: {
          startingBalance: 200000,
          weeklyInflows,
          weeklyOutflows,
          stressInflowDropPct: 10,
          stressOutflowSurgePct: 5,
          minSafeCashBuffer: 100000
        }
      }
    });
    const cashRes = await waitForResponse(4);
    const cashPayload = JSON.parse(cashRes.result.content[0].text);
    assert(cashPayload.success === true, 'Nakit akışı simülasyonu başarıyla hesaplandı');
    assert(cashPayload.trajectory.length === 13, '13 haftalık yörünge üretildi');
    assert(typeof cashPayload.endingBalance === 'number', `Kapanış bakiyesi hesaplandı: ${cashPayload.endingBalance}`);

    // 5. Calculate Breakeven Tool Call
    console.log('\n--- Test 5: tools/call (calculate_breakeven) ---');
    sendRequest({
      jsonrpc: '2.0',
      id: 5,
      method: 'tools/call',
      params: {
        name: 'calculate_breakeven',
        arguments: {
          fixedCosts: 150000,
          unitPrice: 500,
          variableCostPerUnit: 200,
          targetProfit: 60000
        }
      }
    });
    const beRes = await waitForResponse(5);
    const bePayload = JSON.parse(beRes.result.content[0].text);
    assert(bePayload.breakevenUnits === 500, 'Başabaş adedi 500 olarak doğru hesaplandı (150000 / 300)');
    assert(bePayload.breakevenRevenue === 250000, 'Başabaş cirosu 250.000 TRY doğru hesaplandı');

    // 6. Resources List
    console.log('\n--- Test 6: resources/list ---');
    sendRequest({
      jsonrpc: '2.0',
      id: 6,
      method: 'resources/list',
      params: {}
    });
    const resListRes = await waitForResponse(6);
    assert(resListRes.result.resources.length >= 2, 'En az 2 MCP kaynağı mevcut');

  } finally {
    proc.stdin.end();
    proc.kill();
  }

  console.log(`\n========================================`);
  console.log(`📊 MCP & KARAR MOTORU TEST SONUCU:`);
  console.log(`   Toplam Geçen: ${passCount}`);
  console.log(`   Toplam Kalan: ${failCount}`);
  console.log(`========================================`);

  if (failCount > 0) {
    process.exit(1);
  }
}

runTest().catch((err) => {
  console.error('Kritik test hatası:', err);
  process.exit(1);
});
