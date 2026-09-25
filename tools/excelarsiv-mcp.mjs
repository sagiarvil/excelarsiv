#!/usr/bin/env node
/**
 * ExcelArşiv MCP (Model Context Protocol) Server v1.0.0
 * Stdio transport ile JSON-RPC 2.0 uyumlu Model Context Protocol sunucusu.
 * Claude Desktop, Cursor, OpenAI Agents ve bağımsız LLM karar ajanları için doğrudan entegrasyon sağlar.
 */

import { readFileSync, existsSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import readline from 'node:readline';
import {
  simulate13WeekCashflow,
  calculateBreakeven,
  calculateLoanScheduleAndDSCR
} from './excelarsiv-decision-kernel.mjs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const rootDir = resolve(__dirname, '..');

// Katalog verisini yükle
let catalog = [];
try {
  const catalogPath = resolve(rootDir, 'public', 'katalog.json');
  if (existsSync(catalogPath)) {
    const raw = readFileSync(catalogPath, 'utf8');
    const parsed = JSON.parse(raw);
    catalog = Array.isArray(parsed) ? parsed : (parsed.products || []);
  }
} catch (err) {
  process.stderr.write(`[ExcelArşiv MCP] Katalog okuma uyarısı: ${err.message}\n`);
}

// llms.txt içeriğini yükle
let llmsText = '';
try {
  const llmsPath = resolve(rootDir, 'public', 'llms.txt');
  if (existsSync(llmsPath)) {
    llmsText = readFileSync(llmsPath, 'utf8');
  }
} catch (err) {
  // Sessiz fallback
}

// Server Tool Tanımları
const TOOLS = [
  {
    name: 'search_templates',
    description: 'Search ready-to-use financial and operational Excel models by problem, keyword, or category.',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Search query or business problem (e.g. 13 haftalık nakit akışı, hakediş)'
        },
        category: {
          type: 'string',
          description: 'Category slug (e.g. nakit-akisi, muhasebe-ve-vergi)'
        },
        limit: {
          type: 'number',
          description: 'Maximum number of results to return (default: 5)'
        }
      },
      required: ['query']
    }
  },
  {
    name: 'get_template_details',
    description: 'Get detailed specifications, formula logic, input/output contracts, and demo links for an Excel template.',
    inputSchema: {
      type: 'object',
      properties: {
        slug: {
          type: 'string',
          description: 'Template unique ID slug (e.g. 13-haftalik-nakit-akisi-stres-testi-karar-paneli)'
        }
      },
      required: ['slug']
    }
  },
  {
    name: 'run_product_finder',
    description: 'Determine the optimal Excel system based on 5 business criteria.',
    inputSchema: {
      type: 'object',
      properties: {
        area: {
          type: 'string',
          description: 'Focus area (e.g. nakit-akisi, butce, vergi)'
        },
        businessType: {
          type: 'string',
          description: 'Business type (e.g. ticaret, uretim, eticaret)'
        },
        volume: {
          type: 'string',
          description: 'Record volume (az, orta, cok)'
        },
        period: {
          type: 'string',
          description: 'Time horizon (gun, hafta, ay, yil)'
        },
        coreProblem: {
          type: 'string',
          description: 'Primary bottleneck'
        }
      },
      required: ['area', 'coreProblem']
    }
  },
  {
    name: 'simulate_cashflow',
    description: 'Run deterministic 13-week cash flow stress testing simulation directly with ExcelArşiv Decision Kernel.',
    inputSchema: {
      type: 'object',
      properties: {
        startingBalance: {
          type: 'number',
          description: 'Opening cash buffer in TRY'
        },
        weeklyInflows: {
          type: 'array',
          items: { type: 'number' },
          description: 'Array of 13 numbers representing weekly cash inflows'
        },
        weeklyOutflows: {
          type: 'array',
          items: { type: 'number' },
          description: 'Array of 13 numbers representing weekly cash outflows'
        },
        stressInflowDropPct: {
          type: 'number',
          description: 'Percentage drop in inflows under stress (e.g. 15 for 15%)'
        },
        stressOutflowSurgePct: {
          type: 'number',
          description: 'Percentage surge in outflows under stress (e.g. 10 for 10%)'
        },
        minSafeCashBuffer: {
          type: 'number',
          description: 'Minimum required safety cash buffer'
        }
      },
      required: ['startingBalance', 'weeklyInflows', 'weeklyOutflows']
    }
  },
  {
    name: 'calculate_breakeven',
    description: 'Calculate breakeven units, revenue thresholds, and margin of safety.',
    inputSchema: {
      type: 'object',
      properties: {
        fixedCosts: {
          type: 'number',
          description: 'Monthly fixed costs in TRY'
        },
        unitPrice: {
          type: 'number',
          description: 'Average selling price per unit'
        },
        variableCostPerUnit: {
          type: 'number',
          description: 'Variable cost per unit'
        },
        targetProfit: {
          type: 'number',
          description: 'Target monthly net profit'
        }
      },
      required: ['fixedCosts', 'unitPrice', 'variableCostPerUnit']
    }
  },
  {
    name: 'calculate_loan_dscr',
    description: 'Calculate commercial loan installment schedule and Debt Service Coverage Ratio (DSCR).',
    inputSchema: {
      type: 'object',
      properties: {
        principal: {
          type: 'number',
          description: 'Principal loan amount in TRY'
        },
        monthlyInterestRatePct: {
          type: 'number',
          description: 'Monthly contractual interest rate (e.g. 3.5 for 3.5%)'
        },
        termMonths: {
          type: 'number',
          description: 'Loan term in months'
        },
        monthlyEbitda: {
          type: 'number',
          description: 'Monthly EBITDA to evaluate debt service covenants'
        }
      },
      required: ['principal', 'monthlyInterestRatePct', 'termMonths']
    }
  }
];

// Server Kaynak Tanımları
const RESOURCES = [
  {
    uri: 'excelarsiv://llms',
    name: 'LLMS Root Knowledge Graph',
    mimeType: 'text/markdown',
    description: 'Root manifest of Excel Arşiv knowledge architecture'
  },
  {
    uri: 'excelarsiv://templates',
    name: 'Excel Templates Full Catalog',
    mimeType: 'application/json',
    description: 'Complete list of corporate Excel decision templates'
  }
];

// Tool İcra Fonksiyonu
function handleToolCall(name, args = {}) {
  switch (name) {
    case 'search_templates': {
      const q = (args.query || '').toLowerCase().trim();
      const cat = (args.category || '').toLowerCase().trim();
      const limit = Number(args.limit) || 5;

      const results = catalog.filter((item) => {
        const titleMatch = (item.title || item.name || '').toLowerCase().includes(q);
        const descMatch = (item.description || '').toLowerCase().includes(q);
        const slugMatch = (item.slug || '').toLowerCase().includes(q);
        const textMatch = !q || titleMatch || descMatch || slugMatch;

        const catMatch = !cat || (item.category || item.categorySlug || '').toLowerCase() === cat;
        return textMatch && catMatch;
      }).slice(0, limit);

      return {
        count: results.length,
        query: args.query,
        category: args.category || null,
        results: results.map((r) => ({
          slug: r.slug,
          name: r.title || r.name,
          category: r.category || r.categorySlug,
          summary: r.description || r.summary,
          url: `https://excelarsiv.com/sablonlar/${r.slug}/`
        }))
      };
    }

    case 'get_template_details': {
      const slug = (args.slug || '').trim();
      const found = catalog.find((c) => c.slug === slug);
      if (!found) {
        return {
          error: `Şablon bulunamadı: '${slug}'. Mevcut şablonları aramak için search_templates aracını kullanın.`
        };
      }
      return {
        template: {
          slug: found.slug,
          name: found.title || found.name,
          category: found.category || found.categorySlug,
          price: found.price || found.fiyat || 'Standart',
          description: found.description || found.summary,
          features: found.features || found.ozellikler || [],
          url: `https://excelarsiv.com/sablonlar/${found.slug}/`,
          decisionOsCompatible: true,
          standard: 'ExcelArşiv Decision OS v15.1'
        }
      };
    }

    case 'run_product_finder': {
      const area = (args.area || '').toLowerCase();
      const coreProblem = (args.coreProblem || '').toLowerCase();

      // Skor bazlı eşleştirme
      const scored = catalog.map((item) => {
        let score = 0;
        const text = `${item.title} ${item.description} ${item.slug} ${item.category}`.toLowerCase();
        if (text.includes(area)) score += 3;
        if (text.includes(coreProblem)) score += 2;
        if (args.businessType && text.includes(args.businessType.toLowerCase())) score += 1;
        return { item, score };
      }).sort((a, b) => b.score - a.score);

      const top = scored.slice(0, 3).map((s) => s.item);
      return {
        matchedArea: args.area,
        topRecommendations: top.map((t) => ({
          slug: t.slug,
          name: t.title || t.name,
          reason: `Seçilen ${args.area} alanı ve ${args.coreProblem} darboğazı için Decision OS v15.1 mimarisinde önerilen model.`
        }))
      };
    }

    case 'simulate_cashflow': {
      return simulate13WeekCashflow(args);
    }

    case 'calculate_breakeven': {
      return calculateBreakeven(args);
    }

    case 'calculate_loan_dscr': {
      return calculateLoanScheduleAndDSCR(args);
    }

    default:
      throw new Error(`Bilinmeyen araç: ${name}`);
  }
}

// JSON-RPC 2.0 İletişim Döngüsü (stdio)
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
  terminal: false
});

function sendResponse(id, result, error = null) {
  const response = {
    jsonrpc: '2.0',
    id
  };
  if (error) {
    response.error = error;
  } else {
    response.result = result;
  }
  process.stdout.write(JSON.stringify(response) + '\n');
}

rl.on('line', (line) => {
  if (!line || !line.trim()) return;

  let message;
  try {
    message = JSON.parse(line.trim());
  } catch (err) {
    sendResponse(null, null, {
      code: -32700,
      message: 'Parse error: Geçersiz JSON verisi'
    });
    return;
  }

  const { id, method, params } = message;

  try {
    switch (method) {
      case 'initialize': {
        sendResponse(id, {
          protocolVersion: '2024-11-05',
          capabilities: {
            tools: { listChanged: false },
            resources: { subscribe: false, listChanged: false }
          },
          serverInfo: {
            name: 'excelarsiv-mcp',
            version: '1.0.0'
          }
        });
        break;
      }

      case 'notifications/initialized': {
        // MCP bildirimleri id içermez, yanıt gerektirmez
        break;
      }

      case 'ping': {
        sendResponse(id, {});
        break;
      }

      case 'tools/list': {
        sendResponse(id, {
          tools: TOOLS
        });
        break;
      }

      case 'tools/call': {
        const { name, arguments: args } = params || {};
        try {
          const toolResult = handleToolCall(name, args);
          sendResponse(id, {
            content: [
              {
                type: 'text',
                text: JSON.stringify(toolResult, null, 2)
              }
            ],
            isError: false
          });
        } catch (callErr) {
          sendResponse(id, {
            content: [
              {
                type: 'text',
                text: JSON.stringify({ error: callErr.message })
              }
            ],
            isError: true
          });
        }
        break;
      }

      case 'resources/list': {
        sendResponse(id, {
          resources: RESOURCES
        });
        break;
      }

      case 'resources/read': {
        const { uri } = params || {};
        if (uri === 'excelarsiv://llms') {
          sendResponse(id, {
            contents: [
              {
                uri,
                mimeType: 'text/markdown',
                text: llmsText || '# Excel Arşiv Knowledge Base'
              }
            ]
          });
        } else if (uri === 'excelarsiv://templates') {
          sendResponse(id, {
            contents: [
              {
                uri,
                mimeType: 'application/json',
                text: JSON.stringify(catalog, null, 2)
              }
            ]
          });
        } else {
          sendResponse(id, null, {
            code: -32602,
            message: `Kaynak bulunamadı: ${uri}`
          });
        }
        break;
      }

      default: {
        if (id !== undefined && id !== null) {
          sendResponse(id, null, {
            code: -32601,
            message: `Desteklenmeyen metod: ${method}`
          });
        }
        break;
      }
    }
  } catch (err) {
    if (id !== undefined && id !== null) {
      sendResponse(id, null, {
        code: -32603,
        message: `İç sunucu hatası: ${err.message}`
      });
    }
  }
});
