/**
 * ExcelArşiv Decision Kernel v1.0.0
 * Deterministik Finansal Karar ve Simülasyon Motoru
 * (13 Haftalık Nakit Akışı, Kredi Taksit/Covenant, Başabaş Noktası)
 */

/**
 * 13 Haftalık Nakit Akışı Stres Testi
 * @param {Object} params
 * @param {number} params.startingBalance - Başlangıç nakit rezervi (TRY)
 * @param {Array<number>} params.weeklyInflows - 13 haftalık nakit girişleri (13 eleman)
 * @param {Array<number>} params.weeklyOutflows - 13 haftalık nakit çıkışları (13 eleman)
 * @param {number} [params.stressInflowDropPct=0] - Girişlerde stres düşüş yüzdesi (örn: 15 için %15 düşüş)
 * @param {number} [params.stressOutflowSurgePct=0] - Çıkışlarda stres artış yüzdesi (örn: 10 için %10 artış)
 * @param {number} [params.minSafeCashBuffer=50000] - Güvenli minimum nakit tamponu
 */
export function simulate13WeekCashflow(params) {
  const {
    startingBalance = 0,
    weeklyInflows = [],
    weeklyOutflows = [],
    stressInflowDropPct = 0,
    stressOutflowSurgePct = 0,
    minSafeCashBuffer = 50000
  } = params;

  if (weeklyInflows.length !== 13 || weeklyOutflows.length !== 13) {
    throw new Error('13 haftalık nakit akışı için tam olarak 13 haftalık giriş ve çıkış verisi gereklidir.');
  }

  const inflowMult = 1 - (stressInflowDropPct / 100);
  const outflowMult = 1 + (stressOutflowSurgePct / 100);

  let currentBalance = startingBalance;
  let minBalanceObserved = startingBalance;
  let minBalanceWeek = 0;
  let totalInflows = 0;
  let totalOutflows = 0;
  const weeklyTrajectory = [];
  const deficitWeeks = [];

  for (let w = 0; w < 13; w++) {
    const adjustedInflow = Math.round(weeklyInflows[w] * inflowMult * 100) / 100;
    const adjustedOutflow = Math.round(weeklyOutflows[w] * outflowMult * 100) / 100;
    const netWeekly = Math.round((adjustedInflow - adjustedOutflow) * 100) / 100;

    currentBalance = Math.round((currentBalance + netWeekly) * 100) / 100;
    totalInflows += adjustedInflow;
    totalOutflows += adjustedOutflow;

    if (currentBalance < minBalanceObserved) {
      minBalanceObserved = currentBalance;
      minBalanceWeek = w + 1;
    }

    if (currentBalance < minSafeCashBuffer) {
      deficitWeeks.push({
        week: w + 1,
        balance: currentBalance,
        deficitAgainstBuffer: Math.round((minSafeCashBuffer - currentBalance) * 100) / 100
      });
    }

    weeklyTrajectory.push({
      week: w + 1,
      inflow: adjustedInflow,
      outflow: adjustedOutflow,
      net: netWeekly,
      endingBalance: currentBalance,
      isBelowBuffer: currentBalance < minSafeCashBuffer
    });
  }

  const isSolvent = minBalanceObserved >= 0;
  const isBufferCompliant = deficitWeeks.length === 0;

  return {
    success: true,
    startingBalance,
    endingBalance: currentBalance,
    totalNetChange: Math.round((currentBalance - startingBalance) * 100) / 100,
    totalInflows: Math.round(totalInflows * 100) / 100,
    totalOutflows: Math.round(totalOutflows * 100) / 100,
    minBalanceObserved,
    minBalanceWeek,
    isSolvent,
    isBufferCompliant,
    deficitWeeksCount: deficitWeeks.length,
    deficitWeeks,
    trajectory: weeklyTrajectory,
    verdict: !isSolvent
      ? `KRİTİK UYARI: ${minBalanceWeek}. haftada nakit negatif bölgeye (${minBalanceObserved} TRY) düşüyor. Acil likidite köprüsü gerekir.`
      : !isBufferCompliant
        ? `DİKKAT: Nakit pozitif fakat ${deficitWeeks.length} hafta boyunca emniyet tamponunun (${minSafeCashBuffer} TRY) altında.`
        : `MÜKEMMEL: 13 hafta boyunca nakit tamponu korunmakta.`
  };
}

/**
 * Başabaş Noktası & Finansal Eşik Hesabı (Breakeven Analysis)
 * @param {Object} params
 * @param {number} params.fixedCosts - Aylık toplam sabit maliyetler (Kira, bordro, SGK vb.)
 * @param {number} params.unitPrice - Ortalama birim satış fiyatı
 * @param {number} params.variableCostPerUnit - Birim başına değişken maliyet (Hammadde, komisyon, sevkiyat vb.)
 * @param {number} [params.targetProfit=0] - Hedeflenen aylık net kâr
 */
export function calculateBreakeven(params) {
  const {
    fixedCosts = 0,
    unitPrice = 0,
    variableCostPerUnit = 0,
    targetProfit = 0
  } = params;

  if (unitPrice <= variableCostPerUnit) {
    throw new Error('Birim satış fiyatı birim değişken maliyetten büyük olmalıdır (Katkı payı pozitif olmalıdır).');
  }

  const unitContributionMargin = Math.round((unitPrice - variableCostPerUnit) * 100) / 100;
  const contributionMarginRatio = Math.round((unitContributionMargin / unitPrice) * 10000) / 10000;

  // Başabaş Adedi: Sabit Maliyet / Birim Katkı Payı
  const breakevenUnits = Math.ceil(fixedCosts / unitContributionMargin);
  // Başabaş Cirosu: Başabaş Adedi * Satış Fiyatı
  const breakevenRevenue = Math.round(breakevenUnits * unitPrice * 100) / 100;

  // Hedef Kâr için Gerekli Adet ve Ciro
  const targetUnits = Math.ceil((fixedCosts + targetProfit) / unitContributionMargin);
  const targetRevenue = Math.round(targetUnits * unitPrice * 100) / 100;

  // Güvenlik Marjı Hesabı (Eğer mevcut satış adedi verilirse veya hedef adede göre)
  const marginOfSafetyPct = targetUnits > 0
    ? Math.round(((targetUnits - breakevenUnits) / targetUnits) * 10000) / 100
    : 0;

  return {
    success: true,
    fixedCosts,
    unitPrice,
    variableCostPerUnit,
    unitContributionMargin,
    contributionMarginRatioPct: Math.round(contributionMarginRatio * 10000) / 100,
    breakevenUnits,
    breakevenRevenue,
    targetProfit,
    targetUnits,
    targetRevenue,
    marginOfSafetyPct,
    verdict: `Aylık sabit giderleri karşılamak için en az ${breakevenUnits} adet (${breakevenRevenue.toLocaleString('tr-TR')} TRY ciro) satış yapılmalıdır.`
  };
}

/**
 * Ticari Kredi Taksit & Borç Servis Karşılama Oranı (DSCR) Analizi
 * @param {Object} params
 * @param {number} params.principal - Kredi tutarı (Anapara)
 * @param {number} params.monthlyInterestRatePct - Aylık akdi faiz oranı (örn: 3.5 için %3.5)
 * @param {number} params.termMonths - Vade (ay)
 * @param {number} [params.kkdfPct=0] - KKDF oranı (Ticari kredilerde genelde 0)
 * @param {number} [params.bsmvPct=5] - BSMV oranı (%5)
 * @param {number} [params.monthlyEbitda=0] - Aylık FAVÖK / Faaliyet Kârı
 */
export function calculateLoanScheduleAndDSCR(params) {
  const {
    principal = 0,
    monthlyInterestRatePct = 0,
    termMonths = 0,
    kkdfPct = 0,
    bsmvPct = 5,
    monthlyEbitda = 0
  } = params;

  if (principal <= 0 || monthlyInterestRatePct <= 0 || termMonths <= 0) {
    throw new Error('Anapara, faiz oranı ve vade pozitif sayılar olmalıdır.');
  }

  // Vergi ve fon yüküyle efektif aylık faiz
  const taxFactor = 1 + (kkdfPct / 100) + (bsmvPct / 100);
  const grossMonthlyRate = (monthlyInterestRatePct / 100) * taxFactor;

  // Standart Eşit Taksit (Annuity) Formülü: P * [r(1+r)^n] / [(1+r)^n - 1]
  const compound = Math.pow(1 + grossMonthlyRate, termMonths);
  const monthlyPayment = Math.round((principal * (grossMonthlyRate * compound) / (compound - 1)) * 100) / 100;
  const totalRepayment = Math.round(monthlyPayment * termMonths * 100) / 100;
  const totalInterestAndTaxes = Math.round((totalRepayment - principal) * 100) / 100;

  // DSCR (Borç Servis Karşılama Oranı) = Aylık FAVÖK / Aylık Taksit
  let dscr = null;
  let dscrStatus = 'Bilinmiyor (FAVÖK girilmedi)';
  if (monthlyEbitda > 0) {
    dscr = Math.round((monthlyEbitda / monthlyPayment) * 100) / 100;
    if (dscr >= 1.35) {
      dscrStatus = 'GÜÇLÜ (DSCR >= 1.35) - Banka kovenantlarını rahat karşılıyor.';
    } else if (dscr >= 1.15) {
      dscrStatus = 'KABUL EDİLEBİLİR (1.15 <= DSCR < 1.35) - Standart limitlerde.';
    } else if (dscr >= 1.0) {
      dscrStatus = 'RİSKLİ (1.0 <= DSCR < 1.15) - Nakit tamponu çok dar.';
    } else {
      dscrStatus = 'TEMERRÜT RİSKİ (DSCR < 1.0) - Faaliyet kârı kredi taksitini karşılamıyor!';
    }
  }

  return {
    success: true,
    principal,
    monthlyInterestRatePct,
    termMonths,
    monthlyPayment,
    totalRepayment,
    totalInterestAndTaxes,
    monthlyEbitda,
    dscr,
    dscrStatus,
    verdict: `Aylık eşit taksit tutarı: ${monthlyPayment.toLocaleString('tr-TR')} TRY. Toplam geri ödeme: ${totalRepayment.toLocaleString('tr-TR')} TRY.`
  };
}
