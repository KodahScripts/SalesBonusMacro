function calculateRetroBonus(input: number): number {
  if (input >= 16) return 7;
  if (input >= 12 && input < 16) return 4;
  return 0;
}

function calculateUnitBonus(input: number): number {
  if (input >= 10 && input < 12) return 375;
  if (input >= 12 && input < 16) return 750;
  if (input >= 16 && input < 20) return 1500;
  if (input >= 20 && input < 24) return 2500;
  if (input >= 24) return 3000;
  return 0;
}

function calculateRollingMini(input: number): number {
  if (input >= 12 && input < 16) return 250;
  if (input >= 16 && input < 20) return 300;
  if (input >= 20 && input < 24) return 350;
  if (input >= 24) return 400;
  return 200;
}

function calculate90DayUnitBonus(input: number): number {
  if (input >= 12 && input < 16) return 250;
  if (input >= 16 && input < 20) return 300;
  if (input >= 20 && input < 24) return 350;
  if (input >= 24) return 400;
  return 200;
}

function calculateTotalEOMBonus8328(retroCommission: number, fniCommission: number, monthEndBonus3122: number, spiffsToPay: number): number {
  return retroCommission + fniCommission + monthEndBonus3122 + spiffsToPay;
}

function calculateTotalCommissions(commission3120: number, totalEOMBonus8328: number): number {
  return commission3120 + totalEOMBonus8328;
}

function calculateYTDBucket(totalCommissions: number, draw3121: number, spiffsToPay: number): number {
  if (totalCommissions - draw3121 > 0) return 0;
  return totalCommissions - draw3121 - spiffsToPay;
}

function calculateDepositGross(totalCommissions: number, draw3121: number, ytdBucket: number): number {
  return totalCommissions - draw3121 - ytdBucket;
}

function calculateDrawToTake(commission3120: number, monthEndBonus3122: number, draw3121: number): number {
  if (commission3120 + monthEndBonus3122 >= draw3121) return draw3121;
  return commission3120 + monthEndBonus3122;
}

function getNPSScore(individualMonthlyScore: number, individual90DayScore: number): number {
  if (individualMonthlyScore > individual90DayScore) return individualMonthlyScore;
  return individual90DayScore;
}

function calculateCSIOutcome(npsScore: number, regionalScore: number): string {
  const plusThreePercent = regionalScore + (regionalScore * 0.03);
  if (npsScore > plusThreePercent) return "3P";
  if (npsScore == regionalScore) return "A";
  return "B";
}

function calculateNPSPercentage(promoterValue: number, passiveValue: number, detractorValue: number): number {
  return ((promoterValue - detractorValue) / (promoterValue + passiveValue + detractorValue));
}

function calculateJVCommissionBonus(commission3120: number, retroCommission: number, fniCommission: number, monthEndBonus3122: number, spiffsToPay: number): number {
  return commission3120 + retroCommission + fniCommission + monthEndBonus3122 + spiffsToPay;
}

function calculateJVTotalDue(draw3121: number, jvCommissionBonus: number): number {
  return jvCommissionBonus - draw3121;
}

function calculateUnitTotals(newUnitCount: number, usedUnitCount: number): number {
  return newUnitCount + usedUnitCount;
}

function getPercentage(value: number, totalAmount: number): number {
  return value / totalAmount;
}

function calculateExpense(jvCommissionBonus: number, commission3120: number, percentage: number): number {
  return ((jvCommissionBonus - commission3120) * percentage);
}