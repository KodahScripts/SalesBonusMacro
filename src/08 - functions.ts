function calculateRetroMini(commissionAmount: number, unitAvg: number, dealUnitCount: number): number {
    if (commissionAmount <= 251) {
        if (unitAvg >= 24) return 400 * dealUnitCount;
        if (unitAvg >= 20 && unitAvg < 24) return 350 * dealUnitCount;
        if (unitAvg >= 16 && unitAvg < 20) return 300 * dealUnitCount;
        if (unitAvg >= 12 && unitAvg < 16) return 250 * dealUnitCount;
        if (unitAvg >= 1 && unitAvg < 12) return 200 * dealUnitCount;
        return 0;
    } else {
        return 0;
    }
}

function calculateUnitBonus(totalUnits: number) {
        if (totalUnits >= 24) return 3000;
        if (totalUnits >= 20 && totalUnits < 24) return 2500;
        if (totalUnits >= 16 && totalUnits < 20) return 1500;
        if (totalUnits >= 12 && totalUnits < 16) return 750;
        if (totalUnits >= 10 && totalUnits < 12) return 375;
        return 0;
}

function calculateRetroOwed(retroMini: number, commissionAmount: number) : number {
    return retroMini > 0 ? retroMini - commissionAmount : 0;
}

function getRetroPercentage(unitCount: number): number {
    if (unitCount >= 16) return 0.07;
    if (unitCount >= 12 && unitCount < 16) return 0.04;
    return 0;
}

function calculateRetroPayout(commissionFront: number, retroPercentage: number): number {
    return commissionFront * retroPercentage;
}

function calculateRetroTotal(retroPayout: number, retroOwed: number): number {
    return retroPayout + retroOwed;
}

function calculateFniReserve(totalCommissionFni: number) {
    return totalCommissionFni * 0.25;
}

function calculateFniGross(totalCommissionFni: number, fniReserve: number) {
    return totalCommissionFni - fniReserve;
}

function calculateFniPayout(fniGross: number) {
    return fniGross * 0.05;
}

function calculateTotalBonus(unitBonus: number, csiBonus: number, topsalesBonus: number) {
    return unitBonus + csiBonus + topsalesBonus;
}

function calculateCsiBonus(surveyCount: number) {
    if(surveyCount >= 3) {}
}

function chooseNpsScore(currentScore: number, averageScore: number) {
    return currentScore > averageScore ? currentScore : averageScore;
}

function getCsiOutcome(npsScore: number, regionalScore: number) {
    if(npsScore > regionalScore + 0.03) return "3P";
    if(npsScore == regionalScore) return "A";
    if(npsScore < regionalScore) return "B";
}