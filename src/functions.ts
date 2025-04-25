function calculateRetroMini(commissionAmount: number, unitAvg: number, dealUnitCount: number): number {
    if (commissionAmount <= 251) {
        if (unitAvg >= 24) return 400 * dealUnitCount;
        if (unitAvg >= 20 && unitAvg < 24) return 350 * dealUnitCount;
        if (unitAvg >= 16 && unitAvg < 20) return 300 * dealUnitCount;
        if (unitAvg >= 12 && unitAvg < 16) return 250 * dealUnitCount;
        if (unitAvg >= 1 && unitAvg < 12) return 200 * dealUnitCount;
        return 0;
    }
    return 0;
}

function calculateCsiOutcome(currentCsiScore: number, averageCsiScore: number, regionalScore: number): string {
    const score = currentCsiScore > averageCsiScore ? currentCsiScore : averageCsiScore;
    if (score > regionalScore + (regionalScore * 0.03)) return "3P"
    if (score == regionalScore) return "A";
    return "B";
}

function calculateUnitBonus(unitCount: number): number {
    if (unitCount >= 24) return 3000;
    if (unitCount >= 20 && unitCount < 24) return 2500;
    if (unitCount >= 16 && unitCount < 20) return 1500;
    if (unitCount >= 12 && unitCount < 16) return 750;
    if (unitCount >= 10 && unitCount < 12) return 375;
    return 0;
}

function caclulateCsiBonus(surveyCount: number, csiOutcome: string, unitCount: number) {
    if (surveyCount >= 3) {
        if (csiOutcome === "3P") return unitCount * 50;
        if (csiOutcome === "A") return unitCount * 0;
        if (csiOutcome === "B") return unitCount * -50;
    }
    return 0;
}

function calculateYtdBucket(totalCommission: number, priorDraw: number, spiffs: number): number {
    if (totalCommission - priorDraw < 0) {
        return totalCommission - priorDraw - spiffs;
    }
    return 0;
}