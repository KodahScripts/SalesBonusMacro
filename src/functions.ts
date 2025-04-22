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