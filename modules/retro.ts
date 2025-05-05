class Retro {
    public mini: number;
    public owed: number;
    public payout: number;
    public total: number;
    constructor(private commissionAmount: number, private commissionGross: number, private averageUnitCount: number, private dealUnitCount: number, private totalUnitCount: number) {
        this.mini = this.calculateMini(commissionAmount, averageUnitCount, dealUnitCount);
        this.owed = this.mini > 0 ? this.mini - commissionAmount : 0;
        this.payout = this.mini === 0 ? commissionGross * this.getPercentage(totalUnitCount) : 0;
        this.total = this.payout + this.owed;
    }

    calculateMini(commissionAmount, averageUnits, dealUnitCount): number {
        if (commissionAmount <= 251) {
            if (averageUnits >= 24) return 400 * dealUnitCount;
            if (averageUnits >= 20 && averageUnits < 24) return 350 * dealUnitCount;
            if (averageUnits >= 16 && averageUnits < 20) return 300 * dealUnitCount;
            if (averageUnits >= 12 && averageUnits < 16) return 250 * dealUnitCount;
            if (averageUnits >= 1 && averageUnits < 12) return 200 * dealUnitCount;
            return 0;
        }
        return 0;
    }

    getPercentage(totalUnitCount: number): number {
        if(totalUnitCount >= 16) return 0.07;
        if(totalUnitCount >= 12 && totalUnitCount < 16) return 0.04;
        return 0;
    }
}