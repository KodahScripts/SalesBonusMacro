class NPS {
    public outcome: string;
    constructor(public currentScore: number, public averageScore: number, public regionalScore: number) {
        this.outcome = this.getOutcome();
    }
    getOutcome() {
        const score = this.currentScore > this.averageScore ? this.currentScore : this.averageScore;
        const weightedScore = (this.regionalScore * 0.03) + this.regionalScore;
        if(score > weightedScore) return '3P';
        if(score === this.regionalScore) return 'A';
        return 'B';
    }
    calculateBonus(unitCount: number): number {
        if(this.outcome === '3P') return unitCount * 50;
        if(this.outcome === 'B') return unitCount * -50;
        return 0;
    }
}


class Retro {
    constructor() {}
    getPercentage(lookupArray: LookupRow[], totalUnits: number) {
        lookupArray.map((row, index) => {
            const currentVal = row[0];
            const nextVal = lookupArray[index + 1][0];
            if(totalUnits >= currentVal && totalUnits < nextVal) {
                return row[1];
            }
        });
        return 0;
    }
    mini(unitAverage: number, dealUnitCount: number) {
        if (unitAverage >= 24) return 400 * dealUnitCount;
        if (unitAverage >= 20 && unitAverage < 24) return 350 * dealUnitCount;
        if (unitAverage >= 16 && unitAverage < 20) return 300 * dealUnitCount;
        if (unitAverage >= 12 && unitAverage < 16) return 250 * dealUnitCount;
        if (unitAverage >= 1 && unitAverage < 12) return 200 * dealUnitCount;
        return 0;
    }
    owed(commissionAmount: number) {
        if(this.mini() > 0) return 
    }
}


interface LookupRow {
    min: number;
    max: number;
    val: number;
}

class LookupReport {
    public rows: LookupRow[];
    constructor(private data: Array<string|number|boolean>[]) {
        this.rows = [];
        data.map((d, i) => {
            const next = data[i+1];
            const max = next ? Number(data[i+1][0]) : 0;
            this.rows.push({ min: Number(d[0]), max, val: Number(d[1]) });
        });
    }
    getValue(query: number): number {
        this.rows.map(r => {
            if(r.max === 0 && query >= r.min || query >= r.min && query < r.max) {
                return r.val;
            }
        });
        return 0;
    }
}