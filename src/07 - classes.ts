class Store {
  public name: string;
  public abbr: string;
  public employees: Array<Employee>;
  constructor(public lookupData: Array<string | number | boolean>[]) {
    this.name = STORE_NAME;
    this.abbr = STORE_ABBR;
    this.lookupData = lookupData;
    this.employees = [];
  }

  createEmployees(data_0432: Array<string | number | boolean>[]) {
    data_0432.forEach((row, index) => {
      const [date, deal_id, emp_id, emp_name, cust_id, cust_name, prefix, veh_id, veh_desc, sale_type, comm_fni, comm_front, unit_count, comm_amount] = row

      if (index > 0) {
        let employee = this.employees.filter(emp => emp.id == emp_id)[0];

        if (!employee) {
          employee = this.addEmployee(Number(emp_id), String(emp_name).toLocaleUpperCase());
        }
        if (Number(unit_count) > 0) {
          employee.addDeal(String(deal_id), Number(date), Number(cust_id), String(cust_name), String(veh_id), String(veh_desc), String(sale_type), Number(unit_count), Number(comm_fni), Number(comm_front), Number(comm_amount));
        }
      }
    })
  }

  addEmployee(id: number, name: string) {
    const employee = new Employee(id, name);
    this.employees.push(employee);
    return employee;
  }
}

abstract class Person {
  constructor(public id: number, public name: string) {
    this.id = id;
    this.name = name;
  }
}

class Customer extends Person {
  constructor(public id: number, public name: string) {
    super(id, name);
  }
}

class Employee extends Person {
  public deals: Array<Deal>;
  constructor(public id: number, public name: string) {
    super(id, name);

    this.deals = [];
  }

  addDeal(id: string, date: number, custId: number, custName: string, vehId: string, vehDesc: string, salesType: string, unitCount: number, commGross: number, grossPercent: number, commAmount: number) {
    const customer = new Customer(custId, custName);
    const vehicle = new Vehicle(vehId, vehDesc, salesType);
    const commission = new Commission(commGross, grossPercent, commAmount);
    this.deals.push(new Deal(id, date, customer, vehicle, unitCount, commission));
  }

  getReportResultRow() {
    return this.deals.length + 8;
  }

  getResultRow(rowNumber: number) {
    return this.getReportResultRow() + (2 * rowNumber);
  }

  getTotalUnits(filter?: string): number {
    let units = { new: 0, used: 0, total: 0 };
    this.deals.forEach(deal => {
      const t = deal.vehicle.salesType.toLowerCase();
      const count = deal.unitCount;
      units[t] += count;
      units.total += count;
    });
    return filter ? Number(units[filter]) : Number(units.total);
  }

  getTotalCommission(filter: string): number {
    let comms = { gross: 0, grossPercentage: 0, grossAmount: 0 };
    this.deals.forEach(deal => {
      comms.gross += deal.commission.gross;
      comms.grossPercentage += deal.commission.grossPercentage;
      comms.grossAmount += deal.commission.amount;
    });
    return comms[filter];
  }
}

class Commission {
  constructor(public gross: number, public grossPercentage: number, public amount: number) {
    this.gross = gross;
    this.grossPercentage = grossPercentage;
    this.amount = amount;
  }
}

class Vehicle {
  public year: string;
  public make: string;
  public model: string;
  public description: string;
  constructor(public id: string, public vehicleDescription: string, public salesType: string) {
    this.id = id;
    [this.year, this.make, this.model, this.description] = vehicleDescription.split(',');
    this.salesType = salesType;
  }
}

class Deal {
  constructor(public id: string, public date: number, public customer: Customer, public vehicle: Vehicle, public unitCount: number, public commission: Commission) {
    this.id = id;
    this.date = date;
    this.customer = customer;
    this.vehicle = vehicle;
    this.unitCount = unitCount;
    this.commission = commission;
  }

  calculateRetroMini() {
    if(this.commission.amount > 251) return 0;
    
  }
}