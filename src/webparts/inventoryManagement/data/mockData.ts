import { IEmployee } from '../models/IEmployee';
import { IInventoryItem } from '../models/IInventoryItem';

export const EMPLOYEES: IEmployee[] = [
  { id: 'L005', name: 'Loka Kiran Reddy', email: 'Kiran.Reddy@3bh3kf.onmicrosoft.com', department: 'IT', jobTitle: 'Admin' },
  { id: 'E2', name: 'Adele Vance', email: 'AdeleV@3bh3kf.onmicrosoft.com', department: 'Operations', jobTitle: 'Inventory Employee' },
  { id: 'A001', name: 'Alex Wilber', email: 'AlexW@3bh3kf.onmicrosoft.com', department: 'Operations', jobTitle: 'Inventory Employee' },
  { id: 'D004', name: 'Diego Siciliani', email: 'DiegoS@3bh3kf.onmicrosoft.com', department: 'Management', jobTitle: 'Inventory Manager' },
  { id: 'G005', name: 'Grady Archie', email: 'GradyA@3bh3kf.onmicrosoft.com', department: 'Management', jobTitle: 'Inventory Manager' }
];


export const ASSET_CATEGORIES = [
  { key: 'IT', text: 'IT' },
  { key: 'Accessories', text: 'Accessories' },
  { key: 'Peripheral', text: 'Peripheral' },
  { key: 'Furniture', text: 'Furniture' },
  { key: 'Other', text: 'Other' }
];
