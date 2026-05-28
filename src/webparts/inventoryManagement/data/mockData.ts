import { IEmployee } from '../models/IEmployee';
import { IInventoryItem } from '../models/IInventoryItem';

export const EMPLOYEES: IEmployee[] = [
  { id: 'E1', name: 'John Doe', email: 'john.doe@example.com', department: 'IT', jobTitle: 'Software Engineer' },
  { id: 'E2', name: 'Jane Smith', email: 'jane.smith@example.com', department: 'HR', jobTitle: 'HR Manager' },
  { id: 'E3', name: 'Bob Johnson', email: 'bob.johnson@example.com', department: 'Finance', jobTitle: 'Accountant' },
  { id: 'E4', name: 'Alice Williams', email: 'alice.williams@example.com', department: 'Marketing', jobTitle: 'Marketing Specialist' },
  { id: 'E5', name: 'Charlie Brown', email: 'charlie.brown@example.com', department: 'IT', jobTitle: 'Network Admin' },
];


export const ASSET_CATEGORIES = [
  { key: 'IT', text: 'IT' },
  { key: 'Accessories', text: 'Accessories' },
  { key: 'Peripheral', text: 'Peripheral' },
  { key: 'Furniture', text: 'Furniture' },
  { key: 'Other', text: 'Other' }
];
