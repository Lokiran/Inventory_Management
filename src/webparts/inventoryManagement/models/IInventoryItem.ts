export interface IInventoryItem {
  id: string;
  title: string;
  assetName: string;
  assetType: string;
  serialNumber: string;
  purchaseDate: string;
  vendor?: string;
  condition?: string;
  status: string;
  assignedTo?: string;
  assignedToEmail?: string;
  assignedDate?: string;
  warrantyExpiry?: string;
  specifications?: string;
  note?: string;
}
