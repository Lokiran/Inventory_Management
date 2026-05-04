export interface IInventoryItem {
  id: string;
  title: string;
  assetName: string;
  assetType: string;
  serialNumber: string;
  purchaseDate: string;
  status: string;
  assignedTo?: string;
}
