export interface IReturnRequest {
  id: string;
  title: string;
  assetId: string;
  assetName: string;
  serialNumber: string;
  requesterName: string;
  requesterEmail?: string;
  requestDate: string;
  returnReason: string;
  proposedCondition: string; // Good, Fair, Poor, Damaged
  status: 'Pending' | 'Approved' | 'Rejected' | 'Completed';
  managerComment?: string;
  completedDate?: string;
}
