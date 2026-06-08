export interface IRequest {
  id: string;
  requestKey: string;
  requesterName: string;
  employeeId?: string;
  assetId: string;
  assetTitle: string;
  assetName?: string;
  priority?: 'High' | 'Medium' | 'Low' | string;
  quantity: number;
  status: 'Pending' | 'Approved' | 'Declined';
  assetStatus?: 'Pending' | 'Approved';
  managerResponse?: string;
  requestDate: string;
  reason?: string;
}
