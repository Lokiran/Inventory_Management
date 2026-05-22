export interface IRequest {
  id: string;
  requestKey: string;
  requesterName: string;
  assetId: string;
  assetTitle: string;
  quantity: number;
  status: 'Pending' | 'Approved' | 'Declined';
  assetStatus?: 'Pending' | 'Approved';
  managerResponse?: string;
  requestDate: string;
  reason?: string;
}
