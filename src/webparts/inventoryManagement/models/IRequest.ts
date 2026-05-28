export interface IRequest {
  id: string;
  requesterName: string;
  assetId: string;
  assetTitle: string;
  quantity: number;
  status: 'Pending' | 'Approved' | 'Declined';
  requestDate: string;
  reason?: string;
}
