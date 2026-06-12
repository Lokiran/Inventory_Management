export interface IEventLog {
  id: string;
  title: string;
  action: string;
  entityType: 'Asset' | 'Request';
  entityId: string;
  assetName?: string;
  details: string;
  user: string;
  timestamp: string;
}
