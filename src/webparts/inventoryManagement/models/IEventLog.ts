export interface IEventLog {
  id: string;
  title: string;
  action: 'Create' | 'Update' | 'Delete';
  entityType: 'Asset' | 'Request';
  entityId: string;
  assetName?: string;
  details: string;
  user: string;
  timestamp: string;
}
