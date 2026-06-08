export interface INotification {
  id: string;
  title: string;
  message: string;
  type: 'info' | 'success' | 'warning' | 'error';
  timestamp: string;
  isRead: boolean;
  actionLink: string; // The Pivot tab key to switch to (e.g. 'MyRequests', 'Approvals')
  category: 'Request' | 'Assignment' | 'Audit';
}
