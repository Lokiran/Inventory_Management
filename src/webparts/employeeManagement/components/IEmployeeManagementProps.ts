import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEmployeeManagementProps {
  userEmail: string;
  userName: string;
  webUrl: string;
  spContext: WebPartContext;
}
