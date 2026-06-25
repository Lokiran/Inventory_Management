import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IInventoryManagementProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  spContext: WebPartContext;
}
