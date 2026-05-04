import * as React from 'react';
import styles from './InventoryManagement.module.scss';
import type { IInventoryManagementProps } from './IInventoryManagementProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { InventoryList } from './InventoryList';
import { RequestList } from './RequestList';
import { IInventoryItem } from '../models/IInventoryItem';
import { IRequest } from '../models/IRequest';
import { IEventLog } from '../models/IEventLog';
import { AssetForm } from './AssetForm';
import { RequestForm } from './RequestForm';
import { EventStream } from './EventStream';
import { PrimaryButton, Pivot, PivotItem } from '@fluentui/react';

import { EMPLOYEES } from '../data/mockData';
import { InventoryService } from '../services/InventoryService';
import { Dashboard } from './Dashboard';

export interface IInventoryManagementState {
  items: IInventoryItem[];
  requests: IRequest[];
  auditLogs: IEventLog[];
  isAssetFormOpen: boolean;
  isRequestFormOpen: boolean;
  loading: boolean;
  auditLogsLoading: boolean;
  errorMessage?: string;
}

export default class InventoryManagement extends React.Component<IInventoryManagementProps, IInventoryManagementState> {
  constructor(props: IInventoryManagementProps) {
    super(props);
    this.state = {
      items: [],
      requests: [],
      auditLogs: [],
      isAssetFormOpen: false,
      isRequestFormOpen: false,
      loading: true,
      auditLogsLoading: true,
      errorMessage: undefined
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadInventory();
    await this._loadRequests();
    await this._loadAuditLogs();
  }

  private _loadInventory = async (): Promise<void> => {
    try {
      this.setState({ loading: true, errorMessage: undefined });
      const items = await InventoryService.getItems(this.props.sp);
      if (items && items.length > 0) {
        this.setState({ items, loading: false });
      } else {
        // List is empty
        this.setState({ 
          items: [], 
          loading: false,
          errorMessage: 'SharePoint list is empty. Please add items.'
        });
      }
    } catch (error: any) {
      console.error("Failed to load inventory:", error);
      
      // Fallback to empty if SharePoint fails so the UI remains functional
      this.setState({ 
        items: [],
        loading: false, 
        errorMessage: `SharePoint Error: ${error.message || JSON.stringify(error)}`
      });
    }
  };

  private _loadRequests = async (): Promise<void> => {
    try {
      const requests = await InventoryService.getRequests(this.props.sp);
      if (requests && requests.length > 0) {
        this.setState({ requests });
      }
    } catch (error) {
      console.error("Failed to load requests:", error);
    }
  };

  private _loadAuditLogs = async (): Promise<void> => {
    try {
      this.setState({ auditLogsLoading: true });
      const auditLogs = await InventoryService.getAuditLogs(this.props.sp);
      this.setState({ auditLogs, auditLogsLoading: false });
    } catch (error) {
      console.error("Failed to load audit logs:", error);
      this.setState({ auditLogsLoading: false });
    }
  };

  private _onAddAsset = async (newAssetData: Omit<IInventoryItem, 'id' | 'status' | 'assignedTo'>): Promise<void> => {
    try {
      const newAsset: Omit<IInventoryItem, 'id'> = {
        ...newAssetData,
        status: 'In Stock'
      };

      await InventoryService.addItem(this.props.sp, newAsset, this.props.userDisplayName);
      await this._loadInventory(); // Refresh list
      await this._loadAuditLogs(); // Refresh audit logs
    } catch (error: any) {
      console.error("Failed to add asset:", error);
    }
  };

  private _onSubmitRequest = async (requestData: Omit<IRequest, 'id' | 'requestDate' | 'status'>): Promise<void> => {
    try {
      await InventoryService.addRequest(this.props.sp, requestData, this.props.userDisplayName);
      console.log('Successfully saved request to SharePoint');
      await this._loadRequests(); // Refresh list from SharePoint
      await this._loadAuditLogs(); // Refresh audit logs
    } catch (error: any) {
      console.error('Failed to save request to SharePoint AssetRequests list:', error);
      
      this.setState({
        errorMessage: `Failed to save Request. SharePoint rejected the column names. Error: ${error.message || JSON.stringify(error)}`
      });
    }
  };

  public render(): React.ReactElement<IInventoryManagementProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const { items, isAssetFormOpen, isRequestFormOpen, auditLogs, auditLogsLoading } = this.state;

    return (
      <section className={`${styles.inventoryManagement} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.dark : ''}`}>
        <div className={styles.mainContent}>
          
          <div className={styles.heroSection}>
            <div className={styles.heroText}>
              <h2>Inventory Management</h2>
              <p>Welcome back, {escape(userDisplayName)}!</p>
              <span className={styles.smallText}>
                {environmentMessage} • Location: {escape(description)}
              </span>
            </div>
            <img 
              alt="Welcome" 
              src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} 
              className={styles.welcomeImage} 
            />
          </div>

          <div className={styles.actionGrid}>
            <div className={styles.actionButtonContainer}>
              <PrimaryButton 
                text="Add New Asset" 
                onClick={() => this.setState({ isAssetFormOpen: true })} 
                iconProps={{ iconName: 'Add' }}
              />
            </div>
            <div className={styles.actionButtonContainer}>
              <PrimaryButton 
                text="Request Asset" 
                onClick={() => this.setState({ isRequestFormOpen: true })} 
                iconProps={{ iconName: 'Send' }}
              />
            </div>
          </div>

          <div className={styles.card}>
            <Pivot aria-label="Inventory Management Views">
              <PivotItem headerText="Dashboard" itemIcon="BarChart4">
                <Dashboard items={items} requests={this.state.requests} />
              </PivotItem>
              <PivotItem headerText="Inventory" itemIcon="List">
                <div style={{ marginTop: '20px' }}>
                  <div className={styles.cardHeader}>
                    <h3>Current Inventory Overview</h3>
                  </div>
                  <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                    Track and manage your organizational assets efficiently within the SharePoint Framework.
                  </p>
                  {this.state.errorMessage ? (
                    <div style={{ color: '#991b1b', backgroundColor: '#fee2e2', padding: '15px', borderRadius: '8px', marginBottom: '20px' }}>
                      <strong>Error:</strong> {this.state.errorMessage}
                      <p style={{ fontSize: '0.8rem', marginTop: '10px' }}>
                        Please check if the list "InventoryList" exists and all column names are correct.
                      </p>
                    </div>
                  ) : this.state.loading ? (
                    <p>Loading inventory...</p>
                  ) : (
                    <InventoryList items={items} />
                  )}
                </div>
              </PivotItem>
              <PivotItem headerText="Requests" itemIcon="Send">
                <div style={{ marginTop: '20px' }}>
                  <div className={styles.cardHeader}>
                    <h3>Asset Requests Overview</h3>
                  </div>
                  <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                    Track and manage all asset requests efficiently.
                  </p>
                  <RequestList items={this.state.requests} />
                </div>
              </PivotItem>
              <PivotItem headerText="Event Stream" itemIcon="ActivityFeed">
                <EventStream 
                  logs={auditLogs} 
                  loading={auditLogsLoading} 
                  errorMessage={undefined} // Pass error message if needed
                />
              </PivotItem>
            </Pivot>
          </div>

        </div>

        <AssetForm 
          isOpen={isAssetFormOpen} 
          onClose={() => this.setState({ isAssetFormOpen: false })} 
          onAddAsset={this._onAddAsset} 
        />

        <RequestForm 
          isOpen={isRequestFormOpen} 
          onClose={() => this.setState({ isRequestFormOpen: false })} 
          availableAssets={items}
          employees={EMPLOYEES}
          onSubmitRequest={this._onSubmitRequest} 
        />
      </section>
    );
  }
}
