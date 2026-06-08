import * as React from 'react';
import styles from './InventoryManagement.module.scss';
import type { IInventoryManagementProps } from '../models/IInventoryManagementProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { InventoryList } from './InventoryList';
import { RequestList } from './RequestList';
import { IInventoryItem } from '../models/IInventoryItem';
import { IRequest } from '../models/IRequest';
import { IEventLog } from '../models/IEventLog';
import { getSP } from '../pnpjsConfig';
import { AssetForm } from './AssetForm';
import { RequestForm } from './RequestForm';
import { EventStream } from './EventStream';
import { PrimaryButton, Pivot, PivotItem, TextField, DetailsList, DetailsListLayoutMode, SelectionMode, IColumn, DetailsRow } from '@fluentui/react';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  ArcElement,
  Title,
  Tooltip,
  Legend,
} from 'chart.js';
import { Pie } from 'react-chartjs-2';

ChartJS.register(
  CategoryScale,
  LinearScale,
  ArcElement,
  Title,
  Tooltip,
  Legend
);
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

import { EMPLOYEES } from '../data/mockData';
import { InventoryService } from '../services/InventoryService';
import { Dashboard } from './Dashboard';
import { AssetTracking } from './AssetTracking';
import { INotification } from '../models/INotification';
import { NotificationCenter } from './NotificationCenter';

export interface IInventoryManagementState {
  items: IInventoryItem[];
  requests: IRequest[];
  auditLogs: IEventLog[];
  userRole: 'Admin' | 'Inventory Manager' | 'Inventory Employee';
  previewRole?: 'Admin' | 'Inventory Manager' | 'Inventory Employee';
  roleGroups: string[];
  roleLoading: boolean;
  requestActionInProgressId?: string;
  requestSearchId: string;
  isAssetFormOpen: boolean;
  isRequestFormOpen: boolean;
  loading: boolean;
  auditLogsLoading: boolean;
  errorMessage?: string;
  isTrackingActionInProgress?: boolean;
  expandedUserEmail?: string;
  selectedTabKey?: string;
  readNotificationIds: string[];
  clearedNotificationIds: string[];
}

export default class InventoryManagement extends React.Component<IInventoryManagementProps, IInventoryManagementState> {
  private _isRequestOwnedByCurrentUser = (requesterName: string, currentUser: string): boolean => {
    const normalize = (value: string): string => (value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const requestOwner = normalize(requesterName);
    const activeUser = normalize(currentUser);
    if (!requestOwner || !activeUser) {
      return false;
    }
    return requestOwner === activeUser || requestOwner.includes(activeUser) || activeUser.includes(requestOwner);
  };

  private _isAssetAssignedToCurrentUser = (item: IInventoryItem, currentUser: string): boolean => {
    const normalize = (value: string | undefined): string => (value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const activeUser = normalize(currentUser);
    if (!activeUser) return false;

    const assignedNorm = normalize(item.assignedTo);
    const isAssigned = assignedNorm && (assignedNorm === activeUser || assignedNorm.includes(activeUser) || activeUser.includes(assignedNorm));
    const isNoted = (item.note || '').toLowerCase().includes('assigned to:') && normalize(item.note).includes(activeUser);
    const isStatus = (item.status || '').toLowerCase().includes('assigned to:') && normalize(item.status).includes(activeUser);

    return !!(isAssigned || isNoted || isStatus);
  };

  private _getNotifications = (): INotification[] => {
    const { items, requests } = this.state;
    const currentUser = this.props.userDisplayName;
    const effectiveRole = this.state.previewRole || this.state.userRole;
    const isAdminOrManager = effectiveRole === 'Admin' || effectiveRole === 'Inventory Manager';
    const isAdmin = effectiveRole === 'Admin';

    const notifications: INotification[] = [];
    const readIds = new Set(this.state.readNotificationIds);
    const clearedIds = new Set(this.state.clearedNotificationIds);

    const normalize = (value: string | undefined): string => (value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const activeUserNorm = normalize(currentUser);

    const formatTime = (isoString: string | undefined): string => {
      if (!isoString) return '';
      try {
        const d = new Date(isoString);
        const pad = (n: number) => n < 10 ? '0' + n : '' + n;
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
      } catch {
        return isoString;
      }
    };

    // 1. Generate Asset Request Notifications
    requests.forEach(req => {
      const requesterNorm = normalize(req.requesterName);
      const isMyRequest = requesterNorm && (requesterNorm === activeUserNorm || activeUserNorm.includes(requesterNorm) || requesterNorm.includes(activeUserNorm));

      if (isAdminOrManager) {
        // Pending requests notify Admins & Managers
        if (req.status === 'Pending') {
          const id = `req-pending-${req.id}`;
          if (!clearedIds.has(id)) {
            notifications.push({
              id,
              title: "Asset Request Pending",
              message: `${req.requesterName} requested ${req.quantity}x ${req.assetTitle} (Reason: ${req.reason || "None"})`,
              type: 'info',
              timestamp: formatTime(req.requestDate),
              isRead: readIds.has(id),
              actionLink: 'Approvals',
              category: 'Request'
            });
          }
        }
      }

      if (isMyRequest) {
        // Approved/Declined requests notify the Employee
        if (req.status === 'Approved' || req.status === 'Declined') {
          const id = `req-resolved-${req.id}-${req.status}`;
          if (!clearedIds.has(id)) {
            notifications.push({
              id,
              title: req.status === 'Approved' ? "Request Approved" : "Request Declined",
              message: req.status === 'Approved'
                ? `Your request for ${req.quantity}x ${req.assetTitle} has been approved.`
                : `Your request for ${req.quantity}x ${req.assetTitle} has been declined.`,
              type: req.status === 'Approved' ? 'success' : 'error',
              timestamp: formatTime(req.requestDate),
              isRead: readIds.has(id),
              actionLink: 'My Requests',
              category: 'Request'
            });
          }
        }
      }

      if (isAdmin && req.status === 'Approved' && req.assetStatus === 'Pending') {
        const id = `req-assign-admin-${req.id}`;
        if (!clearedIds.has(id)) {
          notifications.push({
            id,
            title: "Asset Ready for Assignment",
            message: `${req.requesterName}'s request for ${req.quantity}x ${req.assetTitle} is approved and ready for assignment.`,
            type: 'info',
            timestamp: formatTime(req.requestDate),
            isRead: readIds.has(id),
            actionLink: 'AssetAssignmentQueue',
            category: 'Request'
          });
        }
      }
    });

    // 2. Generate Asset Assignment & Audit Notifications
    items.forEach(item => {
      const assignedNorm = normalize(item.assignedTo);
      const isMyAsset = assignedNorm && (assignedNorm === activeUserNorm || activeUserNorm.includes(assignedNorm) || assignedNorm.includes(activeUserNorm));

      const isNotedMyAsset = (item.note || '').toLowerCase().includes('assigned to:') && normalize(item.note).includes(activeUserNorm);

      if (isMyAsset || isNotedMyAsset) {
        // Asset Assignment notifies the Employee
        const id = `asset-assigned-${item.id}`;
        if (!clearedIds.has(id)) {
          notifications.push({
            id,
            title: "Asset Assigned",
            message: `Asset '${item.assetName || item.title}' (${item.serialNumber || 'N/A'}) has been assigned to you.`,
            type: 'success',
            timestamp: formatTime(item.assignedDate || item.purchaseDate),
            isRead: readIds.has(id),
            actionLink: 'My Assets',
            category: 'Assignment'
          });
        }
      }

      if (isAdminOrManager) {
        // When status is 'Assigned', notify Admin/Manager of assignments
        if (item.status === 'Assigned') {
          const id = `asset-assigned-admin-${item.id}`;
          if (!clearedIds.has(id)) {
            notifications.push({
              id,
              title: "Asset Assigned to Employee",
              message: `Asset '${item.assetName || item.title}' (${item.serialNumber || 'N/A'}) is assigned to ${item.assignedTo || "Employee"}.`,
              type: 'info',
              timestamp: formatTime(item.assignedDate || item.purchaseDate),
              isRead: readIds.has(id),
              actionLink: 'Asset Tracking',
              category: 'Assignment'
            });
          }
        }

        // Audit/Maintenance warnings
        if (item.status === 'Under Maintenance' || item.condition === 'Damaged' || item.condition === 'Poor') {
          const id = `asset-maintenance-${item.id}-${item.status}-${item.condition}`;
          if (!clearedIds.has(id)) {
            notifications.push({
              id,
              title: "Asset Status Alert",
              message: `Asset '${item.assetName || item.title}' is in ${item.condition} condition and marked as ${item.status}.`,
              type: 'warning',
              timestamp: formatTime(item.purchaseDate),
              isRead: readIds.has(id),
              actionLink: 'Inventory',
              category: 'Audit'
            });
          }
        }
      }
    });

    // Sort notifications by timestamp descending
    notifications.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
    return notifications;
  };

  private _markNotificationAsRead = (id: string): void => {
    const readNotificationIds = [...this.state.readNotificationIds, id];
    this.setState({ readNotificationIds });
    localStorage.setItem('inventory_read_notifications', JSON.stringify(readNotificationIds));
  };

  private _markAllNotificationsAsRead = (): void => {
    const notifications = this._getNotifications();
    const readNotificationIds = Array.from(new Set([...this.state.readNotificationIds, ...notifications.map(n => n.id)]));
    this.setState({ readNotificationIds });
    localStorage.setItem('inventory_read_notifications', JSON.stringify(readNotificationIds));
  };

  private _clearNotification = (id: string): void => {
    const clearedNotificationIds = [...this.state.clearedNotificationIds, id];
    this.setState({ clearedNotificationIds });
    localStorage.setItem('inventory_cleared_notifications', JSON.stringify(clearedNotificationIds));
  };

  private _clearAllNotifications = (): void => {
    const notifications = this._getNotifications();
    const clearedNotificationIds = Array.from(new Set([...this.state.clearedNotificationIds, ...notifications.map(n => n.id)]));
    this.setState({ clearedNotificationIds });
    localStorage.setItem('inventory_cleared_notifications', JSON.stringify(clearedNotificationIds));
  };

  private _handleNotificationAction = (actionLink: string, notificationId: string): void => {
    this._markNotificationAsRead(notificationId);
    this.setState({ selectedTabKey: actionLink });
  };

  constructor(props: IInventoryManagementProps) {
    super(props);

    let readIds: string[] = [];
    let clearedIds: string[] = [];
    try {
      readIds = JSON.parse(localStorage.getItem('inventory_read_notifications') || '[]');
      clearedIds = JSON.parse(localStorage.getItem('inventory_cleared_notifications') || '[]');
    } catch (e) {
      console.warn("localStorage parsing failed", e);
    }

    this.state = {
      items: [],
      requests: [],
      auditLogs: [],
      userRole: 'Inventory Employee',
      previewRole: undefined,
      roleGroups: [],
      roleLoading: true,
      requestActionInProgressId: undefined,
      requestSearchId: '',
      isAssetFormOpen: false,
      isRequestFormOpen: false,
      loading: true,
      auditLogsLoading: true,
      errorMessage: undefined,
      selectedTabKey: 'Dashboard',
      readNotificationIds: readIds,
      clearedNotificationIds: clearedIds
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._resolveUserRole();
    await this._loadInventory();
    await this._loadRequests();
    await this._loadAuditLogs();

    // Dynamically auto-sync existing assigned assets of our 5 active users to the Mapping List
    try {
      await InventoryService.syncExistingAssignmentsToMappingList(this.props.userDisplayName);
    } catch (e) {
      console.warn("Failed to auto-sync existing assignments to Mapping List:", e);
    }
  }

  private _resolveUserRole = async (): Promise<void> => {
    try {
      const normalizedDisplayName = (this.props.userDisplayName || '').toLowerCase().trim();
      const compactDisplayName = normalizedDisplayName.replace(/\s+/g, '');
      const normalizedEmail = (this.props.userEmail || '').toLowerCase().trim();

      const forcedRoleByUserKey: Record<string, 'Admin' | 'Inventory Manager' | 'Inventory Employee'> = {
        'kiran.reddy@3bh3kf.onmicrosoft.com': 'Admin',
        'adelev@3bh3kf.onmicrosoft.com': 'Inventory Employee',
        'alexw@3bh3kf.onmicrosoft.com': 'Inventory Employee',
        'diegos@3bh3kf.onmicrosoft.com': 'Inventory Manager',
        'gradya@3bh3kf.onmicrosoft.com': 'Inventory Manager'
      };

      let forcedRole: 'Admin' | 'Inventory Manager' | 'Inventory Employee' | undefined;
      for (const userKey in forcedRoleByUserKey) {
        if (
          compactDisplayName.includes(userKey) ||
          normalizedDisplayName.includes(userKey) ||
          normalizedEmail.includes(userKey)
        ) {
          forcedRole = forcedRoleByUserKey[userKey];
          break;
        }
      }

      if (forcedRole) {
        this.setState({
          userRole: forcedRole,
          roleGroups: [],
          roleLoading: false
        });
        return;
      }

      const sp = getSP();
      const groups = await sp.web.currentUser.groups();
      const groupNames = groups.map((group: any) => (group.Title || '').toLowerCase());

      const isAdmin = groupNames.some((name: string) => name.includes('admin'));
      const isInventoryManager = groupNames.some((name: string) => name.includes('inventory manager'));
      const isInventoryEmployee = groupNames.some((name: string) => name.includes('inventory employee'));

      let userRole: 'Admin' | 'Inventory Manager' | 'Inventory Employee' = 'Inventory Employee';

      if (isAdmin) {
        userRole = 'Admin';
      } else if (isInventoryManager) {
        userRole = 'Inventory Manager';
      } else if (isInventoryEmployee) {
        userRole = 'Inventory Employee';
      }

      this.setState({
        userRole,
        roleGroups: groups.map((group: any) => group.Title || ''),
        roleLoading: false
      });
    } catch (error) {
      console.error("Failed to resolve SharePoint group role:", error);
      this.setState({
        userRole: 'Inventory Employee',
        roleGroups: [],
        roleLoading: false
      });
    }
  };

  private _loadInventory = async (): Promise<void> => {
    try {
      this.setState({ loading: true, errorMessage: undefined });
      const items = await InventoryService.getItems();
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
      const requests = await InventoryService.getRequests();
      this.setState({ requests });
    } catch (error: any) {
      console.error("Failed to load requests:", error);
      this.setState({
        errorMessage: `Failed to load Requests. Error: ${error.message || JSON.stringify(error)}`
      });
    }
  };

  private _loadAuditLogs = async (): Promise<void> => {
    try {
      this.setState({ auditLogsLoading: true });
      const auditLogs = await InventoryService.getAuditLogs();
      this.setState({ auditLogs, auditLogsLoading: false });
    } catch (error) {
      console.error("Failed to load audit logs:", error);
      this.setState({ auditLogsLoading: false });
    }
  };

  private _onAddAsset = async (newAssetData: Omit<IInventoryItem, 'id' | 'status' | 'assignedTo'>): Promise<void> => {
    try {
      this.setState({ loading: true, errorMessage: undefined });
      const newAsset: Omit<IInventoryItem, 'id'> = {
        ...newAssetData,
        status: 'In Stock'
      };

      await InventoryService.addItem(newAsset, this.props.userDisplayName);
      await this._loadInventory(); // Refresh list
      await this._loadAuditLogs(); // Refresh audit logs
    } catch (error: any) {
      console.error("Failed to add asset:", error);
      this.setState({
        loading: false,
        errorMessage: `Failed to add Asset. SharePoint rejected the save. Error: ${error.message || JSON.stringify(error)}`
      });
    }
  };

  private _onSubmitRequest = async (requestData: Omit<IRequest, 'id' | 'requestKey' | 'requestDate' | 'status'>): Promise<void> => {
    try {
      const requesterEmployee = EMPLOYEES.find(e => e.name.toLowerCase() === requestData.requesterName.toLowerCase());
      const requesterRole = requesterEmployee ? requesterEmployee.jobTitle : 'Inventory Employee';
      const initialStatus = requesterRole === 'Inventory Manager' ? 'Approved' : 'Pending';

      const tempId = `temp-${Date.now()}`;
      const localRequest: IRequest = {
        id: tempId,
        requestKey: `REQ-${("000000" + (this.state.requests.length + 1)).slice(-6)}`,
        requesterName: requestData.requesterName,
        employeeId: (requestData as any).employeeId || "",
        assetId: requestData.assetId || "1",
        assetTitle: requestData.assetTitle,
        assetName: "",
        priority: (requestData as any).priority || "Medium",
        quantity: requestData.quantity || 1,
        status: initialStatus,
        assetStatus: 'Pending',
        requestDate: new Date().toISOString().split('T')[0],
        reason: requestData.reason || ""
      };

      // Optimistic update so it gets added to the RequestList directly
      this.setState(prevState => ({
        requests: [localRequest, ...prevState.requests]
      }));

      await InventoryService.addRequest({
        ...requestData,
        status: initialStatus
      } as any, this.props.userDisplayName);
      console.log('Successfully saved request to SharePoint');
      await this._loadRequests(); // Refresh list from SharePoint
      await this._loadAuditLogs(); // Refresh audit logs
    } catch (error: any) {
      console.error('Failed to save request to SharePoint AssetRequests list:', error);

      this.setState({
        errorMessage: `Failed to save Request to SharePoint. A local copy was added. Error: ${error.message || JSON.stringify(error)}`
      });
    }
  };

  private _onApproveRequest = async (request: IRequest): Promise<void> => {
    try {
      this.setState({ requestActionInProgressId: request.id, errorMessage: undefined });
      if (request.id.indexOf('temp-') === 0) {
        this.setState(prevState => ({
          requests: prevState.requests.map(r => r.id === request.id ? { ...r, status: 'Approved' } : r)
        }));
      } else {
        await InventoryService.updateRequestStatus(parseInt(request.id, 10), 'Approved', this.props.userDisplayName);
        await this._loadRequests();
        await this._loadAuditLogs();
      }
    } catch (error: any) {
      this.setState({
        errorMessage: `Failed to approve request #${request.id}. ${error.message || JSON.stringify(error)}`
      });
    } finally {
      this.setState({ requestActionInProgressId: undefined });
    }
  };

  private _onRejectRequest = async (request: IRequest, reason: string): Promise<void> => {
    try {
      this.setState({ requestActionInProgressId: request.id, errorMessage: undefined });
      if (request.id.indexOf('temp-') === 0) {
        this.setState(prevState => ({
          requests: prevState.requests.map(r => r.id === request.id ? { ...r, status: 'Declined', managerResponse: reason } : r)
        }));
      } else {
        await InventoryService.updateRequestStatus(
          parseInt(request.id, 10),
          'Declined',
          this.props.userDisplayName,
          reason
        );
        await this._loadRequests();
        await this._loadAuditLogs();
      }
    } catch (error: any) {
      this.setState({
        errorMessage: `Failed to reject request #${request.id}. ${error.message || JSON.stringify(error)}`
      });
    } finally {
      this.setState({ requestActionInProgressId: undefined });
    }
  };

  private _onApproveAsset = async (request: IRequest): Promise<void> => {
    try {
      this.setState({ requestActionInProgressId: request.id, errorMessage: undefined });
      if (request.id.indexOf('temp-') === 0) {
        this.setState(prevState => ({
          requests: prevState.requests.map(r => r.id === request.id ? { ...r, assetStatus: 'Approved' } : r)
        }));
      } else {
        await InventoryService.updateAssetStatus(parseInt(request.id, 10), 'Approved', this.props.userDisplayName);
        await this._loadRequests();
        await this._loadAuditLogs();
      }
    } catch (error: any) {
      this.setState({
        errorMessage: `Failed to approve asset status for request #${request.requestKey || request.id}. ${error.message || JSON.stringify(error)}`
      });
    } finally {
      this.setState({ requestActionInProgressId: undefined });
    }
  };

  private _onAssignAssets = async (employeeName: string, employeeEmail: string, assetIds: string[]): Promise<void> => {
    try {
      this.setState({ isTrackingActionInProgress: true, errorMessage: undefined });
      const employee = EMPLOYEES.find(e => e.name.toLowerCase() === employeeName.toLowerCase());
      const employeeId = employee ? employee.id : "";
      await InventoryService.assignAssetsToEmployee(assetIds, employeeName, employeeEmail, this.props.userDisplayName, employeeId);
      await this._loadInventory();
      await this._loadAuditLogs();
    } catch (error: any) {
      this.setState({
        errorMessage: `Failed to assign assets. ${error.message || JSON.stringify(error)}`
      });
    } finally {
      this.setState({ isTrackingActionInProgress: false });
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

    const { items, isAssetFormOpen, isRequestFormOpen, auditLogs, auditLogsLoading, userRole, previewRole, roleLoading, roleGroups, requestActionInProgressId, requestSearchId } = this.state;


    const effectiveRole = previewRole || userRole;

    const isAdmin = effectiveRole === 'Admin';
    const isManager = effectiveRole === 'Inventory Manager';
    const isEmployee = effectiveRole === 'Inventory Employee';

    const myAssets = items.filter(item => this._isAssetAssignedToCurrentUser(item, userDisplayName || ''));

    const myRequests = this.state.requests.filter(request => this._isRequestOwnedByCurrentUser(request.requesterName || '', userDisplayName || ''));
    const myApprovedRequests = myRequests.filter(request => (request.status || '').toLowerCase().includes('approv'));

    const adminQueueRequests = this.state.requests.filter(request => (request.status || '').toLowerCase().includes('approv'));
    const managerQueueRequests = this.state.requests;

    const normalizedSearch = (requestSearchId || '').trim().toLowerCase();
    const filterRequests = (reqs: IRequest[]) => normalizedSearch
      ? reqs.filter(request =>
        (request.requestKey || '').toLowerCase().includes(normalizedSearch) ||
        (request.id || '').toLowerCase().includes(normalizedSearch)
      )
      : reqs;

    const visibleAdminRequests = filterRequests(adminQueueRequests);
    const visibleManagerRequests = filterRequests(managerQueueRequests);
    const notifications = this._getNotifications();

    return (
      <section className={`${styles.inventoryManagement} ${hasTeamsContext ? styles.teams : ''} ${isDarkTheme ? styles.dark : ''}`}>
        <div className={styles.mainContent}>

          <div className={styles.heroSection}>
            <div className={styles.heroText}>
              <h2>Inventory Management</h2>
              <p>Welcome back, {escape(userDisplayName)}!</p>
              <p className={styles.smallText}>
                Role: <strong>{effectiveRole}</strong>
              </p>

              <span className={styles.smallText}>
                {environmentMessage} • Location: {escape(description)}
              </span>
              {isAdmin && roleGroups.length > 0 && (
                <p className={styles.smallText}>
                  SharePoint Groups: {escape(roleGroups.join(', '))}
                </p>
              )}
            </div>
            <img
              alt="Welcome"
              src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')}
              className={styles.welcomeImage}
            />
          </div>

          {this.state.errorMessage && (
            <div style={{ color: '#991b1b', backgroundColor: '#fee2e2', padding: '15px', borderRadius: '8px', marginBottom: '20px', position: 'relative' }}>
              <strong>Error:</strong> {this.state.errorMessage}
              <button
                onClick={() => this.setState({ errorMessage: undefined })}
                style={{ position: 'absolute', right: '15px', top: '12px', border: 'none', background: 'none', cursor: 'pointer', fontWeight: 'bold', fontSize: '1.1rem', color: '#991b1b' }}
                aria-label="Dismiss error"
              >
                &times;
              </button>
            </div>
          )}

          {!roleLoading && (
            <div className={styles.actionGrid}>
              {(isAdmin || isManager) && (
                <div className={styles.actionButtonContainer}>
                  <PrimaryButton
                    text={isAdmin ? "Add New Asset" : "Assign / Manage Assets"}
                    onClick={() => this.setState({ isAssetFormOpen: true })}
                    iconProps={{ iconName: 'Add' }}
                  />
                </div>
              )}
              {(isAdmin || isManager || isEmployee) && (
                <div className={styles.actionButtonContainer}>
                  <PrimaryButton
                    text="Request Asset"
                    onClick={() => this.setState({ isRequestFormOpen: true })}
                    iconProps={{ iconName: 'Send' }}
                  />
                </div>
              )}
            </div>
          )}

          <div className={styles.card}>
            <Pivot
              aria-label="Inventory Management Views"
              selectedKey={this.state.selectedTabKey}
              onLinkClick={(item) => item && this.setState({ selectedTabKey: item.props.itemKey })}
            >
              <PivotItem headerText="Dashboard" itemIcon="BarChart4" itemKey="Dashboard">
                <Dashboard
                  items={isAdmin || isManager ? items : myAssets}
                  requests={isAdmin || isManager ? this.state.requests : myRequests}
                  isAdmin={isAdmin}
                  isInventoryManager={isManager}
                />
              </PivotItem>
              {/* Universal Tabs for Everyone */}
              <PivotItem headerText="My Assets" itemIcon="Broom" itemKey="MyAssets">
                <div style={{ marginTop: '20px' }}>
                  <div className={styles.cardHeader}>
                    <h3>My Assigned Assets</h3>
                  </div>
                  <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                    View assets currently assigned to you.
                  </p>
                  <InventoryList items={myAssets} isAdmin={false} />
                </div>
              </PivotItem>

              <PivotItem headerText="My Requests" itemIcon="Send" itemKey="MyRequests">
                <div style={{ marginTop: '20px' }}>
                  <div className={styles.cardHeader}>
                    <h3>My Asset Requests</h3>
                  </div>
                  <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                    Track your submitted requests and approval status.
                  </p>
                  <RequestList items={myRequests} showResponseColumns={true} />
                  <div style={{ marginTop: '20px' }}>
                    <h4 style={{ marginBottom: '8px' }}>Approved Asset Details (from Requests)</h4>
                    <p style={{ color: 'var(--text-muted)', marginBottom: '12px' }}>
                      These are your approved asset requests, even if assignment is still pending.
                    </p>
                    <RequestList items={myApprovedRequests} showResponseColumns={true} />
                  </div>
                </div>
              </PivotItem>

              <PivotItem
                headerText={notifications.filter(n => !n.isRead).length > 0 ? `Notifications (${notifications.filter(n => !n.isRead).length})` : "Notifications"}
                itemIcon="Ringer"
                itemKey="Notifications"
              >
                <NotificationCenter
                  notifications={notifications}
                  onMarkAsRead={this._markNotificationAsRead}
                  onMarkAllAsRead={this._markAllNotificationsAsRead}
                  onClearNotification={this._clearNotification}
                  onClearAllNotifications={this._clearAllNotifications}
                  onNotificationAction={this._handleNotificationAction}
                />
              </PivotItem>

              {/* Admin-only Inventory */}
              {isAdmin && (
                <PivotItem headerText="Inventory" itemIcon="List" itemKey="Inventory">
                  <div style={{ marginTop: '20px' }}>
                    <div className={styles.cardHeader}>
                      <h3>Current Inventory Overview</h3>
                    </div>
                    <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                      Track and manage your organizational assets efficiently within the SharePoint Framework.
                    </p>
                    {this.state.loading ? (
                      <p>Loading inventory...</p>
                    ) : (
                      <InventoryList items={items} isAdmin={true} />
                    )}
                  </div>
                </PivotItem>
              )}

              {/* Manager Approvals Queue */}
              {isManager && (
                <PivotItem headerText="Approvals" itemIcon="DoubleChevronRight12" itemKey="Approvals">
                  <div style={{ marginTop: '20px' }}>
                    <div className={styles.cardHeader}>
                      <h3>Request Approvals & Assignment Queue</h3>
                    </div>
                    <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                      Track and manage all asset requests efficiently.
                    </p>
                    <TextField
                      label="Search by Request ID"
                      placeholder="e.g. REQ-000123"
                      value={requestSearchId}
                      onChange={(_, value) => this.setState({ requestSearchId: value || '' })}
                      styles={{ root: { marginBottom: '12px', maxWidth: 320 } }}
                    />
                    <div style={{ marginBottom: '20px', padding: '15px', backgroundColor: 'var(--surface-color, #ffffff)', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                      <h4 style={{ marginBottom: '10px' }}>Request Approval Distribution</h4>
                      <div style={{ height: '250px', position: 'relative' }}>
                        <Pie
                          data={{
                            labels: Object.keys(
                              managerQueueRequests.reduce((acc, req) => {
                                const status = req.status || 'Pending';
                                acc[status] = (acc[status] || 0) + 1;
                                return acc;
                              }, {} as Record<string, number>)
                            ).length ? Object.keys(
                              managerQueueRequests.reduce((acc, req) => {
                                const status = req.status || 'Pending';
                                acc[status] = (acc[status] || 0) + 1;
                                return acc;
                              }, {} as Record<string, number>)
                            ) : ['No data'],
                            datasets: [
                              {
                                label: 'Requests by Status',
                                data: Object.keys(
                                  managerQueueRequests.reduce((acc, req) => {
                                    const status = req.status || 'Pending';
                                    acc[status] = (acc[status] || 0) + 1;
                                    return acc;
                                  }, {} as Record<string, number>)
                                ).length ? Object.keys(
                                  managerQueueRequests.reduce((acc, req) => {
                                    const status = req.status || 'Pending';
                                    acc[status] = (acc[status] || 0) + 1;
                                    return acc;
                                  }, {} as Record<string, number>)
                                ).map(k =>
                                  (managerQueueRequests.reduce((acc, req) => {
                                    const status = req.status || 'Pending';
                                    acc[status] = (acc[status] || 0) + 1;
                                    return acc;
                                  }, {} as Record<string, number>))[k]
                                ) : [1],
                                backgroundColor: [
                                  'rgba(255, 206, 86, 0.6)',
                                  'rgba(75, 192, 192, 0.6)',
                                  'rgba(255, 99, 132, 0.6)',
                                  'rgba(153, 102, 255, 0.6)',
                                  'rgba(54, 162, 235, 0.6)',
                                ],
                                borderColor: [
                                  'rgba(255, 206, 86, 1)',
                                  'rgba(75, 192, 192, 1)',
                                  'rgba(255, 99, 132, 1)',
                                  'rgba(153, 102, 255, 1)',
                                  'rgba(54, 162, 235, 1)',
                                ],
                                borderWidth: 1,
                              },
                            ],
                          }}
                          options={{ maintainAspectRatio: false }}
                        />
                      </div>
                    </div>
                    <RequestList
                      items={visibleManagerRequests}
                      canApproveReject={true}
                      canApproveAsset={false}
                      hideStatusColumn={false}
                      showResponseColumns={false}
                      onApproveRequest={this._onApproveRequest}
                      onRejectRequest={this._onRejectRequest}
                      actionInProgressId={requestActionInProgressId}
                    />
                  </div>
                </PivotItem>
              )}

              {/* Admin Asset Assignment Queue */}
              {isAdmin && (
                <PivotItem headerText="Asset Assignment Queue" itemIcon="Send" itemKey="AssetAssignmentQueue">
                  <div style={{ marginTop: '20px' }}>
                    <div className={styles.cardHeader}>
                      <h3>Approved Requests for Asset Assignment</h3>
                    </div>
                    <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                      Only approved requests are shown here so assets can be assigned.
                    </p>
                    <TextField
                      label="Search by Request ID"
                      placeholder="e.g. REQ-000123"
                      value={requestSearchId}
                      onChange={(_, value) => this.setState({ requestSearchId: value || '' })}
                      styles={{ root: { marginBottom: '12px', maxWidth: 320 } }}
                    />
                    <RequestList
                      items={visibleAdminRequests}
                      canApproveReject={false}
                      canApproveAsset={true}
                      hideStatusColumn={true}
                      showResponseColumns={false}
                      onApproveAsset={this._onApproveAsset}
                      actionInProgressId={requestActionInProgressId}
                    />
                  </div>
                </PivotItem>
              )}
              {(isAdmin || isManager) && (
                <PivotItem headerText="Event Stream" itemIcon="ActivityFeed" itemKey="EventStream">
                  <EventStream
                    logs={auditLogs}
                    loading={auditLogsLoading}
                    errorMessage={undefined}
                    currentUserRole={effectiveRole}
                    currentUserName={userDisplayName}
                  />
                </PivotItem>
              )}
              {isAdmin && (
                <PivotItem headerText="Users" itemIcon="People" itemKey="Users">
                  <div style={{ marginTop: '20px' }}>
                    <div className={styles.cardHeader}>
                      <h3>User Administration</h3>
                    </div>
                    <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                      Admin-only area. Manage SharePoint groups and user onboarding from your site permissions.
                    </p>
                    <div style={{ marginBottom: '20px', padding: '15px', backgroundColor: '#f0f6ff', borderRadius: '8px', borderLeft: '4px solid #0078d4' }}>
                      <h4 style={{ marginTop: 0, marginBottom: '10px', color: '#0078d4' }}>SharePoint Group Management</h4>
                      <p style={{ margin: 0, fontSize: '0.9rem', color: '#323130', marginBottom: '15px' }}>
                        To onboard new employees, grant them Admin access, or assign them as Inventory Managers, you must add them to the respective SharePoint Site Groups.
                      </p>
                      <PrimaryButton
                        text="Manage Site Permissions"
                        iconProps={{ iconName: 'Permissions' }}
                        onClick={() => {
                          const siteUrl = window.location.pathname.substring(0, window.location.pathname.toLowerCase().indexOf('/sitepages'));
                          window.open(`${window.location.origin}${siteUrl}/_layouts/15/user.aspx`, '_blank');
                        }}
                      />
                    </div>

                    <h4 style={{ marginBottom: '15px' }}>Employee Directory & Asset Ownership</h4>
                    <div style={{ backgroundColor: 'var(--surface-color, #ffffff)', padding: '15px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                      <DetailsList
                        items={EMPLOYEES.map(emp => {
                          const realName = emp.jobTitle === 'Admin' ? (this.props.userDisplayName || emp.name) : emp.name;
                          const assignedItems = items.filter(i => this._isAssetAssignedToCurrentUser(i, realName));
                          const assetTypes = Array.from(new Set(assignedItems.map(a => a.assetType))).filter(t => t).join(', ');
                          return {
                            ...emp,
                            assignedAssets: assignedItems.length,
                            assignedItems: assignedItems,
                            assetTypes: assetTypes || 'None'
                          };
                        })}
                        columns={[
                          { key: 'col1', name: 'Name', fieldName: 'name', minWidth: 100, maxWidth: 150, isResizable: true },
                          { key: 'col2', name: 'Email', fieldName: 'email', minWidth: 150, maxWidth: 200, isResizable: true },
                          { key: 'col3', name: 'Department', fieldName: 'department', minWidth: 100, maxWidth: 120, isResizable: true },
                          { key: 'col4', name: 'Job Title', fieldName: 'jobTitle', minWidth: 120, maxWidth: 150, isResizable: true },
                          {
                            key: 'col5',
                            name: 'Assigned Assets',
                            fieldName: 'assignedAssets',
                            minWidth: 100,
                            maxWidth: 120,
                            isResizable: true,
                            onRender: (item) => (
                              <span style={{
                                backgroundColor: item.assignedAssets > 0 ? '#dbeafe' : '#f3f4f6',
                                color: item.assignedAssets > 0 ? '#1e40af' : '#4b5563',
                                padding: '4px 10px',
                                borderRadius: '9999px',
                                fontWeight: 'bold'
                              }}>
                                {item.assignedAssets}
                              </span>
                            )
                          },
                          { key: 'col6', name: 'Asset Types', fieldName: 'assetTypes', minWidth: 120, maxWidth: 250, isResizable: true }
                        ]}
                        setKey="usersList"
                        layoutMode={DetailsListLayoutMode.justified}
                        selectionMode={SelectionMode.none}
                        onRenderRow={(rowProps) => {
                          if (!rowProps) return null;
                          const isExpanded = this.state.expandedUserEmail === rowProps.item.email;

                          return (
                            <div>
                              <div
                                onClick={() => this.setState({ expandedUserEmail: isExpanded ? undefined : rowProps.item.email })}
                                style={{ cursor: 'pointer', '&:hover': { backgroundColor: '#f3f2f1' } } as any}
                              >
                                <DetailsRow {...rowProps} />
                              </div>
                              {isExpanded && (
                                <div style={{ padding: '20px 40px', backgroundColor: '#f9fafb', borderBottom: '1px solid #e5e7eb' }}>
                                  <h4 style={{ marginTop: 0, marginBottom: '15px', color: '#111827' }}>Assets assigned to {rowProps.item.name}</h4>
                                  {rowProps.item.assignedItems.length > 0 ? (
                                    <InventoryList items={rowProps.item.assignedItems} isAdmin={false} />
                                  ) : (
                                    <p style={{ color: '#6b7280', fontSize: '0.9rem', margin: 0 }}>This user currently has no assets assigned to them.</p>
                                  )}
                                </div>
                              )}
                            </div>
                          );
                        }}
                      />
                    </div>
                  </div>
                </PivotItem>
              )}
              {isAdmin && (
                <PivotItem headerText="Asset Tracking" itemIcon="EntitlementPolicy" itemKey="AssetTracking">
                  <div style={{ marginTop: '20px' }}>
                    <div className={styles.cardHeader}>
                      <h3>Asset Tracking & Direct Assignment</h3>
                    </div>
                    <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                      Admin and Manager area. Select an employee to view their assigned assets or directly assign new assets from the inventory.
                    </p>
                    <AssetTracking
                      items={items}
                      employees={EMPLOYEES}
                      currentUserRole={effectiveRole}
                      currentUserName={this.props.userDisplayName}
                      currentUserEmail={this.props.userEmail}
                      onAssignAssets={this._onAssignAssets}
                      isActionInProgress={!!this.state.isTrackingActionInProgress}
                    />
                  </div>
                </PivotItem>
              )}
              {isAdmin && (
                <PivotItem headerText="Reports" itemIcon="ReportDocument" itemKey="Reports">
                  <div style={{ marginTop: '20px' }}>
                    <div className={styles.cardHeader}>
                      <h3>Reporting & Insights</h3>
                    </div>
                    <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                      Use dashboard and event history to derive utilization, approval trends, and asset aging reports.
                    </p>

                    <div style={{ marginTop: '20px', padding: '15px', backgroundColor: 'var(--surface-color, #ffffff)', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                      <h4 style={{ marginBottom: '10px' }}>Asset Utilization</h4>
                      <div style={{ display: 'flex', gap: '20px' }}>
                        <div style={{ padding: '10px 15px', backgroundColor: '#f3f4f6', borderRadius: '6px', flex: 1 }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', color: '#4b5563', marginBottom: '4px' }}>Total Assets</span>
                          <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#111827' }}>{items.length}</span>
                        </div>
                        <div style={{ padding: '10px 15px', backgroundColor: '#dbeafe', borderRadius: '6px', flex: 1 }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', color: '#1e40af', marginBottom: '4px' }}>In Use / Assigned</span>
                          <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#1e3a8a' }}>{items.length - items.filter(i => i.status === 'In Stock' || i.status === 'Yes').length}</span>
                        </div>
                        <div style={{ padding: '10px 15px', backgroundColor: '#dcfce7', borderRadius: '6px', flex: 1 }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', color: '#166534', marginBottom: '4px' }}>Utilization Rate</span>
                          <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#14532d' }}>
                            {items.length > 0 ? Math.round(((items.length - items.filter(i => i.status === 'In Stock' || i.status === 'Yes').length) / items.length) * 100) : 0}%
                          </span>
                        </div>
                      </div>
                    </div>

                    <div style={{ marginTop: '20px', padding: '15px', backgroundColor: 'var(--surface-color, #ffffff)', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                      <h4 style={{ marginBottom: '10px' }}>Approval Trends</h4>
                      <div style={{ display: 'flex', gap: '20px' }}>
                        <div style={{ padding: '10px 15px', backgroundColor: '#f3f4f6', borderRadius: '6px', flex: 1 }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', color: '#4b5563', marginBottom: '4px' }}>Total Requests</span>
                          <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#111827' }}>{this.state.requests.length}</span>
                        </div>
                        <div style={{ padding: '10px 15px', backgroundColor: '#dcfce7', borderRadius: '6px', flex: 1 }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', color: '#166534', marginBottom: '4px' }}>Approved</span>
                          <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#14532d' }}>{this.state.requests.filter(r => (r.status || '').toLowerCase().includes('approv')).length}</span>
                        </div>
                        <div style={{ padding: '10px 15px', backgroundColor: '#fee2e2', borderRadius: '6px', flex: 1 }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', color: '#991b1b', marginBottom: '4px' }}>Declined</span>
                          <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#7f1d1d' }}>{this.state.requests.filter(r => (r.status || '').toLowerCase().includes('declin') || (r.status || '').toLowerCase().includes('reject')).length}</span>
                        </div>
                        <div style={{ padding: '10px 15px', backgroundColor: '#fef3c7', borderRadius: '6px', flex: 1 }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', color: '#92400e', marginBottom: '4px' }}>Approval Rate</span>
                          <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#78350f' }}>
                            {this.state.requests.length > 0 ? Math.round((this.state.requests.filter(r => (r.status || '').toLowerCase().includes('approv')).length / this.state.requests.length) * 100) : 0}%
                          </span>
                        </div>
                      </div>
                    </div>

                    {(() => {
                      const now = new Date();
                      const aging = items.reduce((acc, item) => {
                        if (!item.purchaseDate) {
                          acc.unknown++;
                          return acc;
                        }
                        const pd = new Date(item.purchaseDate);
                        const diffYears = Math.abs(now.getTime() - pd.getTime()) / (1000 * 60 * 60 * 24 * 365);
                        if (diffYears < 1) acc.under1++;
                        else if (diffYears <= 3) acc.between1and3++;
                        else acc.over3++;
                        return acc;
                      }, { under1: 0, between1and3: 0, over3: 0, unknown: 0 });

                      return (
                        <div style={{ marginTop: '20px', padding: '15px', backgroundColor: 'var(--surface-color, #ffffff)', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                          <h4 style={{ marginBottom: '10px' }}>Asset Aging</h4>
                          <div style={{ display: 'flex', gap: '20px' }}>
                            <div style={{ padding: '10px 15px', backgroundColor: '#dcfce7', borderRadius: '6px', flex: 1 }}>
                              <span style={{ display: 'block', fontSize: '0.85rem', color: '#166534', marginBottom: '4px' }}>&lt; 1 Year Old (New)</span>
                              <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#14532d' }}>{aging.under1}</span>
                            </div>
                            <div style={{ padding: '10px 15px', backgroundColor: '#fef3c7', borderRadius: '6px', flex: 1 }}>
                              <span style={{ display: 'block', fontSize: '0.85rem', color: '#92400e', marginBottom: '4px' }}>1 - 3 Years Old</span>
                              <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#78350f' }}>{aging.between1and3}</span>
                            </div>
                            <div style={{ padding: '10px 15px', backgroundColor: '#fee2e2', borderRadius: '6px', flex: 1 }}>
                              <span style={{ display: 'block', fontSize: '0.85rem', color: '#991b1b', marginBottom: '4px' }}>&gt; 3 Years Old (Aging)</span>
                              <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#7f1d1d' }}>{aging.over3}</span>
                            </div>
                            <div style={{ padding: '10px 15px', backgroundColor: '#f3f4f6', borderRadius: '6px', flex: 1 }}>
                              <span style={{ display: 'block', fontSize: '0.85rem', color: '#4b5563', marginBottom: '4px' }}>Unknown Age</span>
                              <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#111827' }}>{aging.unknown}</span>
                            </div>
                          </div>
                        </div>
                      );
                    })()}

                    <div style={{ marginTop: '20px', padding: '15px', backgroundColor: 'var(--surface-color, #ffffff)', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
                      <h4 style={{ marginBottom: '10px' }}>Warranty Expiry Report</h4>
                      <div style={{ marginBottom: '15px', display: 'flex', gap: '20px' }}>
                        <div style={{ padding: '10px 15px', backgroundColor: '#f3f4f6', borderRadius: '6px' }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', color: '#4b5563', marginBottom: '4px' }}>Total Assets Count</span>
                          <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#111827' }}>{items.length}</span>
                        </div>
                        <div style={{ padding: '10px 15px', backgroundColor: '#f3f4f6', borderRadius: '6px' }}>
                          <span style={{ display: 'block', fontSize: '0.85rem', color: '#4b5563', marginBottom: '4px' }}>Assets with Warranty Data</span>
                          <span style={{ fontSize: '1.25rem', fontWeight: 'bold', color: '#111827' }}>{items.filter(i => i.warrantyExpiry).length}</span>
                        </div>
                      </div>
                      <DetailsList
                        items={items}
                        columns={[
                          { key: 'col1', name: 'Asset Name', fieldName: 'assetName', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.assetName || item.title },
                          { key: 'col2', name: 'Asset Type', fieldName: 'assetType', minWidth: 100, maxWidth: 150, isResizable: true },
                          { key: 'col3', name: 'Status', fieldName: 'status', minWidth: 80, maxWidth: 100, isResizable: true },
                          { key: 'col4', name: 'Purchase Date', fieldName: 'purchaseDate', minWidth: 100, maxWidth: 120, isResizable: true },
                          {
                            key: 'col5',
                            name: 'Warranty Expiry Date',
                            fieldName: 'warrantyExpiry',
                            minWidth: 140,
                            maxWidth: 200,
                            isResizable: true,
                            onRender: (item) => {
                              const isExpired = item.warrantyExpiry && new Date(item.warrantyExpiry) < new Date();
                              return (
                                <span style={{
                                  color: isExpired ? '#ef4444' : 'inherit',
                                  fontWeight: isExpired ? 'bold' : 'normal'
                                }}>
                                  {item.warrantyExpiry || 'N/A'} {isExpired && '(Expired)'}
                                </span>
                              );
                            }
                          }
                        ]}
                        setKey="warrantyReport"
                        layoutMode={DetailsListLayoutMode.justified}
                        selectionMode={SelectionMode.none}
                      />
                    </div>
                  </div>
                </PivotItem>
              )}
              {isAdmin && (
                <PivotItem headerText="Config" itemIcon="Settings" itemKey="Config">
                  <div style={{ marginTop: '20px' }}>
                    <div className={styles.cardHeader}>
                      <h3>Configuration</h3>
                    </div>
                    <p style={{ color: 'var(--text-muted)', marginBottom: '20px' }}>
                      Admin-only configuration area for list schema, process settings, and environment setup.
                    </p>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                      <div style={{ backgroundColor: 'var(--surface-color, #ffffff)', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
                        <h4 style={{ marginBottom: '15px', color: '#111827', marginTop: 0 }}>SharePoint List Connections</h4>
                        <ul style={{ listStyle: 'none', padding: 0, margin: 0, fontSize: '0.9rem' }}>
                          <li style={{ marginBottom: '12px', display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid #f3f4f6', paddingBottom: '8px' }}>
                            <span style={{ color: '#4b5563' }}>Inventory Database:</span>
                            <strong style={{ color: '#111827' }}>InventoryList</strong>
                          </li>
                          <li style={{ marginBottom: '12px', display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid #f3f4f6', paddingBottom: '8px' }}>
                            <span style={{ color: '#4b5563' }}>Approvals & Requests:</span>
                            <strong style={{ color: '#111827' }}>RequestList</strong>
                          </li>
                          <li style={{ display: 'flex', justifyContent: 'space-between' }}>
                            <span style={{ color: '#4b5563' }}>System Audit Logs:</span>
                            <strong style={{ color: '#111827' }}>AuditLogList</strong>
                          </li>
                        </ul>
                      </div>

                      <div style={{ backgroundColor: 'var(--surface-color, #ffffff)', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
                        <h4 style={{ marginBottom: '15px', color: '#111827', marginTop: 0 }}>Role Based Access Control</h4>
                        <ul style={{ listStyle: 'none', padding: 0, margin: 0, fontSize: '0.9rem' }}>
                          <li style={{ marginBottom: '12px', display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid #f3f4f6', paddingBottom: '8px' }}>
                            <span style={{ color: '#4b5563' }}>Admin Group:</span>
                            <strong style={{ color: '#111827' }}>Inventory Administrators</strong>
                          </li>
                          <li style={{ marginBottom: '12px', display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid #f3f4f6', paddingBottom: '8px' }}>
                            <span style={{ color: '#4b5563' }}>Manager Group:</span>
                            <strong style={{ color: '#111827' }}>Inventory Managers</strong>
                          </li>
                          <li style={{ display: 'flex', justifyContent: 'space-between' }}>
                            <span style={{ color: '#4b5563' }}>Employee Access:</span>
                            <strong style={{ color: '#111827' }}>Site Visitors</strong>
                          </li>
                        </ul>
                      </div>
                    </div>

                    <div style={{ marginTop: '20px', backgroundColor: 'var(--surface-color, #ffffff)', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
                      <h4 style={{ marginBottom: '5px', color: '#111827', marginTop: 0 }}>Required List Schema (Developer Reference)</h4>
                      <p style={{ fontSize: '0.85rem', color: '#6b7280', marginBottom: '15px', marginTop: 0 }}>Ensure your SharePoint lists contain the following columns exactly as written to prevent validation errors.</p>

                      <h5 style={{ marginTop: '15px', marginBottom: '8px', color: '#374151' }}>InventoryList <span style={{ fontWeight: 'normal', color: '#9ca3af' }}>(Asset Database)</span></h5>
                      <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', marginBottom: '25px' }}>
                        {['Title', 'AssetName', 'AssetType', 'SerialNumber', 'PurchaseDate', 'Status', 'Specifications', 'AssignedTo (Person/Group)'].map(col => (
                          <span key={col} style={{ backgroundColor: '#f3f4f6', padding: '4px 10px', borderRadius: '4px', fontSize: '0.75rem', color: '#374151', border: '1px solid #e5e7eb' }}>{col}</span>
                        ))}
                      </div>

                      <h5 style={{ marginBottom: '8px', color: '#374151' }}>RequestList <span style={{ fontWeight: 'normal', color: '#9ca3af' }}>(Approval Workflows)</span></h5>
                      <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                        {['Title', 'Employee', 'AssetType', 'Quantity', 'ReasonforRequest', 'RequestStatus', 'RequestKey', 'AssetStatus'].map(col => (
                          <span key={col} style={{ backgroundColor: '#f3f4f6', padding: '4px 10px', borderRadius: '4px', fontSize: '0.75rem', color: '#374151', border: '1px solid #e5e7eb' }}>{col}</span>
                        ))}
                      </div>
                    </div>
                  </div>
                </PivotItem>
              )}
            </Pivot>
          </div>

        </div>

        {(isAdmin || isManager) && (
          <AssetForm
            isOpen={isAssetFormOpen}
            onClose={() => this.setState({ isAssetFormOpen: false })}
            currentUserRole={effectiveRole}
            onAddAsset={this._onAddAsset}
          />
        )}

        {(isAdmin || isManager || isEmployee) && (
          <RequestForm
            isOpen={isRequestFormOpen}
            onClose={() => this.setState({ isRequestFormOpen: false })}
            availableAssets={items}
            employees={EMPLOYEES}
            currentUserRole={effectiveRole}
            currentUserName={userDisplayName}
            onSubmitRequest={this._onSubmitRequest}
          />
        )}
      </section>
    );
  }
}
