import * as React from 'react';
import { IRequest } from '../models/IRequest';
import { 
  DetailsList, 
  DetailsListLayoutMode, 
  SelectionMode, 
  IColumn
} from '@fluentui/react/lib/DetailsList';
import { PrimaryButton } from '@fluentui/react';

export interface IRequestListProps {
  items: IRequest[];
  canApproveReject?: boolean;
  canApproveAsset?: boolean;
  showResponseColumns?: boolean;
  statusColumnLabel?: string;
  statusField?: 'status' | 'assetStatus';
  hideStatusColumn?: boolean;
  onApproveRequest?: (request: IRequest) => Promise<void>;
  onRejectRequest?: (request: IRequest, reason: string) => Promise<void>;
  onApproveAsset?: (request: IRequest) => Promise<void>;
  actionInProgressId?: string;
}

export const RequestList: React.FC<IRequestListProps> = (props) => {
  const columns: IColumn[] = [
    {
      key: 'columnRequestKey',
      name: 'Request ID',
      fieldName: 'requestKey',
      minWidth: 90,
      maxWidth: 125,
      isResizable: true
    },
    {
      key: 'columnEmployeeName',
      name: 'Employee Name',
      fieldName: 'requesterName',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true
    },
    {
      key: 'columnEmployeeId',
      name: 'Employee ID',
      fieldName: 'employeeId',
      minWidth: 90,
      maxWidth: 110,
      isResizable: true,
      onRender: (item: IRequest) => item.employeeId || '-'
    },
    {
      key: 'columnAssetType',
      name: 'Asset Type',
      fieldName: 'assetTitle',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true
    },
    {
      key: 'columnQuantity',
      name: 'Quantity',
      fieldName: 'quantity',
      minWidth: 60,
      maxWidth: 80,
      isResizable: true
    },
    {
      key: 'columnReason',
      name: 'Reason for Request',
      fieldName: 'reason',
      minWidth: 150,
      maxWidth: 250,
      isResizable: true,
      onRender: (item: IRequest) => item.reason || '-'
    },
    {
      key: 'columnPriority',
      name: 'Priority',
      fieldName: 'priority',
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IRequest) => {
        const priority = item.priority || 'Medium';
        let color = '#4b5563'; // default medium (gray)
        let backgroundColor = '#f3f4f6';
        if (priority === 'High') {
          color = '#b91c1c';
          backgroundColor = '#fee2e2';
        } else if (priority === 'Low') {
          color = '#1e3a8a';
          backgroundColor = '#dbeafe';
        }
        return (
          <span style={{
            backgroundColor,
            color,
            padding: '4px 10px',
            borderRadius: '9999px',
            fontSize: '0.75rem',
            fontWeight: 600
          }}>
            {priority}
          </span>
        );
      }
    },
    {
      key: 'columnRequestDate',
      name: 'Request Date',
      fieldName: 'requestDate',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true
    },
    ...(props.hideStatusColumn ? [] : [{ 
      key: 'column6', 
      name: props.statusColumnLabel || 'Status', 
      fieldName: props.statusField || 'status', 
      minWidth: 80, 
      maxWidth: 100, 
      isResizable: true,
      onRender: (item: IRequest) => {
        let val: string = item[props.statusField || 'status'] || 'Pending';
        if (val === 'Pending') val = 'Pending';

        let backgroundColor = '#fef3c7'; // default pending (yellow)
        let textColor = '#92400e';
        
        if (val === 'Approved') {
          backgroundColor = '#dcfce7';
          textColor = '#166534';
        } else if (val === 'Declined') {
          backgroundColor = '#fee2e2';
          textColor = '#991b1b';
        }
        
        return (
          <span style={{ 
            backgroundColor, 
            color: textColor, 
            padding: '4px 12px', 
            borderRadius: '9999px', 
            fontSize: '0.75rem', 
            fontWeight: 600,
            display: 'inline-block'
          }}>
            {val}
          </span>
        );
      }
    } as IColumn]),
    ...(props.canApproveAsset ? [{
      key: 'columnAssetStatus',
      name: 'Asset Status',
      fieldName: 'assetStatus',
      minWidth: 140,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: IRequest) => {
        const value = item.assetStatus || 'Pending';
        const isApproved = value.toLowerCase().includes('approv');
        const isBusy = props.actionInProgressId === item.id;

        return (
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{
              backgroundColor: isApproved ? '#dcfce7' : '#fef3c7',
              color: isApproved ? '#166534' : '#92400e',
              padding: '4px 10px',
              borderRadius: '9999px',
              fontSize: '0.75rem',
              fontWeight: 600,
              display: 'inline-block'
            }}>
              {value}
            </span>
            {!isApproved && (
              <PrimaryButton
                text="Approve"
                onClick={() => props.onApproveAsset && props.onApproveAsset(item)}
                disabled={isBusy}
              />
            )}
          </div>
        );
      }
    } as IColumn] : []),
    ...(props.showResponseColumns ? [
      { key: 'columnManagerResponse', name: 'Manager Response', fieldName: 'managerResponse', minWidth: 170, maxWidth: 240, isResizable: true },
      {
        key: 'columnAdminResponse',
        name: 'Admin Response',
        fieldName: 'assetStatus',
        minWidth: 140,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: IRequest) => {
          const managerStatus = (item.status || '').toLowerCase();
          if (managerStatus === 'pending') {
            return <span style={{ color: '#92400e', fontStyle: 'italic' }}>Waiting on Manager</span>;
          }
          if (managerStatus === 'declined' || managerStatus === 'rejected') {
            return <span style={{ color: '#991b1b', fontStyle: 'italic' }}>N/A (Rejected)</span>;
          }
          const isApproved = (item.assetStatus || '').toLowerCase().includes('approv');
          return isApproved ? (
            <span style={{ color: '#166534', fontWeight: 600 }}>Asset Allocated</span>
          ) : (
            <span style={{ color: '#92400e', fontWeight: 600 }}>Pending Admin Approval</span>
          );
        }
      }
    ] as IColumn[] : []),
    ...(props.canApproveReject ? [{
      key: 'column8',
      name: 'Actions',
      fieldName: 'actions',
      minWidth: 220,
      maxWidth: 260,
      isResizable: true,
      onRender: (item: IRequest) => {
        const isPending = (item.status || '').toLowerCase() === 'pending';
        const isBusy = props.actionInProgressId === item.id;

        if (!isPending) {
          return <span style={{ color: 'var(--text-muted)' }}>No action</span>;
        }

        return (
          <div style={{ display: 'flex', gap: '8px' }}>
            <PrimaryButton
              text="Approve"
              onClick={() => props.onApproveRequest && props.onApproveRequest(item)}
              disabled={isBusy}
            />
            <PrimaryButton
              text="Reject"
              onClick={() => {
                if (!props.onRejectRequest) {
                  return;
                }

                const rejectionReason = window.prompt('Enter rejection reason for this request:');
                if (!rejectionReason || !rejectionReason.trim()) {
                  return;
                }

                void props.onRejectRequest(item, rejectionReason.trim());
              }}
              disabled={isBusy}
              styles={{
                root: { backgroundColor: '#991b1b', borderColor: '#991b1b' },
                rootHovered: { backgroundColor: '#7f1d1d', borderColor: '#7f1d1d' }
              }}
            />
          </div>
        );
      }
    } as IColumn] : [])
  ];

  return (
    <div style={{ marginTop: '10px' }}>
      {props.items.length === 0 ? (
        <p style={{ fontStyle: 'italic', color: 'var(--text-muted)' }}>No asset requests found.</p>
      ) : (
        <DetailsList
          items={props.items}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
        />
      )}
    </div>
  );
};
