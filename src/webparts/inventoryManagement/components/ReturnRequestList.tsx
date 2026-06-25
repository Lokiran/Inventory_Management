import * as React from 'react';
import { useState, useMemo } from 'react';
import { 
  DetailsList, 
  DetailsListLayoutMode, 
  SelectionMode, 
  IColumn 
} from '@fluentui/react/lib/DetailsList';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { TextField } from '@fluentui/react/lib/TextField';
import { Stack } from '@fluentui/react/lib/Stack';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { IReturnRequest } from '../models/IReturnRequest';

export interface IReturnRequestListProps {
  items: IReturnRequest[];
  isAdmin: boolean;
  isManager: boolean;
  onUpdateStatus: (requestId: string, status: 'Approved' | 'Rejected' | 'Completed', comment: string, finalCondition?: string) => Promise<void>;
  loading: boolean;
}

const conditionOptions: IDropdownOption[] = [
  { key: 'Good', text: 'Good (No damage, fully functional)' },
  { key: 'Fair', text: 'Fair (Minor wear, fully functional)' },
  { key: 'Poor', text: 'Poor (Significant wear, needs repair)' },
  { key: 'Damaged', text: 'Damaged (Broken, non-functional)' }
];

export const ReturnRequestList: React.FC<IReturnRequestListProps> = (props) => {
  const { items, isAdmin, isManager, onUpdateStatus, loading } = props;
  
  const [searchQuery, setSearchQuery] = useState<string>('');
  
  // Dialog / State for Actions
  const [activeRequest, setActiveRequest] = useState<IReturnRequest | null>(null);
  const [actionType, setActionType] = useState<'Approve' | 'Reject' | 'Complete' | null>(null);
  const [comment, setComment] = useState<string>('');
  const [finalCondition, setFinalCondition] = useState<string>('Good');
  const [submitting, setSubmitting] = useState<boolean>(false);

  // Search filter
  const filteredItems = useMemo(() => {
    if (!searchQuery) return items;
    const q = searchQuery.toLowerCase();
    return items.filter(item => 
      (item.assetName || '').toLowerCase().includes(q) ||
      (item.serialNumber || '').toLowerCase().includes(q) ||
      (item.requesterName || '').toLowerCase().includes(q) ||
      (item.status || '').toLowerCase().includes(q) ||
      (item.returnReason || '').toLowerCase().includes(q)
    );
  }, [items, searchQuery]);

  const openDialog = (request: IReturnRequest, type: 'Approve' | 'Reject' | 'Complete'): void => {
    setActiveRequest(request);
    setActionType(type);
    setComment('');
    setFinalCondition(request.proposedCondition || 'Good');
  };

  const closeDialog = (): void => {
    setActiveRequest(null);
    setActionType(null);
    setComment('');
    setSubmitting(false);
  };

  const handleAction = async (): Promise<void> => {
    if (!activeRequest || !actionType) return;
    
    if (actionType === 'Reject' && !comment.trim()) {
      alert('Please provide a reason/comment for rejection.');
      return;
    }

    try {
      setSubmitting(true);
      if (actionType === 'Approve') {
        await onUpdateStatus(activeRequest.id, 'Approved', comment || 'Approved by Manager');
      } else if (actionType === 'Reject') {
        await onUpdateStatus(activeRequest.id, 'Rejected', comment);
      } else if (actionType === 'Complete') {
        await onUpdateStatus(activeRequest.id, 'Completed', comment || 'Checked in & Verified', finalCondition);
      }
      closeDialog();
    } catch (e: any) {
      alert('Action failed: ' + (e.message || JSON.stringify(e)));
    } finally {
      setSubmitting(false);
    }
  };

  const getStatusStyles = (status: string): { bg: string; fg: string } => {
    switch (status) {
      case 'Pending':
        return { bg: '#ffedd5', fg: '#9a3412' }; // Light orange
      case 'Approved':
        return { bg: '#dbeafe', fg: '#1e40af' }; // Light blue
      case 'Rejected':
        return { bg: '#fee2e2', fg: '#991b1b' }; // Light red
      case 'Completed':
        return { bg: '#dcfce7', fg: '#166534' }; // Light green
      default:
        return { bg: '#f3f4f6', fg: '#374151' };
    }
  };

  const columns: IColumn[] = [
    { key: 'id', name: 'ID', fieldName: 'id', minWidth: 50, maxWidth: 80, isResizable: true, onRender: (item: IReturnRequest) => item.id.replace('RR-', '#') },
    { key: 'assetName', name: 'Asset Name', fieldName: 'assetName', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'serialNumber', name: 'Serial Number', fieldName: 'serialNumber', minWidth: 90, maxWidth: 120, isResizable: true },
    { key: 'requesterName', name: 'Employee', fieldName: 'requesterName', minWidth: 100, maxWidth: 130, isResizable: true },
    { key: 'returnReason', name: 'Reason', fieldName: 'returnReason', minWidth: 150, maxWidth: 220, isResizable: true, isMultiline: true },
    { key: 'proposedCondition', name: 'Condition', fieldName: 'proposedCondition', minWidth: 80, maxWidth: 110, isResizable: true },
    { 
      key: 'status', 
      name: 'Status', 
      fieldName: 'status', 
      minWidth: 90, 
      maxWidth: 110, 
      isResizable: true,
      onRender: (item: IReturnRequest) => {
        const { bg, fg } = getStatusStyles(item.status);
        return (
          <span style={{ 
            backgroundColor: bg, 
            color: fg, 
            padding: '4px 10px', 
            borderRadius: '9999px', 
            fontSize: '0.75rem', 
            fontWeight: 600,
            display: 'inline-block'
          }}>
            {item.status}
          </span>
        );
      }
    },
    { key: 'requestDate', name: 'Requested Date', fieldName: 'requestDate', minWidth: 90, maxWidth: 120, isResizable: true },
    { 
      key: 'managerComment', 
      name: 'Manager Notes', 
      fieldName: 'managerComment', 
      minWidth: 120, 
      maxWidth: 200, 
      isResizable: true,
      onRender: (item: IReturnRequest) => item.managerComment || <span style={{ color: '#9ca3af', fontStyle: 'italic' }}>None</span>
    },
    // Actions Column (Visible to managers/admins)
    ...((isManager || isAdmin) ? [
      {
        key: 'actions',
        name: 'Actions',
        minWidth: 200,
        maxWidth: 260,
        isResizable: true,
        onRender: (item: IReturnRequest) => {
          if (item.status === 'Completed') {
            return <span style={{ color: '#16a34a', fontWeight: 600, fontSize: '0.8rem' }}>Check-in Complete</span>;
          }
          if (item.status === 'Rejected') {
            return <span style={{ color: '#ef4444', fontWeight: 600, fontSize: '0.8rem' }}>Rejected</span>;
          }
          
          return (
            <Stack horizontal tokens={{ childrenGap: 6 }}>
              {item.status === 'Pending' && (
                <>
                  <PrimaryButton
                    text="Approve"
                    onClick={() => openDialog(item, 'Approve')}
                    styles={{ root: { height: 26, padding: '4px 8px', fontSize: '0.75rem' } }}
                  />
                  <DefaultButton
                    text="Reject"
                    onClick={() => openDialog(item, 'Reject')}
                    styles={{ root: { height: 26, padding: '4px 8px', fontSize: '0.75rem', color: '#b91c1c', borderColor: '#fee2e2' } }}
                  />
                </>
              )}
              {item.status === 'Approved' && (
                <PrimaryButton
                  text="Verify & Complete"
                  onClick={() => openDialog(item, 'Complete')}
                  styles={{ root: { height: 26, padding: '4px 8px', fontSize: '0.75rem', backgroundColor: '#107c41', borderColor: '#107c41' } }}
                />
              )}
            </Stack>
          );
        }
      }
    ] : [])
  ];

  return (
    <div style={{ marginTop: '10px' }}>
      <div style={{ marginBottom: '15px' }}>
        <SearchBox
          placeholder="Search return requests..."
          value={searchQuery}
          onChange={(_, val) => setSearchQuery(val || '')}
          styles={{ root: { maxWidth: 350 } }}
        />
      </div>

      {loading ? (
        <p>Loading return requests...</p>
      ) : filteredItems.length === 0 ? (
        <p style={{ fontStyle: 'italic', color: '#6b7280' }}>No return requests found.</p>
      ) : (
        <DetailsList
          items={filteredItems}
          columns={columns}
          setKey="returnSet"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
        />
      )}

      {/* Confirmation Dialog */}
      <Dialog
        hidden={!activeRequest}
        onDismiss={closeDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: actionType === 'Approve' ? 'Approve Return Request' :
                 actionType === 'Reject' ? 'Reject Return Request' : 'Verify & Complete Return',
          subText: activeRequest ? `Request by ${activeRequest.requesterName} for asset ${activeRequest.assetName}` : ''
        }}
        modalProps={{ isBlocking: true }}
      >
        <Stack tokens={{ childrenGap: 15 }} style={{ marginTop: '15px' }}>
          {actionType === 'Complete' && (
            <Dropdown
              label="Final Verified Condition"
              selectedKey={finalCondition}
              options={conditionOptions}
              onChange={(_, opt) => setFinalCondition(opt ? (opt.key as string) : 'Good')}
            />
          )}

          <TextField
            label={actionType === 'Reject' ? 'Rejection Reason (Required)' : 'Manager Comments / Verification Notes'}
            placeholder={actionType === 'Reject' ? 'Please specify why this return request is being rejected...' : 'Add check-in details (e.g. checked power cords, verified serial number...)'}
            multiline
            rows={3}
            value={comment}
            onChange={(_, val) => setComment(val || '')}
            required={actionType === 'Reject'}
          />
        </Stack>

        <DialogFooter>
          <PrimaryButton 
            text={actionType === 'Approve' ? 'Approve' :
                  actionType === 'Reject' ? 'Reject' : 'Verify & Complete'} 
            onClick={handleAction} 
            disabled={submitting || (actionType === 'Reject' && !comment.trim())} 
          />
          <DefaultButton text="Cancel" onClick={closeDialog} disabled={submitting} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
