import * as React from 'react';
import { useState, useMemo } from 'react';
import { IEventLog } from '../models/IEventLog';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn
} from '@fluentui/react/lib/DetailsList';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { RoleUtils, UserRole } from '../utils/RoleUtils';

export interface IEventStreamProps {
  logs: IEventLog[];
  loading: boolean;
  errorMessage?: string;
  currentUserRole: UserRole;
  currentUserName: string;
}

export const EventStream: React.FC<IEventStreamProps> = (props) => {
  const [searchQuery, setSearchQuery] = useState<string>('');

  const isAdmin = props.currentUserRole === 'Admin';
  const isManager = props.currentUserRole === 'Inventory Manager';
  const isEmployee = props.currentUserRole === 'Inventory Employee';

  const columns: IColumn[] = [
    {
      key: 'column_action',
      name: 'Action',
      fieldName: 'action',
      minWidth: 120,
      maxWidth: 220,
      isResizable: true,
      onRender: (item: IEventLog) => {
        let backgroundColor = '#f3f4f6';
        let textColor = '#374151';
        let displayText = item.action || '';

        const normalizedAction = displayText.toLowerCase().trim();

        if (normalizedAction === 'created' || normalizedAction === 'create') {
          backgroundColor = '#dbeafe'; // Light blue
          textColor = '#1e40af';      // Dark blue
          displayText = 'created';
        } else if (normalizedAction === 'manager approved') {
          backgroundColor = '#dcfce7'; // Light green
          textColor = '#166534';      // Dark green
          displayText = 'manager approved';
        } else if (normalizedAction === 'manager rejected') {
          backgroundColor = '#fee2e2'; // Light red
          textColor = '#991b1b';      // Dark red
          displayText = 'manager rejected';
        } else if (normalizedAction === 'admin assigned') {
          backgroundColor = '#f3e8ff'; // Light purple
          textColor = '#6b21a8';      // Dark purple
          displayText = 'admin assigned';
        } else if (normalizedAction === 'status updated to in progress') {
          backgroundColor = '#ffedd5'; // Light orange/yellow
          textColor = '#9a3412';      // Dark orange
          displayText = 'status updated to in progress';
        } else if (normalizedAction === 'status updated to resolved') {
          backgroundColor = '#ccfbf1'; // Light teal
          textColor = '#115e59';      // Dark teal
          displayText = 'status updated to resolved';
        } else if (normalizedAction === 'deleted' || normalizedAction === 'delete') {
          backgroundColor = '#fee2e2'; // Light red
          textColor = '#991b1b';      // Dark red
          displayText = 'deleted';
        } else if (normalizedAction === 'return requested') {
          backgroundColor = '#ffedd5'; // Light orange
          textColor = '#9a3412';      // Dark orange
          displayText = 'return requested';
        } else if (normalizedAction === 'return approved') {
          backgroundColor = '#dcfce7'; // Light green
          textColor = '#166534';      // Dark green
          displayText = 'return approved';
        } else if (normalizedAction === 'return completed') {
          backgroundColor = '#ccfbf1'; // Light teal
          textColor = '#115e59';      // Dark teal
          displayText = 'return completed';
        } else if (normalizedAction === 'return rejected') {
          backgroundColor = '#fee2e2'; // Light red
          textColor = '#991b1b';      // Dark red
          displayText = 'return rejected';
        } else if (normalizedAction === 'activated') {
          backgroundColor = '#dcfce7'; // Light green
          textColor = '#166534';      // Dark green
          displayText = 'activated';
        } else if (normalizedAction === 'inactivated') {
          backgroundColor = '#fef3c7'; // Light amber
          textColor = '#92400e';      // Dark amber
          displayText = 'inactivated';
        } else if (normalizedAction === 'deactivated') {
          backgroundColor = '#fee2e2'; // Light red
          textColor = '#991b1b';      // Dark red
          displayText = 'deactivated';
        } else if (normalizedAction === 'update') {
          backgroundColor = '#ffedd5'; // Light orange/yellow (fallback for generic Update)
          textColor = '#9a3412';
          displayText = 'updated';
        }

        return (
          <span style={{
            backgroundColor,
            color: textColor,
            padding: '4px 12px',
            borderRadius: '9999px',
            fontSize: '0.75rem',
            fontWeight: 600,
            display: 'inline-block',
            textTransform: 'lowercase'
          }}>
            {displayText}
          </span>
        );
      }
    },
    { key: 'column_type', name: 'Type', fieldName: 'entityType', minWidth: 60, maxWidth: 80, isResizable: true },
    { key: 'column_title', name: 'Title', fieldName: 'title', minWidth: 150, maxWidth: 200, isResizable: true },
    { key: 'column_assetName', name: 'Asset Name', fieldName: 'assetName', minWidth: 100, maxWidth: 150, isResizable: true },
    ...(RoleUtils.canViewAuditLogs(props.currentUserRole) ? [
      { key: 'column_user', name: 'User', fieldName: 'user', minWidth: 100, maxWidth: 150, isResizable: true }
    ] : []),
    { key: 'column_timestamp', name: 'Timestamp', fieldName: 'timestamp', minWidth: 120, maxWidth: 160, isResizable: true },
    ...(RoleUtils.canViewAuditLogs(props.currentUserRole) ? [
      { key: 'column_details', name: 'Details', fieldName: 'details', minWidth: 200, maxWidth: 400, isResizable: true, isMultiline: true }
    ] : [])
  ];

  const roleBasedFilteredLogs = useMemo(() => {
    if (isEmployee) {
      return props.logs.filter(log =>
        (log.user || '').toLowerCase().includes(props.currentUserName.toLowerCase()) ||
        (log.details || '').toLowerCase().includes(props.currentUserName.toLowerCase())
      );
    }
    return props.logs;
  }, [props.logs, isEmployee, props.currentUserName]);

  const filteredLogs = useMemo(() => {
    if (!searchQuery) {
      return roleBasedFilteredLogs;
    }
    const lowerQuery = searchQuery.toLowerCase();
    return roleBasedFilteredLogs.filter(log =>
      log.title?.toLowerCase().includes(lowerQuery) ||
      log.assetName?.toLowerCase().includes(lowerQuery) ||
      log.details?.toLowerCase().includes(lowerQuery) ||
      log.user?.toLowerCase().includes(lowerQuery) ||
      log.action?.toLowerCase().includes(lowerQuery) ||
      log.entityType?.toLowerCase().includes(lowerQuery) ||
      log.entityId?.toLowerCase().includes(lowerQuery)
    );
  }, [roleBasedFilteredLogs, searchQuery]);

  return (
    <div style={{ marginTop: '20px' }}>
      <div style={{ marginBottom: '20px' }}>
        {props.errorMessage && (
          <div style={{ color: '#991b1b', backgroundColor: '#fee2e2', padding: '15px', borderRadius: '8px', marginBottom: '15px' }}>
            <strong>Notice:</strong> {props.errorMessage}
          </div>
        )}

        <SearchBox
          placeholder="Search by asset name, title, details, user, or action..."
          value={searchQuery}
          onChange={(_, newValue) => setSearchQuery(newValue || '')}
          onClear={() => setSearchQuery('')}
          styles={{ root: { maxWidth: 400 } }}
        />
      </div>

      {props.loading ? (
        <p>Loading audit logs...</p>
      ) : roleBasedFilteredLogs.length === 0 ? (
        <p style={{ fontStyle: 'italic', color: 'var(--text-muted)' }}>No audit events {isEmployee ? 'for you' : ''} recorded yet.</p>
      ) : filteredLogs.length === 0 ? (
        <p style={{ fontStyle: 'italic', color: 'var(--text-muted)' }}>No audit events match your search query.</p>
      ) : (
        <DetailsList
          items={filteredLogs}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
        />
      )}
    </div>
  );
};
