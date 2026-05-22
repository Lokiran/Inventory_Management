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
      minWidth: 60,
      maxWidth: 80,
      isResizable: true,
      onRender: (item: IEventLog) => {
        let backgroundColor = '#e5e7eb';
        let textColor = '#374151';

        if (item.action === 'Create') {
          backgroundColor = '#dcfce7';
          textColor = '#166534';
        } else if (item.action === 'Update') {
          backgroundColor = '#fef3c7';
          textColor = '#92400e';
        } else if (item.action === 'Delete') {
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
            {item.action}
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
        <h3 style={{ margin: '0 0 10px 0', fontSize: '1.25rem', fontWeight: 600 }}>Audit Event Stream</h3>
        <p style={{ color: 'var(--text-muted)', margin: '0 0 15px 0' }}>
          {isEmployee
            ? 'Track your own asset and request operations.'
            : 'Track and search all asset and request operations.'}
        </p>

        {isEmployee && (
          <MessageBar messageBarType={MessageBarType.info}>
            You can only view audit logs related to your own actions. Contact your manager for other employees' activity logs.
          </MessageBar>
        )}

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
