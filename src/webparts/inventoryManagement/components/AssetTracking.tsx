import * as React from 'react';
import {
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Stack,
  IStackTokens,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { IInventoryItem } from '../models/IInventoryItem';
import { IEmployee } from '../models/IEmployee';
import { RoleUtils, UserRole } from '../utils/RoleUtils';

export interface IAssetTrackingProps {
  items: IInventoryItem[];
  employees: IEmployee[];
  currentUserRole: UserRole;
  onAssignAssets: (employeeName: string, employeeEmail: string, assetIds: string[]) => Promise<void>;
  isActionInProgress: boolean;
}

export const AssetTracking: React.FC<IAssetTrackingProps> = (props) => {
  const [selectedEmployeeId, setSelectedEmployeeId] = React.useState<string | undefined>();
  const [selectedAssetIds, setSelectedAssetIds] = React.useState<string[]>([]);

  const stackTokens: IStackTokens = { childrenGap: 20 };

  if (!RoleUtils.canAssignAssets(props.currentUserRole)) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>
        You do not have permission to access the Asset Tracking section. Only Admins and Managers can assign assets.
      </MessageBar>
    );
  }

  const employeeOptions: IDropdownOption[] = props.employees.map(emp => ({
    key: emp.id,
    text: `${emp.name} (${emp.department})`
  }));

  const availableAssets = props.items.filter(item => item.status === 'In Stock' || item.status === 'Yes');
  const assetOptions: IDropdownOption[] = availableAssets.map(asset => ({
    key: asset.id,
    text: `${asset.assetName || asset.title} (${asset.serialNumber || 'N/A'}) - ${asset.assetType || 'Unknown'}`
  }));

  const selectedEmployee = props.employees.find(e => e.id === selectedEmployeeId);
  const normalize = (val: string | undefined) => (val || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const employeeAssignedAssets = props.items.filter(i => {
    if (!selectedEmployee) return false;
    const nameNorm = normalize(selectedEmployee.name);
    const assignedNorm = normalize(i.assignedTo);
    const isAssigned = assignedNorm && (assignedNorm === nameNorm || assignedNorm.includes(nameNorm) || nameNorm.includes(assignedNorm));
    const isNoted = (i.note || '').toLowerCase().includes('assigned to:') && normalize(i.note).includes(nameNorm);
    return isAssigned || isNoted;
  });

  const handleAssign = async () => {
    if (selectedEmployee && selectedAssetIds.length > 0) {
      await props.onAssignAssets(selectedEmployee.name, selectedEmployee.email, selectedAssetIds);
      setSelectedAssetIds([]);
    }
  };

  const handleAssetSelect = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      setSelectedAssetIds(
        option.selected
          ? [...selectedAssetIds, option.key as string]
          : selectedAssetIds.filter(key => key !== option.key)
      );
    }
  };

  return (
    <div style={{ marginTop: '20px' }}>
      <div style={{ backgroundColor: 'var(--surface-color, #ffffff)', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)', marginBottom: '20px' }}>
        <h4 style={{ marginTop: 0, marginBottom: '15px', color: '#111827' }}>Select Employee</h4>
        <Dropdown
          placeholder="Select an employee to view or assign assets"
          options={employeeOptions}
          selectedKey={selectedEmployeeId}
          onChange={(_, option) => setSelectedEmployeeId(option?.key as string)}
          styles={{ dropdown: { width: 400 } }}
        />
      </div>

      {selectedEmployee && (
        <Stack tokens={stackTokens}>
          {/* Currently Assigned Assets Section */}
          <div style={{ backgroundColor: 'var(--surface-color, #ffffff)', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.1)' }}>
            <h4 style={{ marginTop: 0, marginBottom: '15px', color: '#111827' }}>Currently Assigned to {selectedEmployee.name}</h4>
            {employeeAssignedAssets.length > 0 ? (
              <DetailsList
                items={employeeAssignedAssets}
                columns={[
                  { key: 'col1', name: 'Asset Name', fieldName: 'assetName', minWidth: 150, maxWidth: 200, isResizable: true, onRender: item => item.assetName || item.title },
                  { key: 'col2', name: 'Type', fieldName: 'assetType', minWidth: 100, maxWidth: 150, isResizable: true },
                  { key: 'col3', name: 'Serial Number', fieldName: 'serialNumber', minWidth: 120, maxWidth: 180, isResizable: true },
                  { key: 'col4', name: 'Status', fieldName: 'status', minWidth: 100, maxWidth: 120, isResizable: true },
                  { 
                    key: 'col5', 
                    name: 'Assigned Date', 
                    fieldName: 'assignedDate', 
                    minWidth: 120, 
                    maxWidth: 150, 
                    isResizable: true,
                    onRender: item => {
                      if (!item.assignedDate) return 'N/A';
                      try {
                        return new Date(item.assignedDate).toLocaleString();
                      } catch {
                        return item.assignedDate;
                      }
                    }
                  }
                ]}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
              />
            ) : (
              <p style={{ color: '#6b7280', fontSize: '0.9rem' }}>This employee has no assets currently assigned.</p>
            )}
          </div>

          {/* Assign New Assets Section */}
          <div style={{ backgroundColor: '#f0f6ff', padding: '20px', borderRadius: '8px', border: '1px solid #cbe0f5' }}>
            <h4 style={{ marginTop: 0, marginBottom: '15px', color: '#0078d4' }}>Assign New Assets</h4>
            <div style={{ display: 'flex', alignItems: 'flex-end', gap: '15px' }}>
              <Dropdown
                label="Select Assets from Inventory (Multiple allowed)"
                multiSelect
                options={assetOptions}
                selectedKeys={selectedAssetIds}
                onChange={handleAssetSelect}
                placeholder={availableAssets.length > 0 ? "Select assets to assign..." : "No assets available in stock"}
                disabled={availableAssets.length === 0 || props.isActionInProgress}
                styles={{ dropdown: { width: 500 } }}
              />
              <PrimaryButton
                text={props.isActionInProgress ? "Assigning..." : "Assign Selected Assets"}
                onClick={handleAssign}
                disabled={selectedAssetIds.length === 0 || props.isActionInProgress}
                iconProps={{ iconName: props.isActionInProgress ? 'Sync' : 'Add' }}
              />
            </div>
            <p style={{ marginTop: '10px', fontSize: '0.85rem', color: '#323130' }}>
              Selected assets will be updated in the Inventory database as 'Assigned' to {selectedEmployee.name}.
            </p>
          </div>
        </Stack>
      )}
    </div>
  );
};
