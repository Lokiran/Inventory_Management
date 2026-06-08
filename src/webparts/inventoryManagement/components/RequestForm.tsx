import * as React from 'react';
import {
  Panel,
  PanelType,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  Stack,
  IStackTokens,
  MessageBar,
  MessageBarType
} from '@fluentui/react';
import { IInventoryItem } from '../models/IInventoryItem';
import { IRequest } from '../models/IRequest';
import { IEmployee } from '../models/IEmployee';
import { RoleUtils, UserRole } from '../utils/RoleUtils';

export interface IRequestFormProps {
  isOpen: boolean;
  onClose: () => void;
  availableAssets: IInventoryItem[];
  employees: IEmployee[];
  currentUserRole: UserRole;
  currentUserName: string;
  onSubmitRequest: (request: Omit<IRequest, 'id' | 'requestKey' | 'requestDate' | 'status'>) => void;
}

const stackTokens: IStackTokens = { childrenGap: 15 };

export const RequestForm: React.FC<IRequestFormProps> = (props) => {
  const [selectedRequesterId, setSelectedRequesterId] = React.useState<string | undefined>(undefined);
  const [employeeId, setEmployeeId] = React.useState('');
  const [selectedAssetType, setSelectedAssetType] = React.useState<string | undefined>(undefined);
  const [priority, setPriority] = React.useState<'High' | 'Medium' | 'Low'>('Medium');
  const [quantity, setQuantity] = React.useState<number>(1);
  const [reason, setReason] = React.useState('');

  const isAdmin = props.currentUserRole === 'Admin';
  const isManager = props.currentUserRole === 'Inventory Manager';
  const isEmployee = props.currentUserRole === 'Inventory Employee';

  const currentUserOption: IEmployee = {
    id: 'current-user',
    name: props.currentUserName,
    email: '',
    department: 'Your Department',
    jobTitle: 'Employee'
  };

  const employeeExists = props.employees.some(emp => emp.name.toLowerCase() === props.currentUserName.toLowerCase());
  const allEmployees = employeeExists ? props.employees : [currentUserOption, ...props.employees];

  const availableEmployees = isEmployee
    ? allEmployees.filter(emp => emp.name.toLowerCase() === props.currentUserName.toLowerCase())
    : allEmployees;

  const employeeOptions: IDropdownOption[] = availableEmployees.map(emp => ({
    key: emp.id,
    text: `${emp.name} (${emp.department})`
  }));

  // Auto-select current user if employee
  React.useEffect(() => {
    if (isEmployee && employeeOptions.length === 1) {
      setSelectedRequesterId(employeeOptions[0].key as string);
      const emp = allEmployees.find(e => e.id === employeeOptions[0].key);
      if (emp) {
        setEmployeeId(emp.id);
      }
    }
  }, [isEmployee, employeeOptions, props.isOpen]);

  const uniqueAssetTypes = Array.from(new Set(props.availableAssets.map(a => a.assetType).filter(Boolean)));
  const dynamicAssetTypeOptions: IDropdownOption[] = uniqueAssetTypes.map(type => ({ key: type, text: type }));

  const assetTypeOptions: IDropdownOption[] = dynamicAssetTypeOptions.length > 0
    ? dynamicAssetTypeOptions
    : [
        { key: 'Laptop', text: 'Laptop' },
        { key: 'Monitor', text: 'Monitor' },
        { key: 'Mouse', text: 'Mouse' },
        { key: 'Keyboard', text: 'Keyboard' },
        { key: 'Headset', text: 'Headset' },
        { key: 'Other', text: 'Other' }
      ];

  const isFormValid = !!selectedRequesterId && !!employeeId.trim() && !!selectedAssetType && quantity > 0;

  const onSave = () => {
    // Validate role: employees can only request for themselves
    if (isEmployee && selectedRequesterId) {
      const selectedEmployee = allEmployees.find(e => e.id === selectedRequesterId);
      if (selectedEmployee && selectedEmployee.name.toLowerCase() !== props.currentUserName.toLowerCase()) {
        alert('Employees can only request assets for themselves.');
        return;
      }
    }

    const employee = allEmployees.find(e => e.id === selectedRequesterId);

    // Find a real asset ID to satisfy SharePoint backend lookups
    let matchingAsset = props.availableAssets.find(
      a => a.assetType === selectedAssetType && 
           (a.status === 'In Stock' || a.status === 'Yes')
    );
    if (!matchingAsset) {
      matchingAsset = props.availableAssets.find(a => a.assetType === selectedAssetType);
    }

    if (selectedAssetType && employee) {
      props.onSubmitRequest({
        requesterName: employee.name,
        employeeId: employeeId,
        assetId: matchingAsset ? matchingAsset.id : '1',
        assetTitle: selectedAssetType,
        priority: priority,
        quantity,
        reason
      } as any);

      setSelectedRequesterId(undefined);
      setEmployeeId('');
      setSelectedAssetType(undefined);
      setPriority('Medium');
      setQuantity(1);
      setReason('');
      props.onClose();
    }
  };

  return (
    <Panel
      isOpen={props.isOpen}
      onDismiss={props.onClose}
      type={PanelType.medium}
      headerText="Request an Asset"
      closeButtonAriaLabel="Close"
    >
      <Stack tokens={stackTokens}>
        {isEmployee && (
          <MessageBar messageBarType={MessageBarType.info}>
            You can only request assets for yourself. Contact your manager to request assets for others.
          </MessageBar>
        )}
        <Dropdown
          label={isEmployee ? "Requester (You)" : "Requester (Employee)"}
          selectedKey={selectedRequesterId}
          options={employeeOptions}
          onChange={(_, opt) => {
            const empId = opt?.key as string;
            setSelectedRequesterId(empId);
            const emp = allEmployees.find(e => e.id === empId);
            if (emp) {
              setEmployeeId(emp.id);
            }
          }}
          required
          disabled={isEmployee && employeeOptions.length === 1}
        />
        <TextField
          label="Employee ID"
          value={employeeId}
          onChange={(_, val) => setEmployeeId(val || '')}
          required
        />
        <Dropdown
          label="Asset Type"
          selectedKey={selectedAssetType}
          options={assetTypeOptions}
          onChange={(_, opt) => {
            setSelectedAssetType(opt?.key as string);
          }}
          required
        />
        <Dropdown
          label="Priority"
          selectedKey={priority}
          options={[
            { key: 'High', text: 'High' },
            { key: 'Medium', text: 'Medium' },
            { key: 'Low', text: 'Low' }
          ]}
          onChange={(_, opt) => setPriority(opt?.key as any)}
          required
        />
        <TextField
          label="Quantity"
          type="number"
          value={quantity.toString()}
          onChange={(_, val) => setQuantity(parseInt(val || '0'))}
          required
        />
        <TextField
          label="Reason for Request"
          multiline
          rows={3}
          value={reason}
          onChange={(_, val) => setReason(val || '')}
        />
        <Stack horizontal tokens={stackTokens} style={{ marginTop: 20 }}>
          <PrimaryButton text="Submit Request" onClick={onSave} disabled={!isFormValid} />
          <DefaultButton text="Cancel" onClick={props.onClose} />
        </Stack>
      </Stack>
    </Panel>
  );
};
