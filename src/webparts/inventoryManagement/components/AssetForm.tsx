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
import { RoleUtils, UserRole } from '../utils/RoleUtils';

const assetTypeOptions: IDropdownOption[] = [
  { key: 'Laptop', text: 'Laptop' },
  { key: 'Monitor', text: 'Monitor' },
  { key: 'Mouse', text: 'Mouse' },
  { key: 'Keyboard', text: 'Keyboard' },
  { key: 'Headset', text: 'Headset' },
  { key: 'Other', text: 'Other' }
];

export interface IAssetFormProps {
  isOpen: boolean;
  onClose: () => void;
  currentUserRole: UserRole;
  onAddAsset: (asset: Omit<IInventoryItem, 'id' | 'status' | 'assignedTo'>) => void;
}

export const AssetForm: React.FC<IAssetFormProps> = (props) => {
  const stackTokens: IStackTokens = { childrenGap: 15 };
  const [title, setTitle] = React.useState('Company Assets');
  const [assetName, setAssetName] = React.useState('');
  const [assetType, setAssetType] = React.useState<string>('Laptop');
  const [serialNumber, setSerialNumber] = React.useState('');
  const [purchaseDate, setPurchaseDate] = React.useState(new Date().toISOString().split('T')[0]);
  const [note, setNote] = React.useState('');

  const isAdmin = props.currentUserRole === 'Admin';
  const isManager = props.currentUserRole === 'Inventory Manager';

  if (!RoleUtils.canAddAssets(props.currentUserRole)) {
    return null;
  }

  const onSave = () => {
    props.onAddAsset({
      title,
      assetName,
      assetType,
      serialNumber,
      purchaseDate,
      note
    });
    setAssetName('');
    setSerialNumber('');
    setNote('');
    props.onClose();
  };

  return (
    <Panel
      isOpen={props.isOpen}
      onDismiss={props.onClose}
      type={PanelType.medium}
      headerText={isAdmin ? "Add New Asset" : "Register New Asset"}
      closeButtonAriaLabel="Close"
    >
      <Stack tokens={stackTokens}>
        {isManager && !isAdmin && (
          <MessageBar messageBarType={MessageBarType.warning}>
            You are registering a new asset. After adding, you can assign it to employees in the Asset Tracking section.
          </MessageBar>
        )}
        <TextField
          label="Title (Group/Category)"
          value={title}
          onChange={(_, val) => setTitle(val || '')}
          required
        />
        <TextField
          label="Asset Name"
          value={assetName}
          onChange={(_, val) => setAssetName(val || '')}
          required
        />
        <Dropdown
          label="Asset Type"
          selectedKey={assetType}
          options={assetTypeOptions}
          onChange={(_, opt) => setAssetType(opt?.key as string || 'Other')}
          required
        />
        <TextField
          label="Serial Number"
          value={serialNumber}
          onChange={(_, val) => setSerialNumber(val || '')}
          required
        />
        <TextField
          label="Purchase Date"
          type="date"
          value={purchaseDate}
          onChange={(_, val) => setPurchaseDate(val || '')}
          required
        />
        <TextField
          label={isAdmin ? "Notes (Employee details, etc.)" : "Notes"}
          multiline
          rows={3}
          value={note}
          onChange={(_, val) => setNote(val || '')}
        />
        <Stack horizontal tokens={stackTokens} style={{ marginTop: 20 }}>
          <PrimaryButton text="Add Asset" onClick={onSave} disabled={!assetName || !serialNumber} />
          <DefaultButton text="Cancel" onClick={props.onClose} />
        </Stack>
      </Stack>
    </Panel>
  );
};
