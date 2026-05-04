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
  IStackTokens
} from '@fluentui/react';
import { IInventoryItem } from '../models/IInventoryItem';

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
  onAddAsset: (asset: Omit<IInventoryItem, 'id' | 'status' | 'assignedTo'>) => void;
}

export const AssetForm: React.FC<IAssetFormProps> = (props) => {
  const stackTokens: IStackTokens = { childrenGap: 15 };
  const [title, setTitle] = React.useState('Company Assets');
  const [assetName, setAssetName] = React.useState('');
  const [assetType, setAssetType] = React.useState<string>('Laptop');
  const [serialNumber, setSerialNumber] = React.useState('');
  const [purchaseDate, setPurchaseDate] = React.useState(new Date().toISOString().split('T')[0]);

  const onSave = () => {
    props.onAddAsset({
      title,
      assetName,
      assetType,
      serialNumber,
      purchaseDate
    });
    setAssetName('');
    setSerialNumber('');
    props.onClose();
  };

  return (
    <Panel
      isOpen={props.isOpen}
      onDismiss={props.onClose}
      type={PanelType.medium}
      headerText="Add New Asset"
      closeButtonAriaLabel="Close"
    >
      <Stack tokens={stackTokens}>
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
        <Stack horizontal tokens={stackTokens} style={{ marginTop: 20 }}>
          <PrimaryButton text="Add Asset" onClick={onSave} disabled={!assetName || !serialNumber} />
          <DefaultButton text="Cancel" onClick={props.onClose} />
        </Stack>
      </Stack>
    </Panel>
  );
};
