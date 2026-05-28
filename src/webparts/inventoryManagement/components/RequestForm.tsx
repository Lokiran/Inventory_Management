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
import { IRequest } from '../models/IRequest';
import { IEmployee } from '../models/IEmployee';

export interface IRequestFormProps {
  isOpen: boolean;
  onClose: () => void;
  availableAssets: IInventoryItem[];
  employees: IEmployee[];
  onSubmitRequest: (request: Omit<IRequest, 'id' | 'requestDate' | 'status'>) => void;
}

const stackTokens: IStackTokens = { childrenGap: 15 };

export const RequestForm: React.FC<IRequestFormProps> = (props) => {
  const [selectedRequesterId, setSelectedRequesterId] = React.useState<string | undefined>(undefined);
  const [selectedAssetId, setSelectedAssetId] = React.useState<string | undefined>(undefined);
  const [quantity, setQuantity] = React.useState<number>(1);
  const [reason, setReason] = React.useState('');

  const employeeOptions: IDropdownOption[] = props.employees.map(emp => ({
    key: emp.id,
    text: `${emp.name} (${emp.department})`
  }));

  const assetOptions: IDropdownOption[] = props.availableAssets.map(asset => ({
    key: asset.id,
    text: `${asset.assetName || asset.title || 'Unknown Asset'} (${asset.status || 'No Status'})`
  }));

  const onSave = () => {
    const asset = props.availableAssets.find(a => a.id === selectedAssetId);
    const employee = props.employees.find(e => e.id === selectedRequesterId);

    if (asset && employee) {
      props.onSubmitRequest({
        requesterName: employee.name,
        assetId: asset.id,
        assetTitle: asset.title,
        quantity,
        reason
      });
      setSelectedRequesterId(undefined);
      setSelectedAssetId(undefined);
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
        <Dropdown
          label="Requester (Employee)"
          selectedKey={selectedRequesterId}
          options={employeeOptions}
          onChange={(_, opt) => setSelectedRequesterId(opt?.key as string)}
          required
        />
        <Dropdown
          label="Select Asset"
          selectedKey={selectedAssetId}
          options={assetOptions}
          onChange={(_, opt) => setSelectedAssetId(opt?.key as string)}
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
          <PrimaryButton text="Submit Request" onClick={onSave} disabled={!selectedRequesterId || !selectedAssetId} />
          <DefaultButton text="Cancel" onClick={props.onClose} />
        </Stack>
      </Stack>
    </Panel>
  );
};
