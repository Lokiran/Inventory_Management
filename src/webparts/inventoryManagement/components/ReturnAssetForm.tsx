import * as React from 'react';
import { useState } from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Stack, IStackTokens } from '@fluentui/react/lib/Stack';
import { IInventoryItem } from '../models/IInventoryItem';

export interface IReturnAssetFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  asset: IInventoryItem | undefined;
  onSubmit: (reason: string, condition: string) => Promise<void>;
}

const conditionOptions: IDropdownOption[] = [
  { key: 'Good', text: 'Good (No visible damage, fully functional)' },
  { key: 'Fair', text: 'Fair (Minor cosmetic wear, fully functional)' },
  { key: 'Poor', text: 'Poor (Significant wear or partial issues)' },
  { key: 'Damaged', text: 'Damaged (Broken, non-functional, physical damage)' }
];

const stackTokens: IStackTokens = { childrenGap: 15 };

export const ReturnAssetForm: React.FC<IReturnAssetFormProps> = (props) => {
  const { isOpen, onDismiss, asset, onSubmit } = props;
  
  const [reason, setReason] = useState<string>('');
  const [condition, setCondition] = useState<string>('Good');
  const [submitting, setSubmitting] = useState<boolean>(false);
  const [error, setError] = useState<string>('');

  if (!asset) return null;

  const handleSubmit = async (): Promise<void> => {
    if (!reason.trim()) {
      setError('Please provide a reason for the return.');
      return;
    }
    
    try {
      setSubmitting(true);
      setError('');
      await onSubmit(reason, condition);
      setReason('');
      setCondition('Good');
      onDismiss();
    } catch (err: any) {
      setError(err.message || 'Failed to submit return request.');
    } finally {
      setSubmitting(false);
    }
  };

  const onRenderFooterContent = (): JSX.Element => (
    <Stack horizontal tokens={{ childrenGap: 10 }}>
      <PrimaryButton 
        text="Submit Request" 
        onClick={handleSubmit} 
        disabled={submitting} 
      />
      <DefaultButton 
        text="Cancel" 
        onClick={onDismiss} 
        disabled={submitting} 
      />
    </Stack>
  );

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText="Request Asset Return"
      closeButtonAriaLabel="Close"
      onRenderFooterContent={onRenderFooterContent}
      isFooterAtBottom={true}
    >
      <div style={{ marginTop: '20px' }}>
        {error && (
          <div style={{ color: '#991b1b', backgroundColor: '#fee2e2', padding: '10px 15px', borderRadius: '4px', marginBottom: '15px', fontSize: '0.88rem' }}>
            <strong>Error:</strong> {error}
          </div>
        )}

        <Stack tokens={stackTokens}>
          <div style={{ backgroundColor: '#f3f4f6', padding: '15px', borderRadius: '6px' }}>
            <h4 style={{ margin: '0 0 10px 0', fontSize: '0.95rem', color: '#1f2937' }}>Asset Information</h4>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px', fontSize: '0.85rem' }}>
              <div><span style={{ color: '#6b7280' }}>Asset Name:</span> <strong>{asset.assetName || asset.title}</strong></div>
              <div><span style={{ color: '#6b7280' }}>Serial Number:</span> <strong>{asset.serialNumber || 'N/A'}</strong></div>
              <div><span style={{ color: '#6b7280' }}>Type:</span> <strong>{asset.assetType}</strong></div>
              <div><span style={{ color: '#6b7280' }}>Current Condition:</span> <strong>{asset.condition || 'Good'}</strong></div>
            </div>
          </div>

          <Dropdown
            label="Returned Asset Condition"
            selectedKey={condition}
            options={conditionOptions}
            onChange={(_, option) => setCondition(option ? (option.key as string) : 'Good')}
            required
          />

          <TextField
            label="Reason for Return"
            placeholder="Please detail why you are returning this asset (e.g. Upgrade received, hardware failure, contract ended...)"
            multiline
            rows={4}
            value={reason}
            onChange={(_, newValue) => setReason(newValue || '')}
            required
            errorMessage={reason.trim() ? '' : error && !reason.trim() ? 'Reason is required' : ''}
          />
        </Stack>
      </div>
    </Panel>
  );
};
