import * as React from 'react';
import { useState } from 'react';
import {
  Stack,
  Text,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { Card, ICardTokens } from '@uifabric/react-cards';
import styles from './AssetRequestModule.module.scss';
import { IEmployeeManagementProps } from '../IEmployeeManagementProps';
import { InventoryService } from '../../services/InventoryService';

interface IAssetRequestForm {
  employeeEmail: string;
  employeeName: string;
  department: string;
  assetType: string;
  assetName: string;
  quantity: number;
  priority: string;
  reasonDescription: string;
  requiredDate: string;
}

const assetTypes: IDropdownOption[] = [
  { key: 'laptop', text: 'Laptop' },
  { key: 'mouse', text: 'Mouse' },
  { key: 'keyboard', text: 'Keyboard' },
  { key: 'monitor', text: 'Monitor' },
  { key: 'headset', text: 'Headset' },
  { key: 'mobile', text: 'Mobile' },
  { key: 'other', text: 'Other' },
];

const priorityOptions: IDropdownOption[] = [
  { key: 'low', text: 'Low' },
  { key: 'medium', text: 'Medium' },
  { key: 'high', text: 'High' },
  { key: 'urgent', text: 'Urgent' },
];

const cardTokens: ICardTokens = { childrenMargin: 12 };

export const AssetRequestModule: React.FC<IEmployeeManagementProps & { setIsLoading: (loading: boolean) => void }> = (props) => {
  const [formData, setFormData] = useState<IAssetRequestForm>({
    employeeEmail: '',
    employeeName: props.userName || '',
    department: '',
    assetType: '',
    assetName: '',
    quantity: 1,
    priority: 'medium',
    reasonDescription: '',
    requiredDate: new Date().toISOString().split('T')[0],
  });

  const [message, setMessage] = useState<{ type: MessageBarType; text: string } | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  const handleInputChange = (fieldName: string, value: any) => {
    setFormData((prev) => ({
      ...prev,
      [fieldName]: value,
    }));
  };

  const handleSubmit = async () => {
    try {
      if (!formData.employeeEmail || !formData.assetType || !formData.assetName) {
        setMessage({ type: MessageBarType.error, text: 'Please fill in all required fields.' });
        return;
      }

      setIsSubmitting(true);
      const service = new InventoryService(props.apiBaseUrl);
      
      const payload = {
        employeeEmail: formData.employeeEmail,
        assetType: formData.assetType,
        assetName: formData.assetName,
        quantity: formData.quantity,
        priority: formData.priority,
        reasonDescription: formData.reasonDescription,
        requiredDate: formData.requiredDate
      };

      await service.createAssetRequest(payload);
      
      setMessage({ type: MessageBarType.success, text: 'Asset request submitted successfully!' });
      
      // Reset form
      setTimeout(() => {
        setFormData({
          employeeEmail: '',
          employeeName: props.userName || '',
          department: '',
          assetType: '',
          assetName: '',
          quantity: 1,
          priority: 'medium',
          reasonDescription: '',
          requiredDate: new Date().toISOString().split('T')[0],
        });
        setMessage(null);
      }, 2000);
    } catch (error) {
      console.error('Error submitting asset request:', error);
      setMessage({ type: MessageBarType.error, text: 'Failed to submit request. Please try again.' });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleCancel = () => {
    setFormData({
      employeeEmail: '',
      employeeName: props.userName || '',
      department: '',
      assetType: '',
      assetName: '',
      quantity: 1,
      priority: 'medium',
      reasonDescription: '',
      requiredDate: new Date().toISOString().split('T')[0],
    });
    setMessage(null);
  };

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Text variant="xLarge" block style={{ fontWeight: 600 }}>
        Request Asset
      </Text>

      {message && (
        <MessageBar messageBarType={message.type} isMultiline>
          {message.text}
        </MessageBar>
      )}

      <Card>
        <Card.Section tokens={cardTokens}>
          <Stack tokens={{ childrenGap: 15 }}>
            {/* Employee Information */}
            <Stack tokens={{ childrenGap: 10 }}>
              <Text variant="large" style={{ fontWeight: 600, color: '#0078d4' }}>
                Employee Information
              </Text>

              <TextField
                label="Employee Email *"
                placeholder="e.g., john.doe@company.com"
                value={formData.employeeEmail}
                onChange={(ev, newValue) => handleInputChange('employeeEmail', newValue)}
                required
              />

              <TextField
                label="Employee Name"
                value={formData.employeeName}
                onChange={(ev, newValue) => handleInputChange('employeeName', newValue)}
              />

              <TextField
                label="Department"
                placeholder="Enter your department"
                value={formData.department}
                onChange={(ev, newValue) => handleInputChange('department', newValue)}
              />
            </Stack>

            {/* Asset Information */}
            <Stack tokens={{ childrenGap: 10 }}>
              <Text variant="large" style={{ fontWeight: 600, color: '#0078d4' }}>
                Asset Information
              </Text>

              <Dropdown
                label="Asset Type *"
                placeholder="Select asset type"
                options={assetTypes}
                selectedKey={formData.assetType}
                onChange={(ev, option) => handleInputChange('assetType', option?.key)}
                required
              />

              <TextField
                label="Asset Name *"
                placeholder="e.g., HP Pavilion 15"
                value={formData.assetName}
                onChange={(ev, newValue) => handleInputChange('assetName', newValue)}
                required
              />

              <TextField
                label="Quantity"
                type="number"
                value={formData.quantity.toString()}
                onChange={(ev, newValue) => handleInputChange('quantity', parseInt(newValue || '1') || 1)}
              />
            </Stack>

            {/* Request Details */}
            <Stack tokens={{ childrenGap: 10 }}>
              <Text variant="large" style={{ fontWeight: 600, color: '#0078d4' }}>
                Request Details
              </Text>

              <Dropdown
                label="Priority"
                options={priorityOptions}
                selectedKey={formData.priority}
                onChange={(ev, option) => handleInputChange('priority', option?.key)}
              />

              <TextField
                label="Reason / Description"
                multiline
                rows={4}
                placeholder="Explain why you need this asset..."
                value={formData.reasonDescription}
                onChange={(ev, newValue) => handleInputChange('reasonDescription', newValue)}
              />

              <TextField
                label="Required Date"
                type="date"
                value={formData.requiredDate}
                onChange={(ev, newValue) => handleInputChange('requiredDate', newValue)}
              />
            </Stack>

            {/* Action Buttons */}
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton
                text="Submit Request"
                onClick={handleSubmit}
                disabled={isSubmitting}
              />
              <DefaultButton
                text="Cancel"
                onClick={handleCancel}
              />
            </Stack>
          </Stack>
        </Card.Section>
      </Card>
    </Stack>
  );
};

