import * as React from 'react';
import { useState, useEffect } from 'react';
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

interface IAssetRequestModuleProps extends IEmployeeManagementProps {
  employeeId?: string;
  department?: string;
  setIsLoading: (loading: boolean) => void;
}

interface IAssetRequestForm {
  employeeName: string;
  employeeId: string;
  employeeEmail?: string;
  department: string;
  serialNo: string;
  assetName: string;
  assetType: string;
  priority: string;
  reasonDescription: string;
  requiredDate: string;
}

const priorityOptions: IDropdownOption[] = [
  { key: 'low', text: 'Low' },
  { key: 'medium', text: 'Medium' },
  { key: 'high', text: 'High' },
  { key: 'urgent', text: 'Urgent' },
];

const assetNameOptions: IDropdownOption[] = [
  { key: 'HP Pavilion 15', text: 'HP Pavilion 15' },
  { key: 'Dell Latitude 5420', text: 'Dell Latitude 5420' },
  { key: 'MacBook Pro 16', text: 'MacBook Pro 16' },
  { key: 'Lenovo ThinkPad T14', text: 'Lenovo ThinkPad T14' },
  { key: 'iPad Air', text: 'iPad Air' },
  { key: 'iPhone 13', text: 'iPhone 13' },
  { key: 'Logitech MX Master 3', text: 'Logitech MX Master 3' },
  { key: 'Dell UltraSharp 27', text: 'Dell UltraSharp 27' },
];

const assetTypeOptions: IDropdownOption[] = [
  { key: 'Laptop', text: 'Laptop' },
  { key: 'Mobile', text: 'Mobile' },
  { key: 'Tablet', text: 'Tablet' },
  { key: 'Accessory', text: 'Accessory' },
  { key: 'Monitor', text: 'Monitor' },
  { key: 'Other', text: 'Other' },
];

const cardTokens: ICardTokens = { childrenMargin: 12 };

export const AssetRequestModule: React.FC<IAssetRequestModuleProps> = (props) => {
  const [formData, setFormData] = useState<IAssetRequestForm>({
    employeeName: props.userName || '',
    employeeId: props.employeeId || '',
    employeeEmail: props.userEmail || '',
    department: props.department || 'General',
    serialNo: '',
    assetName: '',
    assetType: '',
    priority: 'medium',
    reasonDescription: '',
    requiredDate: new Date().toISOString().split('T')[0],
  });

  const [message, setMessage] = useState<{ type: MessageBarType; text: string } | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  // Sync with props when employee context changes
  useEffect(() => {
    setFormData((prev) => ({
      ...prev,
      employeeName: props.userName || '',
      employeeId: props.employeeId || '',
      department: props.department || 'General',
      employeeEmail: props.userEmail || '',
    }));
  }, [props.userName, props.employeeId, props.department, props.userEmail]);

  // Lookup employee details when typed manually
  useEffect(() => {
    const timer = setTimeout(async () => {
      if (!formData.employeeName.trim()) return;
      try {
        const service = new InventoryService(props.spContext);
        const details = await service.getEmployeeDetailsByName(formData.employeeName);
        setFormData(prev => ({
          ...prev,
          employeeId: details.employeeId || prev.employeeId,
          department: details.department || prev.department,
          employeeEmail: details.email || prev.employeeEmail,
        }));
      } catch (e) {
        console.error('Failed to resolve employee details in asset request', e);
      }
    }, 500);
    return () => clearTimeout(timer);
  }, [formData.employeeName]);

  const handleInputChange = (fieldName: string, value: any) => {
    setFormData((prev) => ({
      ...prev,
      [fieldName]: value,
    }));
  };

  const handleSubmit = async () => {
    try {
      if (!formData.assetName || !formData.assetType || !formData.reasonDescription) {
        setMessage({ type: MessageBarType.error, text: 'Please fill in all required fields.' });
        return;
      }

      setIsSubmitting(true);
      const service = new InventoryService(props.spContext);
      
      const payload = {
        employeeEmail: formData.employeeEmail || props.userEmail,
        employeeName: formData.employeeName,
        employeeId: formData.employeeId,
        department: formData.department,
        serialNo: formData.serialNo || 'N/A',
        assetType: formData.assetType,
        assetName: formData.assetName,
        priority: formData.priority,
        reasonDescription: formData.reasonDescription,
        requiredDate: formData.requiredDate,
      };

      console.log('Submitting asset request to SharePoint', {
        siteUrl: props.webUrl,
        userEmail: payload.employeeEmail,
        payload,
      });

      await service.createAssetRequest(payload);
      
      setMessage({ type: MessageBarType.success, text: 'Asset request submitted successfully!' });
      
      // Reset form fields
      setTimeout(() => {
        setFormData({
          employeeName: props.userName || '',
          employeeId: props.employeeId || '',
          employeeEmail: props.userEmail || '',
          department: props.department || 'General',
          serialNo: '',
          assetName: '',
          assetType: '',
          priority: 'medium',
          reasonDescription: '',
          requiredDate: new Date().toISOString().split('T')[0],
        });
        setMessage(null);
      }, 2000);
    } catch (error: any) {
      console.error('Error submitting asset request:', error);
      const fullError = error && typeof error === 'object' ? JSON.stringify(error, Object.getOwnPropertyNames(error)) : String(error);
      const errorMessage = `Failed to submit request: ${error.message || fullError || 'Check browser console (F12) for details.'}`;
      setMessage({ type: MessageBarType.error, text: errorMessage });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleCancel = () => {
    setFormData({
      employeeName: props.userName || '',
      employeeId: props.employeeId || '',
      employeeEmail: props.userEmail || '',
      department: props.department || 'General',
      serialNo: '',
      assetName: '',
      assetType: '',
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
            <Stack tokens={{ childrenGap: 10 }}>
              <Text variant="large" style={{ fontWeight: 600, color: '#0078d4' }}>
                Employee Information
              </Text>
              <TextField 
                label="Employee Name *" 
                value={formData.employeeName} 
                onChange={(ev, val) => handleInputChange('employeeName', val || '')}
                required
              />
            </Stack>

            <Stack tokens={{ childrenGap: 10 }}>
              <Text variant="large" style={{ fontWeight: 600, color: '#0078d4' }}>
                Asset Details
              </Text>

              <Dropdown
                label="Asset Name *"
                placeholder="Choose an asset model"
                options={assetNameOptions}
                selectedKey={formData.assetName}
                onChange={(ev, option) => handleInputChange('assetName', option?.key)}
                required
              />

              <Dropdown
                label="Asset Type *"
                placeholder="Choose an asset category"
                options={assetTypeOptions}
                selectedKey={formData.assetType}
                onChange={(ev, option) => handleInputChange('assetType', option?.key)}
                required
              />

              <Dropdown
                label="Priority"
                options={priorityOptions}
                selectedKey={formData.priority}
                onChange={(ev, option) => handleInputChange('priority', option?.key)}
              />

              <TextField
                label="Reason for Request *"
                multiline
                rows={4}
                placeholder="Explain why you need this..."
                value={formData.reasonDescription}
                onChange={(ev, newValue) => handleInputChange('reasonDescription', newValue)}
                required
              />
            </Stack>

            <Stack tokens={{ childrenGap: 10 }}>
              <Text variant="large" style={{ fontWeight: 600, color: '#0078d4' }}>
                Timeline
              </Text>

              <TextField
                label="Requested Date"
                type="date"
                value={formData.requiredDate}
                onChange={(ev, newValue) => handleInputChange('requiredDate', newValue)}
              />
            </Stack>

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

