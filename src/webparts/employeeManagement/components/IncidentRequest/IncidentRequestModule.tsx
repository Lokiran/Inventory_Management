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
import styles from './IncidentRequestModule.module.scss';
import { IEmployeeManagementProps } from '../IEmployeeManagementProps';
import { InventoryService } from '../../services/InventoryService';

interface IIncidentForm {
  employeeEmail: string;
  assetId: string;
  assetName: string;
  issueType: string;
  issueDescription: string;
  priority: string;
  attachmentUrl?: string;
  reportedDate: string;
}

const issueTypes: IDropdownOption[] = [
  { key: 'hardware_damage', text: 'Hardware Damage' },
  { key: 'software_issue', text: 'Software Issue' },
  { key: 'not_working', text: 'Not Working' },
  { key: 'missing_parts', text: 'Missing Parts' },
  { key: 'performance', text: 'Performance Issue' },
  { key: 'other', text: 'Other' },
];

const priorityOptions: IDropdownOption[] = [
  { key: 'low', text: 'Low' },
  { key: 'medium', text: 'Medium' },
  { key: 'high', text: 'High' },
  { key: 'critical', text: 'Critical' },
];

const cardTokens: ICardTokens = { childrenMargin: 12 };

export const IncidentRequestModule: React.FC<IEmployeeManagementProps & { setIsLoading: (loading: boolean) => void }> = (props) => {
  const [formData, setFormData] = useState<IIncidentForm>({
    employeeEmail: '',
    assetId: '',
    assetName: '',
    issueType: '',
    issueDescription: '',
    priority: 'medium',
    reportedDate: new Date().toISOString().split('T')[0],
  });

  const [message, setMessage] = useState<{ type: MessageBarType; text: string } | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);

  const handleInputChange = (fieldName: string, value: any) => {
    setFormData((prev) => ({
      ...prev,
      [fieldName]: value,
    }));
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      setSelectedFile(file);
    }
  };

  const handleSubmit = async () => {
    try {
      if (!formData.employeeEmail || !formData.assetId || !formData.issueType || !formData.issueDescription) {
        setMessage({ type: MessageBarType.error, text: 'Please fill in all required fields.' });
        return;
      }

      setIsSubmitting(true);
      const service = new InventoryService(props.apiBaseUrl);
      
      const payload = {
        ...formData,
        employeeEmail: formData.employeeEmail
      };
      await service.createIncidentRequest(payload, selectedFile || undefined);

      setMessage({ type: MessageBarType.success, text: 'Incident reported successfully! We will investigate shortly.' });

      // Reset form
      setTimeout(() => {
        setFormData({
          employeeEmail: '',
    assetId: '',
          assetName: '',
          issueType: '',
          issueDescription: '',
          priority: 'medium',
          reportedDate: new Date().toISOString().split('T')[0],
        });
        setSelectedFile(null);
        setMessage(null);
      }, 2000);
    } catch (error) {
      console.error('Error submitting incident:', error);
      setMessage({ type: MessageBarType.error, text: 'Failed to report incident. Please try again.' });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleCancel = () => {
    setFormData({
      employeeEmail: '',
    assetId: '',
      assetName: '',
      issueType: '',
      issueDescription: '',
      priority: 'medium',
      reportedDate: new Date().toISOString().split('T')[0],
    });
    setSelectedFile(null);
    setMessage(null);
  };

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Text variant="xLarge" block style={{ fontWeight: 600 }}>
        Raise Incident
      </Text>

      {message && (
        <MessageBar messageBarType={message.type} isMultiline>
          {message.text}
        </MessageBar>
      )}

      <Card>
        <Card.Section tokens={cardTokens}>
          <Stack tokens={{ childrenGap: 15 }}>
            {/* Asset Information */}
            <Stack tokens={{ childrenGap: 10 }}>
              <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c' }}>
                Employee Information
              </Text>
              <TextField
                label="Employee Email *"
                placeholder="e.g., john.doe@company.com"
                value={formData.employeeEmail}
                onChange={(ev, newValue) => handleInputChange('employeeEmail', newValue)}
                required
              />
              <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c', marginTop: '15px' }}>
                Asset Information
              </Text>

              <TextField
                label="Asset ID *"
                placeholder="Enter asset ID"
                value={formData.assetId}
                onChange={(ev, newValue) => handleInputChange('assetId', newValue)}
                required
              />

              <TextField
                label="Asset Name"
                placeholder="Enter asset name"
                value={formData.assetName}
                onChange={(ev, newValue) => handleInputChange('assetName', newValue)}
              />
            </Stack>

            {/* Issue Information */}
            <Stack tokens={{ childrenGap: 10 }}>
              <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c' }}>
                Issue Information
              </Text>

              <Dropdown
                label="Issue Type *"
                placeholder="Select issue type"
                options={issueTypes}
                selectedKey={formData.issueType}
                onChange={(ev, option) => handleInputChange('issueType', option?.key)}
                required
              />

              <TextField
                label="Issue Description *"
                multiline
                rows={5}
                placeholder="Describe the issue in detail..."
                value={formData.issueDescription}
                onChange={(ev, newValue) => handleInputChange('issueDescription', newValue)}
                required
              />

              <Dropdown
                label="Priority"
                options={priorityOptions}
                selectedKey={formData.priority}
                onChange={(ev, option) => handleInputChange('priority', option?.key)}
              />
            </Stack>

            {/* Attachment */}
            <Stack tokens={{ childrenGap: 10 }}>
              <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c' }}>
                Attachment
              </Text>

              <input
                type="file"
                accept="image/*,.pdf,.doc,.docx"
                onChange={handleFileChange}
                style={{
                  padding: '10px',
                  border: '1px solid #ccc',
                  borderRadius: '4px',
                  cursor: 'pointer',
                }}
              />
              {selectedFile && (
                <Text variant="small" style={{ color: '#27ae60' }}>
                  Selected: {selectedFile.name}
                </Text>
              )}
            </Stack>

            {/* Reported Date */}
            <Stack tokens={{ childrenGap: 10 }}>
              <TextField
                label="Reported Date"
                type="date"
                value={formData.reportedDate}
                onChange={(ev, newValue) => handleInputChange('reportedDate', newValue)}
              />
            </Stack>

            {/* Action Buttons */}
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <PrimaryButton
                text="Report Incident"
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

